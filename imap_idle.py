import imaplib
import threading
import time
from collections.abc import Callable

from .imap_client import _connect

# 本模块实现基于 IMAP IDLE（RFC 2177）的长连接实时监听，与 imap_client.py 中的
# 轮询式操作分离：IDLE 需要一个长期占用、手工接管 socket 的阻塞线程，
# 而轮询操作是短平快的请求-响应，两者生命周期与错误处理方式完全不同。

# RFC 2177 §3 要求客户端不要在 IDLE 中停留超过 30 分钟，这里留出余量。
_IDLE_REFRESH_SECONDS = 25 * 60
# IDLE 期间 socket 读取的超时粒度。用它代替 select()：既能周期性检查停止标志
# 与保活截止时间，又不会被 SSL 层残留数据误导（select() 可能在应用层没有完整
# 响应时报告可读，从而让一次读取无限阻塞）。
_IDLE_POLL_SECONDS = 1.0
# 单条响应行的长度上限，防止异常服务器无限堆积缓冲区。
_IDLE_MAX_LINE = 1024 * 1024


class IdleUnsupportedError(Exception):
    """服务器未声明 IDLE 能力，调用方应回退为轮询。"""


def idle_watch(
    account: dict,
    proxy: dict | None,
    stop_event: threading.Event,
    on_new_mail: Callable[[], object],
    idle_timeout: float = _IDLE_REFRESH_SECONDS,
) -> None:
    """用 IMAP IDLE 长连接实时监听收件箱，有变化时回调 on_new_mail()。

    本函数是阻塞的，必须运行在独立线程中：IDLE 会话会长时间占用线程，放进
    asyncio 的默认线程池会挤占其它任务。连接中断时向上抛出异常，由调用方
    决定何时重连。同步函数，调用方应将其放在专用 threading.Thread 中执行。

    实现上不使用 imaplib 的 _command()：它无法解析 "+" 续行，且 Gmail 在 IDLE
    期间推送的 FETCH 未标记响应会让 _get_tagged_response() 抛出 KeyError，
    因此这里手拼裸命令并自行读取原始响应行。

    Args:
        account:      账户配置字典（imap_server、email、password 等）。
        proxy:        代理描述，None 表示直连。
        stop_event:   置位后本函数会尽快返回，用于插件卸载时结束监听。
        on_new_mail:  检测到新邮件的回调，在本线程中直接调用。
        idle_timeout: 单次 IDLE 最长持续秒数，到期后重发 IDLE 保活。

    Raises:
        IdleUnsupportedError: 服务器未声明 IDLE 能力，调用方应回退为轮询。
        imaplib.IMAP4.abort:  连接被服务器关闭或 IDLE 异常结束。
    """
    conn = _connect(account, proxy)
    try:
        # imaplib 在 3.14 之前只在登录前读取一次 CAPABILITY（_get_capabilities
        # 只在 _connect 中被调用），而 Gmail 等服务器仅在登录后才声明 IDLE，
        # 因此这里主动刷新一次能力列表。
        status, data = conn.capability()
        caps = data[-1] if status == "OK" and data else b""
        if isinstance(caps, bytes):
            caps = caps.decode("ascii", "ignore")
        conn.capabilities = tuple(str(caps).upper().split())
        if "IDLE" not in conn.capabilities:
            raise IdleUnsupportedError(f"{account.get('imap_server')} 未声明 IDLE 能力")

        # 从这里开始自己接管 socket 读取，不再调用任何 imaplib 读操作。
        # 原因：imaplib 3.12 通过 BufferedReader + SocketIO 读取，而 SocketIO
        # 一旦遇到 socket 超时就永久置位错误状态，之后所有读取都抛
        # "cannot read from timed out object"；可长时间的 IDLE 等待恰恰依赖
        # 反复超时（CPython 直到 3.14 才为此自建缓冲读取，issues/51571）。
        # 原始 socket 的 recv() 超时是可恢复的。
        # 先用极短超时把 imaplib 已经预读到缓冲区里的字节取出来，避免丢响应；
        # 这一步可能让 conn.file 进入上述失效状态，但此后不再使用它。
        buffered = bytearray()
        conn.sock.settimeout(0.2)
        peek = getattr(conn.file, "peek", None)
        if peek is not None:
            # conn.file 运行时是 BufferedReader，peek() 能把 imaplib 已经
            # 预读走的字节交还回来；超时（TimeoutError 是 OSError 子类）
            # 只说明缓冲区为空，不影响后续从 socket 读取。
            try:
                buffered += peek(1)
            except OSError:
                pass
        conn.sock.settimeout(_IDLE_POLL_SECONDS)

        def next_line() -> bytes:
            """读取一行响应（不含 CRLF），无数据时返回空字节串。

            Returns:
                下一行响应的内容；socket 超时且没有完整行时返回 b""。

            Raises:
                imaplib.IMAP4.abort: 连接被关闭，或响应行超过长度上限。
            """
            while True:
                index = buffered.find(b"\r\n")
                if index >= 0:
                    line = bytes(buffered[:index])
                    del buffered[: index + 2]
                    return line
                if len(buffered) > _IDLE_MAX_LINE:
                    raise imaplib.IMAP4.abort("IDLE 响应行过长")
                chunk = conn.sock.recv(4096)
                if not chunk:
                    raise imaplib.IMAP4.abort("IDLE 连接被服务器关闭")
                # 用 extend() 而不是 +=，避免在闭包里重新绑定 buffered
                buffered.extend(chunk)

        while not stop_event.is_set():
            # IDLE 是手拼的裸命令：_command() 读不懂 "+" 续行，
            # 因此自己发标签，并在结束后自行清理 tagged_commands
            # （_new_tag() 会顺手注册 tag，不清理会残留到会话结束）。
            tag = conn._new_tag()
            got_mail = False
            try:
                conn.send(tag + b" IDLE\r\n")
                deadline = time.monotonic() + idle_timeout
                idling = False

                # 阶段一：等待 "* n EXISTS" 之类的主动推送
                while not stop_event.is_set() and time.monotonic() < deadline:
                    try:
                        line = next_line()
                    except TimeoutError:
                        # 空闲，顺带检查停止标志与保活截止时间
                        continue
                    if not line:
                        continue  # 服务器偶尔会插入空行
                    if line.startswith(tag + b" "):
                        # 还没收到 "+" 续行就被 tagged 响应终止，说明 IDLE 被拒
                        raise imaplib.IMAP4.abort(f"IDLE 被服务器拒绝: {line!r}")
                    if line.startswith(b"+"):
                        # 服务器确认进入 IDLE。此前到达的未标记响应（例如命令
                        # 发出前就已在途的 EXISTS）不算推送，需要清掉。
                        idling = True
                        got_mail = False
                    elif idling and b" EXISTS" in line:
                        got_mail = True
                    if got_mail:
                        break

                # 阶段二：发 DONE 收尾，并把到达 DONE 之前的残留响应读到
                # tagged 完成行为止，否则它们会被下一轮 IDLE 误判。
                try:
                    conn.send(b"DONE\r\n")
                    drain_deadline = time.monotonic() + 30
                    while not stop_event.is_set() and time.monotonic() < drain_deadline:
                        try:
                            line = next_line()
                        except TimeoutError:
                            continue
                        if line.startswith(tag + b" "):
                            break
                except (OSError, imaplib.IMAP4.abort):
                    # 连接在推送之后、DONE 收尾之前断开：推送本身已经拿到，
                    # 不能因为它丢掉通知。没有新邮件时按正常故障上报。
                    if not got_mail:
                        raise
            finally:
                conn.tagged_commands.pop(tag, None)

            if got_mail:
                on_new_mail()
    finally:
        # 接管 socket 之后 conn.file 已与 socket 不同步，不能再调用
        # conn.logout()（它会从旧缓冲区读取）。这里直接关闭连接。
        try:
            conn.shutdown()
        except Exception:
            pass
