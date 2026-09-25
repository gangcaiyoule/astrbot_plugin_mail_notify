import email as email_lib
import imaplib
from datetime import datetime, timezone
from email.utils import parseaddr, parsedate_to_datetime

from .mail_utils import decode_mime_header, extract_text_body, truncate_text
from .network_utils import open_tcp_socket

# 本模块封装轮询式的 IMAP 同步操作。
# 由于 imaplib 是阻塞式 I/O，调用方需通过 asyncio.to_thread() 在线程池中执行这些函数，
# 避免阻塞 asyncio 事件循环。
# 基于 IDLE 的实时监听另见 imap_idle.py（它会长期占用一个专用线程）。


class _IMAP4PreferredIPv4(imaplib.IMAP4):
    """Plain IMAP4 connection preferring IPv4, optionally through a proxy."""

    def __init__(self, *args, proxy: dict | None = None, **kwargs):
        # Store the proxy before super().__init__() opens the socket, so that
        # _create_socket() can pick it up.
        self._proxy = proxy
        super().__init__(*args, **kwargs)

    def _create_socket(self, timeout):
        return open_tcp_socket(self.host, self.port, timeout=timeout, proxy=self._proxy)


class _IMAP4SSLPreferredIPv4(imaplib.IMAP4_SSL):
    """IMAP4_SSL connection preferring IPv4, optionally through a proxy."""

    def __init__(self, *args, proxy: dict | None = None, **kwargs):
        self._proxy = proxy
        super().__init__(*args, **kwargs)

    def _create_socket(self, timeout):
        sock = open_tcp_socket(self.host, self.port, timeout=timeout, proxy=self._proxy)
        return self.ssl_context.wrap_socket(sock, server_hostname=self.host)


def is_recent_email(mail_info: dict, init_dt: datetime) -> bool:
    """判断邮件是否晚于插件初始化时间，用于过滤首次运行时的历史存量邮件。

    Args:
        mail_info: 由 _parse_email() 返回的邮件信息字典，需包含 'date_raw' 字段（ISO 格式）。
        init_dt:   插件首次运行时记录的初始化时间，作为过滤基准线。

    Returns:
        True  — 邮件时间 >= init_dt，属于新邮件，应当推送。
        False — 邮件时间 < init_dt，属于历史邮件，跳过。
        若解析失败则保守地返回 True，避免漏推。
    """
    date_str = mail_info.get("date_raw", "")
    if not date_str:
        # 无法获取邮件时间时，保守处理：视为新邮件
        return True
    try:
        mail_dt = datetime.fromisoformat(date_str)
        # 若时间戳不带时区信息，统一补充为 UTC，确保跨时区比较正确
        if mail_dt.tzinfo is None:
            mail_dt = mail_dt.replace(tzinfo=timezone.utc)
        if init_dt.tzinfo is None:
            init_dt = init_dt.replace(tzinfo=timezone.utc)
        return mail_dt >= init_dt
    except Exception:
        # 日期解析异常时保守处理：视为新邮件
        return True


def _connect(
    account: dict, proxy: dict | None = None
) -> imaplib.IMAP4 | imaplib.IMAP4_SSL:
    """根据账户配置建立 IMAP 连接并完成认证，返回已选中 INBOX 的连接对象。

    Args:
        account: 包含 imap_server、imap_port、use_ssl、email、password 的账户配置字典。
        proxy:   parse_socks_proxy() 生成的代理描述，None 表示直连。

    Returns:
        已登录并选中 INBOX（只读模式）的 IMAP 连接对象。
    """
    server = account["imap_server"]
    port = account.get("imap_port", 993)  # SSL 端口默认 993
    use_ssl = account.get("use_ssl", True)

    # 根据是否启用 SSL 选择对应的连接类，超时设置为 30 秒
    if use_ssl:
        conn = _IMAP4SSLPreferredIPv4(server, port, timeout=30, proxy=proxy)
    else:
        conn = _IMAP4PreferredIPv4(server, port, timeout=30, proxy=proxy)

    conn.login(account["email"], account["password"])

    # 网易邮箱兼容：发送 IMAP ID 命令，避免 Unsafe Login
    email_addr = account.get("email", "")
    if email_addr.endswith("@163.com") or email_addr.endswith("@126.com"):
        id_payload = (
            '("name" "AstrBot-MailNotify" "version" "1.3.2" "vendor" "AstrBot")'
        )
        try:
            conn.xatom("ID", id_payload)
        except Exception:
            pass  # 非163/不支持ID的服务器可忽略

    # 以只读模式优先选中 INBOX，不会意外修改服务器上的已读状态。
    # 某些邮箱服务商（如部分网易环境）可能对只读模式兼容性较弱，
    # 因此在只读失败时再回退到非只读模式。
    status, data = conn.select("INBOX", readonly=True)
    if status != "OK":
        status, data = conn.select("INBOX", readonly=False)

    if status != "OK":
        raise ValueError(f"无法选择 INBOX: status={status!r}, data={data!r}")
    return conn


def _parse_email(
    msg, uid: bytes, preview_body_len: int, filter_body_len: int | None = None
) -> dict:
    """将原始 email.message.Message 对象解析为结构化字典。

    Args:
        msg:         由 email.message_from_bytes() 解析得到的邮件对象。
        uid:         IMAP UID（bytes），用于去重与增量跟踪。
        preview_body_len: Maximum number of body characters used in notifications.
        filter_body_len:  Maximum number of body characters used for filtering.

    Returns:
        A dictionary containing uid, subject, from_name, from_addr, date,
        date_raw, body, and filter_body.
    """
    # 解码 MIME 编码的主题（可能是 Base64 或 Quoted-Printable 编码的中文）
    subject = decode_mime_header(msg.get("Subject", ""))

    # 解析发件人：拆分显示名与邮件地址，并对显示名做 MIME 解码
    from_raw = msg.get("From", "")
    from_name, from_addr = parseaddr(from_raw)
    from_name = decode_mime_header(from_name) if from_name else from_addr

    # 解析邮件时间：格式化为可读字符串，同时保留 ISO 格式用于时间比较
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        date_formatted = dt.strftime("%Y-%m-%d %H:%M")  # 用于展示
        date_raw = dt.isoformat()  # 用于 is_recent_email() 比较
    except Exception:
        # 日期格式非标准时降级处理，截取原始字符串前 25 位
        date_formatted = date_str[:25] if date_str else "未知"
        date_raw = ""

    # 提取纯文本正文，优先 text/plain，降级到剥离标签的 text/html
    body = extract_text_body(msg)
    preview_body = truncate_text(body, preview_body_len)
    filter_text = truncate_text(body, filter_body_len or preview_body_len)

    return {
        "uid": int(uid),
        "subject": subject or "(无主题)",
        "from_name": from_name,
        "from_addr": from_addr,
        "date": date_formatted,
        "date_raw": date_raw,
        "body": preview_body,
        "filter_body": filter_text,
    }


def imap_fetch_new(
    account: dict,
    last_uid: int,
    preview_body_len: int,
    filter_body_len: int | None = None,
    proxy: dict | None = None,
) -> tuple[list[dict], int]:
    """基于 UID 增量拉取收件箱中的新邮件。同步函数，调用方应通过 asyncio.to_thread() 执行。

    增量机制：
        - last_uid == 0 表示首次运行，此时仅记录当前最大 UID 作为基准，不返回任何邮件，
          防止历史存量邮件在插件初次启动时被批量推送。
        - 后续运行搜索 UID > last_uid 的邮件，实现真正的增量拉取。
        - 每次最多处理最新的 10 封（uid_list[-10:]），避免积压时一次性拉取过多。

    Args:
        account:      账户配置字典（imap_server、email、password 等）。
        last_uid:     上次成功处理的最大邮件 UID，存储于 KV 数据库（data/data_v4.db）。
        preview_body_len: Maximum number of body characters used in notifications.
        filter_body_len:  Maximum number of body characters used for filtering.
        proxy:        代理描述，None 表示直连。

    Returns:
        (new_emails, new_max_uid)：新邮件列表 与 本次处理后的最大 UID。
    """
    conn = _connect(account, proxy)
    new_emails: list[dict] = []
    new_max_uid = last_uid

    try:
        if last_uid == 0:
            # 首次运行：搜索全部邮件，仅记录最大 UID 作为后续增量基准，不推送任何邮件
            status, data = conn.uid("search", "ALL")
            if status == "OK" and data[0]:
                uid_list = data[0].split()
                if uid_list:
                    new_max_uid = int(uid_list[-1])
            return [], new_max_uid

        # 增量搜索：只搜索 UID 大于 last_uid 的邮件
        status, data = conn.uid("search", f"UID {last_uid + 1}:*")
        if status != "OK" or not data[0]:
            return [], last_uid

        uid_list = data[0].split()
        # 服务器可能返回 UID <= last_uid 的结果（部分服务器行为），需二次过滤
        uid_list = [u for u in uid_list if int(u) > last_uid]
        if not uid_list:
            return [], last_uid

        # 记录本次处理到的最大 UID，下次轮询将以此为起点
        new_max_uid = max(int(u) for u in uid_list)

        # 限制每次最多拉取最新的 10 封，防止积压时造成推送风暴
        for uid in uid_list[-10:]:
            status, msg_data = conn.uid("fetch", uid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, bytes):
                continue
            msg = email_lib.message_from_bytes(raw)
            new_emails.append(_parse_email(msg, uid, preview_body_len, filter_body_len))
    finally:
        # 无论成功或异常，确保 IMAP 连接被正常关闭
        try:
            conn.logout()
        except Exception:
            pass

    return new_emails, new_max_uid


def imap_query_since(
    account: dict,
    since_dt: datetime,
    max_body_len: int,
    proxy: dict | None = None,
) -> list[dict]:
    """查询指定日期之后的邮件列表，供 /mail_query 指令使用。同步函数，调用方应通过 asyncio.to_thread() 执行。

    与 imap_fetch_new() 不同，此函数按日期搜索而非 UID，适合用户主动发起的历史查询。
    每次最多返回最新的 20 封，避免结果过长。

    Args:
        account:      账户配置字典。
        since_dt:     查询起始日期（含当天），由 /mail_query 指令传入。
        max_body_len: 正文截取最大字符数。
        proxy:        代理描述，None 表示直连。

    Returns:
        满足条件的邮件字典列表，按 UID 升序排列（最多 20 封）。
    """
    conn = _connect(account, proxy)
    results: list[dict] = []

    try:
        # IMAP SINCE 搜索条件要求日期格式为 DD-Mon-YYYY，如 01-Mar-2026
        since_str = since_dt.strftime("%d-%b-%Y")
        status, data = conn.uid("search", f"SINCE {since_str}")
        if status != "OK" or not data[0]:
            return []

        uid_list = data[0].split()
        # 限制最多返回最新 20 封，防止结果过多导致消息过长
        for uid in uid_list[-20:]:
            status, msg_data = conn.uid("fetch", uid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, bytes):
                continue
            msg = email_lib.message_from_bytes(raw)
            results.append(_parse_email(msg, uid, max_body_len))
    finally:
        # 无论成功或异常，确保 IMAP 连接被正常关闭
        try:
            conn.logout()
        except Exception:
            pass

    return results
