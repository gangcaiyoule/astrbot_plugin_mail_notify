import asyncio
import re
import threading
from datetime import datetime, timezone

from astrbot.api import AstrBotConfig, logger
from astrbot.api.event import AstrMessageEvent, MessageChain, filter
from astrbot.api.star import Context, Star, register

from .imap_client import imap_fetch_new, imap_query_since, is_recent_email
from .imap_idle import IdleUnsupportedError, idle_watch
from .network_utils import parse_socks_proxy
from .smtp_client import smtp_send_mail
from .timezone_labels import (
    format_mail_display_time,
    format_runtime_display_time,
    get_display_timezone_name,
)


@register(
    "astrbot_plugin_mail_notify",
    "YourName",
    "监控邮箱新邮件并通过 QQ 私聊发送通知",
    "1.3.6",
)
class MailNotifyPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config
        # Background polling task created during initialize().
        self._check_task: asyncio.Task | None = None
        # Consumer for IDLE wake-ups, created during initialize().
        self._idle_task: asyncio.Task | None = None
        # IDLE threads notify the event loop through this queue.
        self._idle_queue: asyncio.Queue = asyncio.Queue()
        # Event loop handle used to hop back from IDLE threads; set in initialize().
        self._loop: asyncio.AbstractEventLoop
        # One dedicated thread per IDLE-capable account: a long-lived IDLE session
        # would otherwise occupy a worker of the shared asyncio thread pool.
        self._idle_threads: dict[str, threading.Thread] = {}
        self._idle_stop = threading.Event()
        # Accounts that reported "no IDLE capability", so they fall back to polling.
        self._idle_unsupported: set[str] = set()
        # Guards against an IDLE wake-up and a safety poll checking one account twice.
        self._check_locks: dict[str, asyncio.Lock] = {}
        # Runtime-only status used by /mail_status; not persisted.
        self._last_check_time: dict[str, str] = {}
        self._account_status: dict[str, str] = {}

    def _get_display_timezone(self) -> str:
        configured_name = self.config.get("display_timezone", "")
        if configured_name == "__custom__":
            configured_name = self.config.get("display_timezone_custom", "")
        return get_display_timezone_name(configured_name)

    def _format_mail_time(self, mail_info: dict) -> str:
        return format_mail_display_time(
            mail_info.get("date_raw", ""),
            mail_info.get("date", "未知"),
            self._get_display_timezone(),
        )

    def _get_imap_proxy(self) -> dict | None:
        """Build the IMAP proxy descriptor from the global proxy settings.

        Returns:
            Proxy descriptor for imap_client, or None when the proxy is off.

        Raises:
            ValueError: The proxy is enabled but its settings are invalid.
        """
        return parse_socks_proxy(self.config.get("proxy_settings", {}))

    def _format_runtime_time(self, date_raw: str) -> str:
        return format_runtime_display_time(date_raw, self._get_display_timezone())

    async def initialize(self):
        """插件初始化后启动后台轮询循环与 IDLE 事件消费协程"""
        # IDLE 线程需要事件循环句柄才能用 call_soon_threadsafe() 回调回来。
        self._loop = asyncio.get_running_loop()
        self._check_task = asyncio.create_task(self._check_loop())
        self._idle_task = asyncio.create_task(self._idle_consumer())
        logger.info("邮件通知插件：后台检查循环已启动。")

    # ── 后台循环 ──────────────────────────────────────────

    def _get_idle_settings(self) -> dict:
        """读取实时监听（IDLE）相关配置。

        Returns:
            配置字典；缺失或类型异常时返回空字典。
        """
        settings = self.config.get("idle_settings", {})
        return settings if isinstance(settings, dict) else {}

    def _is_check_due(self, account_email: str, interval_minutes: int) -> bool:
        """判断某账户是否到达下一次检查时间。

        Args:
            account_email: 账户地址，同时作为运行态时间戳的键。
            interval_minutes: 两次检查之间的最短分钟数。

        Returns:
            从未检查过或已超过间隔时返回 True。
        """
        last = self._last_check_time.get(account_email)
        if not last:
            return True
        try:
            last_dt = datetime.fromisoformat(last)
        except ValueError:
            return True
        if last_dt.tzinfo is None:
            last_dt = last_dt.replace(tzinfo=timezone.utc)
        elapsed = (datetime.now(timezone.utc) - last_dt).total_seconds()
        return elapsed >= interval_minutes * 60

    def _sync_idle_workers(self):
        """为每个支持 IDLE 的账户启动一个专职监听线程（幂等）。"""
        if not self._get_idle_settings().get("enable", True):
            return
        for account in self.config.get("mail_accounts", []):
            email_addr = (account.get("email") or "").strip()
            if not email_addr or not account.get("imap_server"):
                continue
            if email_addr in self._idle_unsupported:
                continue
            worker = self._idle_threads.get(email_addr)
            if worker and worker.is_alive():
                continue
            worker = threading.Thread(
                target=self._idle_worker,
                # 传副本，避免运行中改配置时线程读到半新半旧的数据。
                args=(dict(account), email_addr),
                name=f"mail-notify-idle-{email_addr}",
                daemon=True,
            )
            self._idle_threads[email_addr] = worker
            worker.start()

    def _idle_worker(self, account: dict, account_email: str):
        """维持 IMAP IDLE 长连接，并在有新邮件时唤醒事件循环。

        运行在独立线程中：阻塞的 IMAP 读写在这里发生，而真正的邮件检查会被
        调度回 asyncio 事件循环，从而完整复用原有的通知、去重与过滤逻辑。

        Args:
            account: 本监听器使用的账户配置快照。
            account_email: 账户地址，同时作为队列消息内容。
        """
        while not self._idle_stop.is_set():
            try:
                refresh_seconds = (
                    max(
                        int(self._get_idle_settings().get("refresh_minutes", 25) or 25),
                        1,
                    )
                    * 60
                )
                idle_watch(
                    account,
                    self._get_imap_proxy(),
                    self._idle_stop,
                    # asyncio.Queue 不是线程安全的，必须回到事件循环中入队。
                    lambda: self._loop.call_soon_threadsafe(
                        self._idle_queue.put_nowait, account_email
                    ),
                    idle_timeout=refresh_seconds,
                )
                return  # idle_watch 正常返回，说明收到了停止信号。
            except IdleUnsupportedError as e:
                self._idle_unsupported.add(account_email)
                logger.info(
                    f"邮件通知插件：{account_email} 不支持 IDLE，改为轮询（{e}）"
                )
                return
            except Exception as e:
                logger.error(
                    f"邮件通知插件：{account_email} IDLE 监听中断，即将重连: {e}"
                )
                if self._idle_stop.wait(30):
                    return

    async def _idle_consumer(self):
        """消费 IDLE 推送事件，在事件循环中执行邮件检查。"""
        while True:
            account_email = await self._idle_queue.get()
            # 合并突发推送：连续到达的邮件只需检查一次，UID 增量会一并取回。
            while not self._idle_queue.empty():
                self._idle_queue.get_nowait()

            notify_umo = self.config.get("notify_umo", "")
            if not notify_umo:
                continue
            account = next(
                (
                    acc
                    for acc in self.config.get("mail_accounts", [])
                    if (acc.get("email") or "").strip() == account_email
                ),
                None,
            )
            if not account or not account.get("imap_server"):
                continue
            try:
                await self._check_account(account, notify_umo)
                self._account_status[account_email] = "✅ 正常"
            except Exception as e:
                self._account_status[account_email] = f"❌ {str(e)[:80]}"
                logger.error(f"邮件通知插件：{account_email} 检查失败: {e}")
            self._last_check_time[account_email] = datetime.now(
                timezone.utc
            ).isoformat()

    async def _check_loop(self):
        """按账户间隔执行检查，IDLE 账户只做低频兜底轮询。

        长连接一旦静默失效（被中间设备掐断而不发 FIN），没有兜底就会一直漏
        邮件，因此保留一个远大于 check_interval 的兜底周期。
        """
        await asyncio.sleep(10)  # 等待系统初始化完成
        while True:
            try:
                notify_umo = self.config.get("notify_umo", "")
                # 只有绑定了通知目标才检查
                if notify_umo:
                    self._sync_idle_workers()
                    idle_enabled = bool(self._get_idle_settings().get("enable", True))
                    safety_minutes = max(
                        int(
                            self._get_idle_settings().get("safety_poll_interval", 15)
                            or 15
                        ),
                        1,
                    )
                    for account in self.config.get("mail_accounts", []):
                        email_addr = (account.get("email") or "").strip()
                        if not email_addr or not account.get("imap_server"):
                            continue
                        watching = (
                            idle_enabled
                            and email_addr in self._idle_threads
                            and email_addr not in self._idle_unsupported
                        )
                        interval = (
                            safety_minutes
                            if watching
                            else max(
                                int(self.config.get("check_interval", 5) or 5),
                                1,
                            )
                        )
                        # 外层每 60 秒醒一次，各账户按自己的间隔决定是否真的检查。
                        if not self._is_check_due(email_addr, interval):
                            continue
                        try:
                            await self._check_account(account, notify_umo)
                            self._account_status[email_addr] = "✅ 正常"
                        except Exception as e:
                            self._account_status[email_addr] = f"❌ {str(e)[:80]}"
                            logger.error(f"邮件通知插件：{email_addr} 检查失败: {e}")
                        self._last_check_time[email_addr] = datetime.now(
                            timezone.utc
                        ).isoformat()
                await asyncio.sleep(60)
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"邮件通知插件：循环异常: {e}")
                await asyncio.sleep(60)

    # ── IMAP逻辑 ───────────────────────────────────────────────

    async def _check_account(self, account: dict, notify_umo: str):
        # 每个邮箱独立存储状态，避免冲突
        account_email = account["email"]
        # IDLE 推送与兜底轮询可能同时触发同一账户。重叠时直接跳过这一次：
        # 正在运行的那次会用 UID 增量把新邮件一并取回，既不会漏也不会重复推送。
        lock = self._check_locks.setdefault(account_email, asyncio.Lock())
        if lock.locked():
            return
        async with lock:
            max_body_len = max(int(self.config.get("max_body_length", 500) or 500), 1)
            filter_body_len = max(
                int(self.config.get("filter_body_length", 3000) or 3000),
                max_body_len,
            )

            uid_key = f"last_uid_{account_email}"
            init_key = f"init_time_{account_email}"
            last_uid = await self.get_kv_data(uid_key, 0) or 0
            init_time = await self.get_kv_data(init_key, "")

            is_first_run = not init_time
            if is_first_run:
                # 首次运行记录初始化时间和当前UID基线，防止历史邮件被推送
                init_time = datetime.now(timezone.utc).isoformat()
                await self.put_kv_data(init_key, init_time)

            # imaplib为阻塞操作，实际查询在工作线程中执行
            new_emails, new_max_uid = await asyncio.to_thread(
                imap_fetch_new,
                account,
                last_uid,
                max_body_len,
                filter_body_len,
                proxy=self._get_imap_proxy(),
            )

            if new_max_uid > last_uid:
                await self.put_kv_data(uid_key, new_max_uid)

            if is_first_run:
                if new_max_uid > 0:
                    logger.info(
                        f"邮件通知插件：{account_email} 初始化完成，最大UID = {new_max_uid}"
                    )
                return

            init_dt = datetime.fromisoformat(init_time)
            for mail_info in new_emails:
                # 二次校验邮件时间，避免刚拉取的邮件属于历史存量
                if is_recent_email(mail_info, init_dt):
                    should_notify, reason = self._should_notify_mail(mail_info)
                    log_mail_from = (
                        mail_info.get("from_addr") or mail_info.get("from_name") or "?"
                    )
                    log_mail_subject = (mail_info.get("subject") or "")[:120]
                    if not should_notify:
                        logger.info(
                            "MailNotify: filtered mail for account=%s, reason=%s, from=%s, subject=%s",
                            account_email,
                            reason,
                            log_mail_from,
                            log_mail_subject,
                        )
                        continue
                    logger.info(
                        "MailNotify: notification allowed for account=%s, reason=%s, from=%s, subject=%s",
                        account_email,
                        reason,
                        log_mail_from,
                        log_mail_subject,
                    )
                    await self._send_notification(account, mail_info, notify_umo)

    # ── Notification ─────────────────────────────────────────────

    def _get_filter_settings(self, prefix: str) -> dict:
        settings = self.config.get(f"{prefix}_settings", {}) or {}
        if isinstance(settings, dict):
            return settings
        return {}

    def _get_filter_enabled(self, prefix: str) -> bool:
        settings = self._get_filter_settings(prefix)
        if "enable" in settings:
            return bool(settings.get("enable", False))
        return bool(self.config.get(f"enable_{prefix}", False))

    def _get_filter_rules(self, prefix: str, field: str) -> list[str]:
        settings = self._get_filter_settings(prefix)
        nested_values = settings.get(f"{field}_rules", [])
        if not nested_values:
            nested_values = self.config.get(f"{field}_{prefix}", []) or []

        values = nested_values or []
        return [
            str(value).strip()
            for value in values
            if isinstance(value, (str, int, float)) and str(value).strip()
        ]

    @staticmethod
    def _matches_sender_rule(mail_info: dict, rule: str) -> bool:
        normalized_rule = rule.strip().casefold()
        if not normalized_rule:
            return False

        from_addr = (mail_info.get("from_addr") or "").strip().casefold()
        from_name = (mail_info.get("from_name") or "").strip().casefold()

        if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", normalized_rule):
            return from_addr == normalized_rule
        if normalized_rule.startswith("@"):
            return from_addr.endswith(normalized_rule)
        return normalized_rule in from_addr or normalized_rule in from_name

    @staticmethod
    def _matches_contains_rule(text: str, rule: str) -> bool:
        normalized_rule = rule.strip().casefold()
        if not normalized_rule:
            return False
        return normalized_rule in (text or "").casefold()

    def _match_rule_group(
        self, mail_info: dict, prefix: str
    ) -> tuple[bool, str | None]:
        sender_rules = self._get_filter_rules(prefix, "sender")
        for rule in sender_rules:
            if self._matches_sender_rule(mail_info, rule):
                return True, f"sender:{rule}"

        subject_rules = self._get_filter_rules(prefix, "subject")
        subject = mail_info.get("subject") or ""
        for rule in subject_rules:
            if self._matches_contains_rule(subject, rule):
                return True, f"subject:{rule}"

        body_rules = self._get_filter_rules(prefix, "body")
        filter_body = mail_info.get("filter_body") or mail_info.get("body") or ""
        for rule in body_rules:
            if self._matches_contains_rule(filter_body, rule):
                return True, f"body:{rule}"

        return False, None

    def _should_notify_mail(self, mail_info: dict) -> tuple[bool, str]:
        enable_blacklist = self._get_filter_enabled("blacklist")
        enable_whitelist = self._get_filter_enabled("whitelist")

        if enable_blacklist:
            is_blacklisted, rule = self._match_rule_group(mail_info, "blacklist")
            if is_blacklisted:
                return False, f"被黑名单屏蔽 ({rule})"

        if enable_whitelist:
            is_whitelisted, rule = self._match_rule_group(mail_info, "whitelist")
            if not is_whitelisted:
                return False, "被白名单屏蔽"
            return True, f"白名单允许 ({rule})"

        return True, "允许，因为不需要匹配白名单限制"

    async def _send_notification(self, account: dict, mail_info: dict, notify_umo: str):
        account_name = account.get("name") or account["email"]
        use_ai = self.config.get("ai_summary", False)
        body_text = mail_info["body"]
        display_time = self._format_mail_time(mail_info)

        if use_ai and body_text:
            body_text = await self._try_ai_summary(mail_info, notify_umo, body_text)

        lines = [
            f"📬 新邮件通知 [{account_name}]",
            "━━━━━━━━━━━━━━━━",
            f"📤 发件人: {mail_info['from_name']}",
        ]
        if mail_info["from_addr"] and mail_info["from_addr"] != mail_info["from_name"]:
            lines[-1] += f" <{mail_info['from_addr']}>"
        lines.append(f"📋 主题: {mail_info['subject']}")
        lines.append(f"🕐 时间: {display_time}")
        if body_text:
            label = "📝 AI摘要" if use_ai else "📝 预览"
            lines.append(f"{label}: {body_text}")

        chain = MessageChain().message("\n".join(lines))
        await self.context.send_message(notify_umo, chain)

    async def _try_ai_summary(
        self, mail_info: dict, notify_umo: str, fallback: str
    ) -> str:
        try:
            provider_id = await self.context.get_current_chat_provider_id(
                umo=notify_umo
            )
            if not provider_id:
                return fallback
            prompt = (
                "请用简洁的中文（不超过100字）总结以下邮件内容，只输出摘要：\n"
                f"主题：{mail_info['subject']}\n"
                f"正文：{fallback}"
            )
            llm_resp = await self.context.llm_generate(
                chat_provider_id=provider_id,
                prompt=prompt,
            )
            if llm_resp and llm_resp.completion_text:
                return llm_resp.completion_text
        except Exception as e:
            logger.warning(f"MailNotify: AI summary failed: {e}")
        return fallback

    def _get_account_by_name_or_email(self, account_name: str) -> dict | None:
        accounts = self.config.get("mail_accounts", [])
        target_name = account_name.strip()
        for acc in accounts:
            name = (acc.get("name") or "").strip()
            addr = (acc.get("email") or "").strip()
            if target_name in (name, addr):
                return acc
        return None

    def _get_admin_uids(self) -> set[str]:
        admin_uids = self.config.get("admin_uids", []) or []
        return {
            str(uid).strip()
            for uid in admin_uids
            if isinstance(uid, (str, int)) and str(uid).strip()
        }

    def _get_admin_denied_message(self) -> str:
        if not self._get_admin_uids():
            return "❌ 还未指定插件管理员。\n请在插件web设置的admin_uid中添加用户id。"
        return "❌ 无权限使用该命令。"

    def _is_plugin_admin(self, event: AstrMessageEvent) -> bool:
        admin_uids = self._get_admin_uids()
        sender_id = str(event.get_sender_id()).strip()
        return bool(sender_id and sender_id in admin_uids)

    def _parse_mail_reply_args(self, message_str: str) -> tuple[str, str, str, str]:
        raw = re.sub(r"\s+", " ", (message_str or "").strip())
        if not raw:
            raise ValueError("参数为空。")

        parts = raw.split(" ", 1)
        if len(parts) < 2:
            raise ValueError("参数缺失。")
        args_text = parts[1].strip()

        args = args_text.split(" ", 2)
        if len(args) < 3:
            raise ValueError("参数不足。")

        account_name, to_addr, subject_body = args[0].strip(), args[1].strip(), args[2]

        if "|" not in subject_body:
            raise ValueError("缺少主题与正文分隔符。")

        subject, body = [s.strip() for s in subject_body.split("|", 1)]
        if not account_name:
            raise ValueError("账户名不能为空。")
        if not to_addr:
            raise ValueError("收件人不能为空。")
        if "@" not in to_addr:
            raise ValueError("收件人邮箱格式错误。")
        if not subject:
            raise ValueError("邮件主题不能为空。")
        if not body:
            raise ValueError("邮件正文不能为空。")
        if len(subject) > 200:
            raise ValueError("邮件主题过长（最多 200 字符）。")
        if len(body) > 5000:
            raise ValueError("邮件正文过长（最多 5000 字符）。")
        return account_name, to_addr, subject, body

    # ── Commands ─────────────────────────────────────────────────

    @filter.command("mail_bind")
    async def mail_bind(self, event: AstrMessageEvent):
        if not self._is_plugin_admin(event):
            yield event.plain_result(self._get_admin_denied_message())
            return
        """绑定当前会话为邮件通知目标"""
        umo = event.unified_msg_origin
        # This is regular plugin config, not KV state, so it is saved in config files.
        self.config["notify_umo"] = umo
        self.config.save_config()
        yield event.plain_result(f"✅ 已绑定当前会话为邮件通知目标。\n会话 ID: {umo}")

    @filter.command("mail_status")
    async def mail_status(self, event: AstrMessageEvent):
        if not self._is_plugin_admin(event):
            yield event.plain_result(self._get_admin_denied_message())
            return
        """查看所有邮箱的监控状态"""
        # Read current config plus runtime cache to render a status snapshot.
        accounts = self.config.get("mail_accounts", [])
        notify_umo = self.config.get("notify_umo", "")
        interval = self.config.get("check_interval", 5)

        if not accounts:
            yield event.plain_result(
                "📭 未配置任何邮箱账户，请在 WebUI 插件配置中添加。"
            )
            return

        lines = [
            f"📊 邮箱监控状态 (间隔: {interval}分钟)",
            f"🔔 通知目标: {'已绑定' if notify_umo else '❗未绑定，请先 /mail_bind'}",
            "━━━━━━━━━━━━━━━━",
        ]
        for acc in accounts:
            addr = acc.get("email", "?")
            name = acc.get("name") or addr
            status = self._account_status.get(addr, "⏳ 等待首次检查")
            last = self._format_runtime_time(self._last_check_time.get(addr, ""))
            lines.append(f"📧 {name} ({addr})")
            lines.append(f"   状态: {status}")
            watching = addr in self._idle_threads and addr not in self._idle_unsupported
            lines.append(f"   模式: {'⚡ 实时监听(IDLE)' if watching else '🕒 轮询'}")
            lines.append(f"   最近检查: {last}")

        yield event.plain_result("\n".join(lines))

    @filter.command("mail_check")
    async def mail_check(self, event: AstrMessageEvent):
        if not self._is_plugin_admin(event):
            yield event.plain_result(self._get_admin_denied_message())
            return
        """立即手动检查所有邮箱"""
        accounts = self.config.get("mail_accounts", [])
        if not accounts:
            yield event.plain_result(
                "📭 未配置任何邮箱账户，请在 WebUI 插件配置中添加。"
            )
            return

        notify_umo = self.config.get("notify_umo", "")
        if not notify_umo:
            yield event.plain_result(
                "❌ Notification target is not bound yet. Please run /mail_bind first."
            )
            return
        yield event.plain_result("🔍 正在检查所有邮箱...")

        # Manual check reuses the same account-checking path as the background loop.
        errors = []
        for account in accounts:
            if not account.get("email") or not account.get("imap_server"):
                continue
            email_addr = account["email"]
            try:
                await self._check_account(account, notify_umo)
                self._account_status[email_addr] = "✅ 正常"
            except Exception as e:
                self._account_status[email_addr] = f"❌ {str(e)[:80]}"
                errors.append(f"{account.get('name') or email_addr}: {e}")
            self._last_check_time[email_addr] = datetime.now(timezone.utc).isoformat()

        if errors:
            yield event.plain_result("⚠️ 部分邮箱检查失败:\n" + "\n".join(errors))
        else:
            yield event.plain_result("✅ 所有邮箱检查完成。")

    @filter.command("mail_query")
    async def mail_query(
        self, event: AstrMessageEvent, account_name: str, since_date: str
    ):
        if not self._is_plugin_admin(event):
            yield event.plain_result(self._get_admin_denied_message())
            return
        """查询指定邮箱自某日期以来的邮件，如 /mail_query qq邮箱 2026-03-01"""
        accounts = self.config.get("mail_accounts", [])

        # Resolve the target account by either display name or full email address.
        target = None
        for acc in accounts:
            name = acc.get("name", "")
            addr = acc.get("email", "")
            if account_name in (name, addr):
                target = acc
                break
        if not target:
            yield event.plain_result(
                f'❌ 未找到名为 "{account_name}" 的邮箱账户。\n'
                f"已配置的账户: {', '.join(a.get('name') or a.get('email', '?') for a in accounts)}"
            )
            return

        # The command accepts only YYYY-MM-DD to keep parsing deterministic.
        try:
            since_dt = datetime.strptime(since_date, "%Y-%m-%d")
        except ValueError:
            yield event.plain_result(
                "❌ 日期格式错误，请使用 YYYY-MM-DD，如 2026-03-01"
            )
            return

        yield event.plain_result(
            f"🔍 正在查询 {account_name} 自 {since_date} 以来的邮件..."
        )

        try:
            max_body_len = self.config.get("max_body_length", 500)
            # History query also uses a worker thread because IMAP access is blocking.
            emails = await asyncio.to_thread(
                imap_query_since,
                target,
                since_dt,
                max_body_len,
                proxy=self._get_imap_proxy(),
            )
        except Exception as e:
            yield event.plain_result(f"❌ 查询失败: {e}")
            return

        if not emails:
            yield event.plain_result(
                f"📭 {account_name} 自 {since_date} 以来没有邮件。"
            )
            return

        lines = [
            f"📬 {account_name} 自 {since_date} 以来共 {len(emails)} 封邮件：",
            "━━━━━━━━━━━━━━━━",
        ]
        for i, m in enumerate(emails, 1):
            lines.append(f"{i}. 📋 {m['subject']}")
            lines.append(f"   📤 {m['from_name']}  🕐 {self._format_mail_time(m)}")
        yield event.plain_result("\n".join(lines))

    @filter.command("mail_reply")
    async def mail_reply(self, event: AstrMessageEvent):
        if not self._is_plugin_admin(event):
            yield event.plain_result(self._get_admin_denied_message())
            return
        """手动发送邮件回复。格式：/mail_reply <账户备注名> <收件人邮箱> <主题>|<正文>"""
        usage = (
            "❌ 用法错误\n"
            "格式: /mail_reply <账户备注名> <收件人邮箱> <主题>|<正文>\n"
            "示例: /mail_reply qq邮箱 test@example.com 回复主题|你好，已收到你的邮件。"
        )

        try:
            account_name, to_addr, subject, body = self._parse_mail_reply_args(
                event.message_str
            )
        except ValueError as e:
            yield event.plain_result(f"{usage}\n原因: {e}")
            return

        account = self._get_account_by_name_or_email(account_name)
        if not account:
            accounts = self.config.get("mail_accounts", [])
            account_names = ", ".join(
                (a.get("name") or a.get("email") or "?") for a in accounts
            )
            yield event.plain_result(
                f'❌ 未找到名为 "{account_name}" 的邮箱账户。\n已配置账户: {account_names or "(空)"}'
            )
            return

        if not account.get("smtp_server"):
            yield event.plain_result(
                "❌ 该账户未配置 SMTP 服务器。请在插件配置中填写 smtp_server、smtp_port、smtp_use_ssl。"
            )
            return

        yield event.plain_result("📤 正在发送邮件...")
        try:
            await asyncio.to_thread(smtp_send_mail, account, to_addr, subject, body)
        except Exception as e:
            yield event.plain_result(f"❌ 发送失败: {e}")
            return

        account_display = account.get("name") or account.get("email") or account_name
        yield event.plain_result(
            f"✅ 发送成功\n账户: {account_display}\n收件人: {to_addr}\n主题: {subject}"
        )

    # ── Lifecycle ────────────────────────────────────────────────

    async def terminate(self):
        """插件卸载时停止后台任务与所有 IDLE 监听线程。"""
        self._idle_stop.set()
        for task in (self._check_task, self._idle_task):
            if task and not task.done():
                task.cancel()
        for task in (self._check_task, self._idle_task):
            if not task:
                continue
            try:
                await task
            except asyncio.CancelledError:
                pass
        # IDLE 读循环最长阻塞 1 秒，稍等即可退出。
        for worker in self._idle_threads.values():
            if worker.is_alive():
                worker.join(timeout=3)
        self._idle_threads.clear()
        logger.info("MailNotify: plugin terminated.")
