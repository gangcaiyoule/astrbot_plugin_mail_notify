import asyncio
from datetime import datetime, timezone

from astrbot.api import AstrBotConfig, logger
from astrbot.api.event import AstrMessageEvent, MessageChain, filter
from astrbot.api.star import Context, Star, register

from .imap_client import imap_fetch_new, imap_query_since, is_recent_email


@register(
    "astrbot_plugin_mail_notify",
    "YourName",
    "监控邮箱新邮件并通过 QQ 私聊发送通知",
    "1.1.3",
)
class MailNotifyPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config
        # Background polling task created during initialize().
        self._check_task: asyncio.Task | None = None
        # Runtime-only status used by /mail_status; not persisted.
        self._last_check_time: dict[str, str] = {}
        self._account_status: dict[str, str] = {}

    async def initialize(self):
        """Called after plugin instantiation, start background mail check loop."""
        # The plugin starts its own polling loop after AstrBot loads it.
        self._check_task = asyncio.create_task(self._check_loop())
        logger.info("MailNotify: background check loop started.")

    # ── Background Loop ──────────────────────────────────────────

    async def _check_loop(self):
        await asyncio.sleep(10)  # wait for everything to settle
        while True:
            try:
                interval = self.config.get("check_interval", 5)
                notify_umo = self.config.get("notify_umo", "")
                # Poll only after a target conversation is bound.
                if notify_umo:
                    accounts = self.config.get("mail_accounts", [])
                    for account in accounts:
                        if not account.get("email") or not account.get("imap_server"):
                            continue
                        try:
                            # Each account is checked independently so one failure
                            # does not block the others.
                            await self._check_account(account, notify_umo)
                            self._account_status[account["email"]] = "✅ 正常"
                        except Exception as e:
                            self._account_status[account["email"]] = f"❌ {str(e)[:80]}"
                            logger.error(
                                f"MailNotify: check failed for {account['email']}: {e}"
                            )
                        self._last_check_time[account["email"]] = (
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        )
                await asyncio.sleep(max(interval, 1) * 60)
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"MailNotify: loop error: {e}")
                await asyncio.sleep(60)

    # ── IMAP Logic ───────────────────────────────────────────────

    async def _check_account(self, account: dict, notify_umo: str):
        # Every mailbox has its own persisted keys so multiple accounts do not clash.
        account_email = account["email"]
        max_body_len = self.config.get("max_body_length", 500)

        uid_key = f"last_uid_{account_email}"
        init_key = f"init_time_{account_email}"
        # KV data is stored by AstrBot in data/data_v4.db.
        last_uid = await self.get_kv_data(uid_key, 0) or 0
        init_time = await self.get_kv_data(init_key, "")

        is_first_run = not init_time
        if is_first_run:
            # First run records the start time and current UID baseline,
            # preventing old inbox history from being pushed as new mail.
            init_time = datetime.now(timezone.utc).isoformat()
            await self.put_kv_data(init_key, init_time)

        # imaplib is blocking, so the actual mailbox query runs in a worker thread.
        new_emails, new_max_uid = await asyncio.to_thread(
            imap_fetch_new, account, last_uid, max_body_len
        )

        if new_max_uid > last_uid:
            # Persist the newest UID after a successful fetch for next incremental sync.
            await self.put_kv_data(uid_key, new_max_uid)

        if is_first_run:
            if new_max_uid > 0:
                logger.info(
                    f"MailNotify: initialized {account_email}, max UID = {new_max_uid}"
                )
            return

        init_dt = datetime.fromisoformat(init_time)
        for mail_info in new_emails:
            # Double-check the parsed mail date against init_time to avoid edge cases
            # where a just-fetched mail still belongs to the historical backlog.
            if is_recent_email(mail_info, init_dt):
                await self._send_notification(account, mail_info, notify_umo)

    # ── Notification ─────────────────────────────────────────────

    async def _send_notification(self, account: dict, mail_info: dict, notify_umo: str):
        account_name = account.get("name") or account["email"]
        use_ai = self.config.get("ai_summary", False)
        body_text = mail_info["body"]

        # Optional AI summary replaces the raw preview text when enabled.
        if use_ai and body_text:
            body_text = await self._try_ai_summary(mail_info, notify_umo, body_text)

        # Build a plain text message and let AstrBot deliver it to the bound session.
        lines = [
            f"📬 新邮件通知 [{account_name}]",
            "━━━━━━━━━━━━━━━━",
            f"📤 发件人: {mail_info['from_name']}",
        ]
        if mail_info["from_addr"] and mail_info["from_addr"] != mail_info["from_name"]:
            lines[-1] += f" <{mail_info['from_addr']}>"
        lines.append(f"📋 主题: {mail_info['subject']}")
        lines.append(f"🕐 时间: {mail_info['date']}")
        if body_text:
            label = "📝 AI摘要" if use_ai else "📝 预览"
            lines.append(f"{label}: {body_text}")

        chain = MessageChain().message("\n".join(lines))
        await self.context.send_message(notify_umo, chain)

    async def _try_ai_summary(
        self, mail_info: dict, notify_umo: str, fallback: str
    ) -> str:
        try:
            # Reuse the chat provider already bound to the target conversation.
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

    # ── Commands ─────────────────────────────────────────────────

    @filter.command("mail_bind")
    async def mail_bind(self, event: AstrMessageEvent):
        """绑定当前会话为邮件通知目标"""
        umo = event.unified_msg_origin
        # This is regular plugin config, not KV state, so it is saved in config files.
        self.config["notify_umo"] = umo
        self.config.save_config()
        yield event.plain_result(f"✅ 已绑定当前会话为邮件通知目标。\n会话 ID: {umo}")

    @filter.command("mail_status")
    async def mail_status(self, event: AstrMessageEvent):
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
            last = self._last_check_time.get(addr, "尚未检查")
            lines.append(f"📧 {name} ({addr})")
            lines.append(f"   状态: {status}")
            lines.append(f"   最近检查: {last}")

        yield event.plain_result("\n".join(lines))

    @filter.command("mail_check")
    async def mail_check(self, event: AstrMessageEvent):
        """立即手动检查所有邮箱"""
        accounts = self.config.get("mail_accounts", [])
        if not accounts:
            yield event.plain_result(
                "📭 未配置任何邮箱账户，请在 WebUI 插件配置中添加。"
            )
            return

        notify_umo = self.config.get("notify_umo", "") or event.unified_msg_origin
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
            self._last_check_time[email_addr] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        if errors:
            yield event.plain_result("⚠️ 部分邮箱检查失败:\n" + "\n".join(errors))
        else:
            yield event.plain_result("✅ 所有邮箱检查完成。")

    @filter.command("mail_query")
    async def mail_query(
        self, event: AstrMessageEvent, account_name: str, since_date: str
    ):
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
                imap_query_since, target, since_dt, max_body_len
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
            lines.append(f"   📤 {m['from_name']}  🕐 {m['date']}")
        yield event.plain_result("\n".join(lines))

    # ── Lifecycle ────────────────────────────────────────────────

    async def terminate(self):
        """Cancel background task on plugin unload."""
        if self._check_task and not self._check_task.done():
            self._check_task.cancel()
            try:
                await self._check_task
            except asyncio.CancelledError:
                pass
        logger.info("MailNotify: plugin terminated.")
