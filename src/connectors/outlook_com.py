from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
import pythoncom
import win32com.client

from src.config import settings
from src.connectors.graph_client import (
    CalendarEvent,
    EmailMessage,
    html_to_clean_markdown,
    markdown_to_html,
)

# Constantes OlDefaultFolders de Outlook COM
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9

# Constantes OlItemType
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1

# Constantes OlBusyStatus
BUSY_STATUS_MAP = {
    0: "free",
    1: "tentative",
    2: "busy",
    3: "oof",
    4: "working_elsewhere",
}


def _extract_smtp_address(mail_item: Any) -> str:
    """
    Obtiene el correo SMTP real incluso cuando el remitente es interno de Exchange
    (donde SenderEmailAddress devuelve una ruta X500 '/O=EXCHANGELABS/...').
    """
    try:
        if getattr(mail_item, "SenderEmailType", "") == "EX":
            sender = getattr(mail_item, "Sender", None)
            if sender is not None:
                ex_user = sender.GetExchangeUser()
                if ex_user is not None and ex_user.PrimarySmtpAddress:
                    return str(ex_user.PrimarySmtpAddress)
                ex_dist = sender.GetExchangeDistributionList()
                if ex_dist is not None and ex_dist.PrimarySmtpAddress:
                    return str(ex_dist.PrimarySmtpAddress)
            # Fallback mediante PropertyAccessor (PR_SENDER_SMTP_ADDRESS)
            pa = mail_item.PropertyAccessor
            smtp = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
            if smtp:
                return str(smtp)
    except Exception:
        pass
    return str(getattr(mail_item, "SenderEmailAddress", "") or "")


def _extract_recipient_list(mail_item: Any) -> tuple[List[str], List[str]]:
    """Separa los destinatarios To (Type=1) y CC (Type=2) en listas de correos SMTP."""
    to_list: List[str] = []
    cc_list: List[str] = []
    try:
        for rec in mail_item.Recipients:
            addr = ""
            try:
                pa = rec.PropertyAccessor
                addr = str(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F") or "")
            except Exception:
                addr = str(getattr(rec, "Address", "") or "")
            if not addr or "/" in addr:
                addr = str(getattr(rec, "Name", "") or "")
            r_type = getattr(rec, "Type", 1)
            if r_type == 1:
                to_list.append(addr)
            elif r_type == 2:
                cc_list.append(addr)
    except Exception:
        pass
    return to_list, cc_list


class OutlookLocalClient:
    """
    Conector local para Microsoft Outlook clásico en Windows (vía COM / MAPI).
    No requiere registro en Azure Entra ID ni consentimiento de administrador TI.
    """

    def __init__(self, account_email: Optional[str] = None) -> None:
        self.account_email = (account_email or settings.m365_account_email).strip().lower()

    def _connect(self) -> tuple[Any, Any]:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            return outlook, namespace
        except Exception as exc:
            raise RuntimeError(
                "No se pudo conectar con Outlook de escritorio en Windows. "
                "Asegúrate de tener instalado y abierto Outlook clásico (no 'New Outlook' web). "
                f"Detalle: {exc}"
            ) from exc

    def _get_folder(self, namespace: Any, folder_name: str = "inbox") -> Any:
        folder_map = {
            "inbox": OL_FOLDER_INBOX,
            "drafts": OL_FOLDER_DRAFTS,
            "sentitems": OL_FOLDER_SENT_MAIL,
            "calendar": OL_FOLDER_CALENDAR,
        }
        folder_id = folder_map.get(folder_name.lower(), OL_FOLDER_INBOX)

        # Si el usuario especificó M365_ACCOUNT_EMAIL en .env, buscar ese buzón específico
        if self.account_email:
            try:
                for store in namespace.Stores:
                    display = str(getattr(store, "DisplayName", "")).lower()
                    if self.account_email in display:
                        return store.GetDefaultFolder(folder_id)
            except Exception:
                pass

        return namespace.GetDefaultFolder(folder_id)

    def _parse_mail_item(self, item: Any, include_body: bool = False) -> EmailMessage:
        entry_id = str(getattr(item, "EntryID", ""))
        conv_id = str(getattr(item, "ConversationID", "") or "")
        subject = str(getattr(item, "Subject", "") or "(Sin asunto)")
        sender_name = str(getattr(item, "SenderName", "") or "")
        sender_email = _extract_smtp_address(item)
        to_list, cc_list = _extract_recipient_list(item)

        received_str = ""
        try:
            rcvd = getattr(item, "ReceivedTime", None)
            if rcvd is not None:
                received_str = rcvd.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass

        unread = bool(getattr(item, "UnRead", False))
        plain_body = str(getattr(item, "Body", "") or "").strip()
        preview = " ".join(plain_body.split())[:280]

        body_md = ""
        if include_body:
            html_body = str(getattr(item, "HTMLBody", "") or "")
            body_md = html_to_clean_markdown(html_body) if html_body else plain_body

        return EmailMessage(
            id=entry_id,
            conversation_id=conv_id,
            subject=subject,
            sender_name=sender_name,
            sender_email=sender_email,
            to_recipients=to_list,
            cc_recipients=cc_list,
            received_datetime=received_str,
            is_read=not unread,
            body_preview=preview,
            body_markdown=body_md,
            web_link=None,
        )

    def get_my_profile(self) -> Dict[str, Any]:
        """Obtiene la cuenta activa en el cliente local de Outlook."""
        _, namespace = self._connect()
        accounts_list: List[str] = []
        primary_name = ""
        primary_email = ""

        try:
            for acc in namespace.Accounts:
                smtp = str(getattr(acc, "SmtpAddress", "") or "")
                disp = str(getattr(acc, "DisplayName", "") or "")
                accounts_list.append(smtp or disp)
                if not primary_email or (self.account_email and self.account_email in smtp.lower()):
                    primary_email = smtp
                    primary_name = disp
        except Exception:
            pass

        if not primary_email:
            try:
                user = namespace.CurrentUser
                primary_name = str(getattr(user, "Name", "Usuario Outlook"))
                primary_email = str(getattr(user, "Address", ""))
            except Exception:
                primary_name = "Usuario Outlook Local"

        return {
            "backend": "outlook_local (Windows COM)",
            "displayName": primary_name,
            "mail": primary_email,
            "userPrincipalName": primary_email,
            "available_accounts": accounts_list,
        }

    def list_emails(
        self,
        unread_only: bool = True,
        limit: int = 10,
        folder: str = "inbox",
        include_body: bool = False,
    ) -> List[EmailMessage]:
        _, namespace = self._connect()
        target_folder = self._get_folder(namespace, folder)
        items = target_folder.Items
        if unread_only:
            items = items.Restrict("[UnRead] = True")
        items.Sort("[ReceivedTime]", True)

        results: List[EmailMessage] = []
        for item in items:
            try:
                # Class 43 = olMail
                if getattr(item, "Class", 0) != 43:
                    continue
                results.append(self._parse_mail_item(item, include_body=include_body))
                if len(results) >= limit:
                    break
            except Exception:
                continue
        return results

    def get_email(self, message_id: str) -> EmailMessage:
        _, namespace = self._connect()
        item = namespace.GetItemFromID(message_id)
        return self._parse_mail_item(item, include_body=True)

    def search_emails(self, query: str, limit: int = 10) -> List[EmailMessage]:
        _, namespace = self._connect()
        inbox = self._get_folder(namespace, "inbox")
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        q_lower = query.strip().lower()
        results: List[EmailMessage] = []
        checked = 0
        for item in items:
            checked += 1
            if checked > 250:
                break
            try:
                if getattr(item, "Class", 0) != 43:
                    continue
                subj = str(getattr(item, "Subject", "") or "").lower()
                s_name = str(getattr(item, "SenderName", "") or "").lower()
                s_email = _extract_smtp_address(item).lower()
                body_snip = str(getattr(item, "Body", "") or "")[:1000].lower()
                if q_lower in subj or q_lower in s_name or q_lower in s_email or q_lower in body_snip:
                    results.append(self._parse_mail_item(item, include_body=False))
                    if len(results) >= limit:
                        break
            except Exception:
                continue
        return results

    def create_reply_draft(self, message_id: str, reply_markdown: str, reply_all: bool = False) -> Dict[str, Any]:
        _, namespace = self._connect()
        item = namespace.GetItemFromID(message_id)
        reply = item.ReplyAll() if reply_all else item.Reply()
        reply_html = markdown_to_html(reply_markdown)
        reply.HTMLBody = reply_html + "<br><hr>" + str(getattr(reply, "HTMLBody", "") or "")
        reply.Save()
        return {
            "id": str(getattr(reply, "EntryID", "")),
            "subject": str(getattr(reply, "Subject", "")),
            "webLink": None,
        }

    def send_reply(self, message_id: str, reply_markdown: str, reply_all: bool = False) -> None:
        _, namespace = self._connect()
        item = namespace.GetItemFromID(message_id)
        reply = item.ReplyAll() if reply_all else item.Reply()
        reply_html = markdown_to_html(reply_markdown)
        reply.HTMLBody = reply_html + "<br><hr>" + str(getattr(reply, "HTMLBody", "") or "")
        reply.Send()

    def send_new_email(
        self,
        to_recipients: List[str],
        subject: str,
        body_markdown: str,
        cc_recipients: Optional[List[str]] = None,
        save_to_sent_items: bool = True,
    ) -> None:
        outlook, namespace = self._connect()
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        mail.To = "; ".join(to_recipients)
        if cc_recipients:
            mail.CC = "; ".join(cc_recipients)
        mail.Subject = subject
        mail.HTMLBody = markdown_to_html(body_markdown)
        if self.account_email:
            try:
                for acc in namespace.Accounts:
                    if self.account_email in str(getattr(acc, "SmtpAddress", "")).lower():
                        mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                        break
            except Exception:
                pass
        mail.Send()

    def create_new_draft(
        self,
        to_recipients: List[str],
        subject: str,
        body_markdown: str,
        cc_recipients: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        outlook, _ = self._connect()
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        mail.To = "; ".join(to_recipients)
        if cc_recipients:
            mail.CC = "; ".join(cc_recipients)
        mail.Subject = subject
        mail.HTMLBody = markdown_to_html(body_markdown)
        mail.Save()
        return {
            "id": str(getattr(mail, "EntryID", "")),
            "subject": subject,
            "webLink": None,
        }

    def mark_as_read(self, message_id: str) -> None:
        _, namespace = self._connect()
        item = namespace.GetItemFromID(message_id)
        item.UnRead = False
        item.Save()

    # =========================================================================
    # FUNCIONES DE CALENDARIO LOCAL EN OUTLOOK (WINDOWS COM)
    # =========================================================================

    def list_calendar_events(
        self,
        start_dt: datetime,
        end_dt: datetime,
        limit: int = 50,
    ) -> List[CalendarEvent]:
        _, namespace = self._connect()
        cal_folder = self._get_folder(namespace, "calendar")
        items = cal_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Filtro en formato compatible con Outlook Restrict
        start_naive = start_dt.replace(tzinfo=None)
        end_naive = end_dt.replace(tzinfo=None)

        results: List[CalendarEvent] = []
        for item in items:
            try:
                ev_start = getattr(item, "Start", None)
                ev_end = getattr(item, "End", None)
                if ev_start is None or ev_end is None:
                    continue
                ev_s_naive = datetime(
                    ev_start.year, ev_start.month, ev_start.day, ev_start.hour, ev_start.minute
                )
                ev_e_naive = datetime(
                    ev_end.year, ev_end.month, ev_end.day, ev_end.hour, ev_end.minute
                )
                if ev_s_naive > end_naive:
                    break
                if ev_e_naive <= start_naive:
                    continue

                busy_code = int(getattr(item, "BusyStatus", 2))
                req_attendees = str(getattr(item, "RequiredAttendees", "") or "")
                att_list = [a.strip() for a in req_attendees.split(";") if a.strip()]

                results.append(
                    CalendarEvent(
                        id=str(getattr(item, "EntryID", "")),
                        subject=str(getattr(item, "Subject", "") or "(Sin título)"),
                        start_local=ev_s_naive.strftime("%Y-%m-%d %H:%M"),
                        end_local=ev_e_naive.strftime("%Y-%m-%d %H:%M"),
                        is_all_day=bool(getattr(item, "AllDayEvent", False)),
                        show_as=BUSY_STATUS_MAP.get(busy_code, "busy"),
                        location=str(getattr(item, "Location", "") or ""),
                        organizer_name=str(getattr(item, "Organizer", "") or ""),
                        organizer_email="",
                        attendees=att_list,
                        is_online_meeting=False,
                        body_preview=" ".join(str(getattr(item, "Body", "") or "").split())[:200],
                        web_link=None,
                    )
                )
                if len(results) >= limit:
                    break
            except Exception:
                continue
        return results

    def get_day_schedule_and_free_slots(
        self,
        target_date_str: Optional[str] = None,
        work_start_hour: int = 8,
        work_end_hour: int = 18,
    ) -> Dict[str, Any]:
        local_tz = datetime.now().astimezone().tzinfo
        now_local = datetime.now(local_tz)

        if target_date_str:
            base_date = datetime.strptime(target_date_str.strip()[:10], "%Y-%m-%d").date()
        else:
            base_date = now_local.date()

        day_start = datetime(base_date.year, base_date.month, base_date.day, 0, 0, 0, tzinfo=local_tz)
        day_end = datetime(base_date.year, base_date.month, base_date.day, 23, 59, 59, tzinfo=local_tz)

        events = self.list_calendar_events(day_start, day_end)

        work_start = datetime(base_date.year, base_date.month, base_date.day, work_start_hour, 0, tzinfo=local_tz)
        work_end = datetime(base_date.year, base_date.month, base_date.day, work_end_hour, 0, tzinfo=local_tz)

        busy_intervals: List[tuple[datetime, datetime]] = []
        for ev in events:
            if ev.show_as == "free":
                continue
            if ev.is_all_day:
                busy_intervals.append((work_start, work_end))
                continue
            try:
                ev_s = datetime.strptime(ev.start_local, "%Y-%m-%d %H:%M").replace(tzinfo=local_tz)
                ev_e = datetime.strptime(ev.end_local, "%Y-%m-%d %H:%M").replace(tzinfo=local_tz)
                s_clamped = max(work_start, ev_s)
                e_clamped = min(work_end, ev_e)
                if e_clamped > s_clamped:
                    busy_intervals.append((s_clamped, e_clamped))
            except Exception:
                pass

        busy_intervals.sort(key=lambda x: x[0])
        merged_busy: List[tuple[datetime, datetime]] = []
        for s, e in busy_intervals:
            if not merged_busy or s > merged_busy[-1][1]:
                merged_busy.append((s, e))
            else:
                merged_busy[-1] = (merged_busy[-1][0], max(merged_busy[-1][1], e))

        free_slots: List[Dict[str, Any]] = []
        cursor = work_start
        for b_start, b_end in merged_busy:
            if b_start > cursor:
                mins = int((b_start - cursor).total_seconds() // 60)
                if mins >= 15:
                    free_slots.append(
                        {
                            "start": cursor.strftime("%H:%M"),
                            "end": b_start.strftime("%H:%M"),
                            "duration_minutes": mins,
                            "period": "mañana" if cursor.hour < 12 else "tarde",
                        }
                    )
            cursor = max(cursor, b_end)

        if cursor < work_end:
            mins = int((work_end - cursor).total_seconds() // 60)
            if mins >= 15:
                free_slots.append(
                    {
                        "start": cursor.strftime("%H:%M"),
                        "end": work_end.strftime("%H:%M"),
                        "duration_minutes": mins,
                        "period": "mañana" if cursor.hour < 12 else "tarde",
                    }
                )

        afternoon_free_mins = sum(
            slot["duration_minutes"]
            for slot in free_slots
            if slot["period"] == "tarde" or int(slot["end"].split(":")[0]) > 12
        )

        return {
            "date": base_date.strftime("%Y-%m-%d"),
            "current_local_time": now_local.strftime("%Y-%m-%d %H:%M"),
            "work_hours_evaluated": f"{work_start_hour:02d}:00 - {work_end_hour:02d}:00",
            "total_events": len(events),
            "events": [e.model_dump() for e in events],
            "free_slots": free_slots,
            "is_afternoon_mostly_free": afternoon_free_mins >= 180,
        }

    def create_calendar_event(
        self,
        subject: str,
        start_local_str: str,
        end_local_str: str,
        body_markdown: str = "",
        attendees_emails: Optional[List[str]] = None,
        location: str = "",
    ) -> Dict[str, Any]:
        outlook, _ = self._connect()
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start_local_str.strip().replace("T", " ")[:16]
        appt.End = end_local_str.strip().replace("T", " ")[:16]
        if location:
            appt.Location = location
        if body_markdown:
            appt.Body = body_markdown
        if attendees_emails:
            appt.MeetingStatus = 1  # olMeeting
            for addr in attendees_emails:
                if addr and "@" in addr:
                    appt.Recipients.Add(addr.strip())
        appt.Save()
        if attendees_emails:
            appt.Send()
        return {
            "id": str(getattr(appt, "EntryID", "")),
            "subject": subject,
            "webLink": None,
        }
