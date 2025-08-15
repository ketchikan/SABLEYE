# utils/email_sender.py
from __future__ import annotations
import os
from typing import Iterable, Optional
import pythoncom  # required if used from background threads
import win32com.client as win32

class OutlookEmailSender:
    """
    Context-managed Outlook sender that accepts a ready-made HTML string.

    Features:
      - Initializes COM once (works in background threads)
      - Sends via a specific Outlook account (SendUsingAccount) if provided
      - Optionally sets SentOnBehalfOfName (requires permissions)
      - Preview vs Send switch
    """

    def __init__(
        self,
        *,
        # account_smtp: Optional[str] = None,   # the SMTP of the Outlook account to send from
        send_on_behalf_of: Optional[str] = None,  # mailbox display/SMTP for On-Behalf-Of (perm required)
        preview: bool = False
    ):
        # self.account_smtp = (account_smtp or "").lower() or None
        self.send_on_behalf_of = send_on_behalf_of
        self.preview = preview

        self._com_inited = False
        self._outlook = None
        self._account = None

    # --- Context manager API ---
    def __enter__(self) -> "OutlookEmailSender":
        pythoncom.CoInitialize()
        self._com_inited = True
        self._outlook = win32.Dispatch("Outlook.Application")
        # (No account_smtp logic here since it's not in __init__)
        return self


    def __exit__(self, exc_type, exc, tb):
        # Clean up COM
        try:
            if self._com_inited:
                pythoncom.CoUninitialize()
        finally:
            self._com_inited = False
            self._outlook = None
            self._account = None

    # --- Public API ---
    def send_html(
        self,
        *,
        html_body: str,
        to: str,
        subject: str,
        cc: str = "",
        bcc: str = "",
        attachments: Optional[Iterable[str]] = None,
        # account_smtp: Optional[str] = None,     # override per-message, if needed
        send_on_behalf_of: Optional[str] = None,# override per-message, if needed
        reply_to: Optional[str] = None,
        preview: Optional[bool] = None
    ):
        """
        Send a single HTML email. Assumes __enter__ has been called (use `with`).

        - If account_smtp is provided, uses that Outlook account via SendUsingAccount.
        - If send_on_behalf_of is provided, sets SentOnBehalfOfName (requires permissions).
        - reply_to sets the ReplyRecipients.
        """
        if self._outlook is None:
            raise RuntimeError("OutlookEmailSender must be used within a context (use `with`).")

        mail = self._outlook.CreateItem(0)  # 0 = olMailItem

        # Choose sending account (preferred way)
        # chosen_smtp = (account_smtp or self.account_smtp or "").lower() or None
        # if chosen_smtp:
        #     acct = self._find_account(chosen_smtp) if account_smtp else self._account
        #     if acct is None:
        #         raise RuntimeError(f"Outlook account not found for SMTP: {chosen_smtp}")
        #     mail.SendUsingAccount = acct

        # On-behalf-of (separate from account; requires Exchange permissions)
        sob = send_on_behalf_of or self.send_on_behalf_of
        if sob:
            mail.SentOnBehalfOfName = sob

        # Headers & body
        mail.To = to or ""
        mail.CC = cc or ""
        mail.BCC = bcc or ""
        mail.Subject = subject or ""
        mail.HTMLBody = html_body or ""

        # Reply-To
        if reply_to:
            recips = mail.ReplyRecipients
            recips.Add(reply_to)

        # Attachments
        if attachments:
            for path in attachments:
                if path and os.path.exists(path):
                    mail.Attachments.Add(path)

        do_preview = self.preview if preview is None else preview
        if do_preview:
            mail.Display(False)   # Show window; user can click Send
        else:
            mail.Send()

        return mail  # return the MailItem in case caller wants to inspect it

    # --- Helpers ---
    def _find_account(self, query_lower: str):
        """Match by SMTP or DisplayName (case-insensitive)."""
        session = self._outlook.Session
        for acct in session.Accounts:
            try:
                smtp = str(getattr(acct, "SmtpAddress", "")).lower()
            except Exception:
                smtp = ""
            name = str(getattr(acct, "DisplayName", "")).lower()

            if query_lower == smtp or query_lower == name:
                return acct
        return None

