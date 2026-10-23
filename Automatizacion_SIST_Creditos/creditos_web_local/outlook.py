from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional


class OutlookUnavailableError(RuntimeError):
    pass


@dataclass
class DraftResult:
    success: bool
    message_id: Optional[str] = None
    preview_html: Optional[str] = None
    error: Optional[str] = None


class OutlookDraftService:
    """Creates Outlook draft items via COM automation.

    - Disabled state is reported honestly (success=False, not a fake success).
    - Persistent deduplication is handled by the caller via draft_history (DB).
    - In-session deduplication via _created_keys prevents double-clicks.
    - COM resources are released in release().
    - No automatic email sending; only drafts are created.
    """

    def __init__(self, enabled: bool = True) -> None:
        self.enabled = enabled
        self._com = None
        self._created_keys: set = set()

    def create_draft(
        self,
        subject: str,
        body_html: str,
        recipient: str,
        *,
        account: str = "",
        dedupe_key: Optional[str] = None,
    ) -> DraftResult:
        if not self.enabled:
            return DraftResult(
                success=False,
                error="Integración Outlook desactivada. Configure la cuenta en Configuración.",
            )

        identity = dedupe_key or f"{subject}|{recipient}"
        if identity in self._created_keys:
            return DraftResult(
                success=True,
                message_id=f"DUPLICATE_{identity}",
                preview_html=body_html,
            )

        is_automated = bool(
            os.environ.get("PYTEST_CURRENT_TEST")
            or os.environ.get("CI")
            or os.environ.get("GITHUB_ACTIONS")
        )
        if is_automated:
            return DraftResult(
                success=False,
                error=(
                    "La creación real de borradores de Outlook está bloqueada"
                    " en entornos automatizados."
                ),
            )

        try:
            import win32com.client  # type: ignore
        except Exception:
            return DraftResult(
                success=False,
                error="Outlook COM no disponible en esta PC (pywin32 no instalado o Outlook ausente).",
            )

        try:
            if self._com is None:
                self._com = win32com.client.Dispatch("Outlook.Application")
            mail = self._com.CreateItem(0)
            mail.Subject = subject
            mail.To = recipient
            mail.HTMLBody = body_html
            if account:
                try:
                    for acc in self._com.Session.Accounts:
                        if acc.SmtpAddress.lower() == account.lower():
                            mail.SendUsingAccount = acc
                            break
                except Exception:
                    pass  # Use default account if selection fails
            mail.Save()
            message_id = getattr(mail, "EntryID", None) or f"DRAFT_{identity}"
            self._created_keys.add(identity)
            return DraftResult(success=True, message_id=str(message_id), preview_html=body_html)
        except Exception as exc:
            return DraftResult(success=False, error=f"No se pudo crear el borrador: {exc}")

    def release(self) -> None:
        """Release COM references."""
        self._com = None


class MockOutlookDraftService:
    """Test-only mock – must not be used in production code."""

    def __init__(self, *, available: bool = True, fail: bool = False):
        self.available = available
        self.fail = fail
        self.released = False
        self._created_keys: set = set()

    def create_draft(
        self,
        subject: str,
        body_html: str,
        recipient: str,
        *,
        account: str = "",
        dedupe_key: Optional[str] = None,
    ) -> DraftResult:
        if not self.available:
            return DraftResult(success=False, error="Outlook de prueba no disponible.")
        if self.fail:
            return DraftResult(success=False, error="Mock de Outlook falló en la creación del borrador.")
        identity = dedupe_key or f"{subject}|{recipient}"
        if identity in self._created_keys:
            return DraftResult(
                success=True,
                message_id=f"MOCK_DUPLICATE_{identity}",
                preview_html=body_html,
            )
        self._created_keys.add(identity)
        return DraftResult(success=True, message_id=f"MOCK_DRAFT_{identity}", preview_html=body_html)

    def release(self) -> None:
        self.released = True
