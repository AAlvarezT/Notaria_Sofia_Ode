from __future__ import annotations

import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional


class OutlookUnavailableError(RuntimeError):
    pass


@dataclass
class DraftResult:
    success: bool
    message_id: Optional[str] = None
    preview_html: Optional[str] = None
    error: Optional[str] = None


class OutlookDraftService:
    def __init__(self, enabled: bool = True) -> None:
        self.enabled = enabled
        self._com = None
        self._created_keys: set[str] = set()

    def create_draft(self, subject: str, body_html: str, recipient: str, *, account: str = "", dedupe_key: Optional[str] = None) -> DraftResult:
        if not self.enabled:
            return DraftResult(success=False, error="Integración Outlook deshabilitada en configuración.")
        identity = dedupe_key or f"{subject}|{recipient}"
        if identity in self._created_keys:
            return DraftResult(success=True, message_id=f"DUPLICATE_{identity}", preview_html=body_html)

        is_automated_runtime = bool(os.environ.get("PYTEST_CURRENT_TEST") or os.environ.get("CI") or os.environ.get("GITHUB_ACTIONS"))
        if is_automated_runtime:
            return DraftResult(success=False, error="La creación real de borradores de Outlook está bloqueada en entornos automatizados o no interactivos.")

        try:
            import win32com.client  # type: ignore
        except Exception as exc:  # pragma: no cover - Windows-specific path
            return DraftResult(success=False, error="Outlook no está disponible en esta PC: no se pudo inicializar la automatización COM.")

        try:
            if self._com is None:
                self._com = win32com.client.Dispatch("Outlook.Application")
            mail = self._com.CreateItem(0)
            mail.Subject = subject
            mail.To = recipient
            mail.HTMLBody = body_html
            mail.Save()
            message_id = getattr(mail, "EntryID", None) or f"DRAFT_{identity}"
            self._created_keys.add(identity)
            return DraftResult(success=True, message_id=str(message_id), preview_html=body_html)
        except Exception as exc:  # pragma: no cover - COM-specific path
            return DraftResult(success=False, error=f"No se pudo crear el borrador: {exc}")

    def release(self) -> None:
        self._com = None


class MockOutlookDraftService:
    def __init__(self, *, available: bool = True, fail: bool = False):
        self.available = available
        self.fail = fail
        self.released = False
        self._created_keys: set[str] = set()

    def create_draft(self, subject: str, body_html: str, recipient: str, *, account: str = "", dedupe_key: Optional[str] = None) -> DraftResult:
        if not self.available:
            return DraftResult(success=False, error="Outlook de prueba no disponible.")
        if self.fail:
            return DraftResult(success=False, error="Mock de Outlook falló en la creación del borrador.")
        identity = dedupe_key or f"{subject}|{recipient}"
        if identity in self._created_keys:
            return DraftResult(success=True, message_id=f"MOCK_DUPLICATE_{identity}", preview_html=body_html)
        self._created_keys.add(identity)
        return DraftResult(success=True, message_id=f"MOCK_DRAFT_{identity}", preview_html=body_html)

    def release(self) -> None:
        self.released = True
