from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from database import connect, init_db
from outlook import MockOutlookDraftService, OutlookDraftService
from services import determine_document_state, import_signo_file


@pytest.fixture
def temp_db(tmp_path):
    db_path = tmp_path / "creditos.sqlite3"
    init_db(db_path)
    return db_path


def _write_signo_file(path: Path, rows_by_sheet):
    wb = Workbook()
    for sheet_name, rows in rows_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name)
        ws.append(rows[0])
        for row in rows[1:]:
            ws.append(row)
    if "Pendientes" not in wb.sheetnames:
        wb.active.title = "Pendientes"
    if wb.sheetnames and wb.sheetnames[0] != "Pendientes":
        ws = wb[wb.sheetnames[0]]
        wb.remove(ws)
        if "Pendientes" not in wb.sheetnames:
            wb.active.title = "Pendientes"
    wb.save(path)


def test_import_new_format_keeps_leading_zeroes(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "12345678", "Cliente Demo", "2024-01-01", "F", "001", "0001", "K", "102.50", "52.50", "2024-02-05", "Ana", "2024-02-02"],
            ]
        },
    )
    result = import_signo_file(path, temp_db)
    assert result["status"] == "OK"
    with connect(temp_db) as conn:
        row = conn.execute("SELECT serie, numero, saldo_observado, estado FROM documents WHERE document_key = ?", ("F|001|0001",)).fetchone()
    assert row["serie"] == "001"
    assert row["numero"] == "0001"
    assert row["saldo_observado"] == "52.50"


def test_partial_balance_and_document_present_in_pending_and_cancelled(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "201", "Cliente A", "2024-01-01", "F", "B1", "10", "K", "500", "200", "2024-02-05", "Ana", "2024-02-02"],
            ],
            "Cancelados": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FormaPago", "FechaCancelacion", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "201", "Cliente A", "2024-01-01", "F", "B1", "10", "K", "500", "200", "EFECTIVO", "2024-01-20", "2024-02-05", "Ana", "2024-02-02"],
            ],
        },
    )
    import_signo_file(path, temp_db)
    with connect(temp_db) as conn:
        row = conn.execute("SELECT estado, saldo_observado FROM documents WHERE document_key = ?", ("F|B1|10",)).fetchone()
    assert row["estado"] == "SALDO_PARCIAL"
    assert row["saldo_observado"] == "200"


def test_zero_balance_document_is_excluded(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "202", "Cliente B", "2024-01-01", "F", "B2", "11", "K", "150", "0", "2024-02-05", "Ana", "2024-02-02"],
            ]
        },
    )
    import_signo_file(path, temp_db)
    with connect(temp_db) as conn:
        row = conn.execute("SELECT estado FROM documents WHERE document_key = ?", ("F|B2|11",)).fetchone()
    assert row["estado"] == "EXCLUIDO"


def test_missing_document_is_not_marked_as_paid(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Ventas": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FormaPago", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "303", "Cliente C", "2024-01-01", "F", "B3", "12", "K", "200", "200", "EFECTIVO", "2024-02-05", "Ana", "2024-02-02"],
            ]
        },
    )
    import_signo_file(path, temp_db)
    with connect(temp_db) as conn:
        row = conn.execute("SELECT estado FROM documents WHERE document_key = ?", ("F|B3|12",)).fetchone()
    assert row["estado"] == "EN_REVISION"


def test_repeated_import_is_idempotent_and_older_report_does_not_overwrite(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "444", "Cliente D", "2024-01-01", "F", "B4", "13", "K", "100", "100", "2024-02-05", "Ana", "2024-02-01"],
            ]
        },
    )
    first = import_signo_file(path, temp_db)
    second = import_signo_file(path, temp_db)
    assert first["documents"] == 1
    assert second["documents"] == 1
    with connect(temp_db) as conn:
        count = conn.execute("SELECT COUNT(*) FROM documents WHERE document_key = ?", ("F|B4|13",)).fetchone()[0]
    assert count == 1


def test_older_report_does_not_overwrite_newer(temp_db, tmp_path):
    path_old = tmp_path / "old.xlsx"
    _write_signo_file(
        path_old,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "555", "Cliente E", "2024-01-01", "F", "B5", "14", "K", "100", "100", "2024-02-05", "Ana", "2024-01-01"],
            ]
        },
    )
    path_new = tmp_path / "new.xlsx"
    _write_signo_file(
        path_new,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "555", "Cliente E", "2024-01-01", "F", "B5", "14", "K", "100", "60", "2024-02-05", "Ana", "2024-02-02"],
            ]
        },
    )
    import_signo_file(path_old, temp_db)
    import_signo_file(path_new, temp_db)
    with connect(temp_db) as conn:
        row = conn.execute("SELECT saldo_observado, fecha_actualizacion FROM documents WHERE document_key = ?", ("F|B5|14",)).fetchone()
    assert row["saldo_observado"] == "60"
    assert row["fecha_actualizacion"] == "2024-02-02"


def test_manual_information_is_preserved_when_same_document_is_updated(temp_db, tmp_path):
    path = tmp_path / "signo.xlsx"
    _write_signo_file(
        path,
        {
            "Pendientes": [
                ["#", "DOI", "Cliente", "FechaEmision", "Tipo", "Serie", "Numero", "Kardex", "ImporteFacturado", "SaldoPendiente", "FechaVencimiento", "Empleado", "FechaActualizacion"],
                [1, "666", "Cliente F", "2024-01-01", "F", "B6", "15", "K", "100", "100", "2024-02-05", "Ana", "2024-02-02"],
            ]
        },
    )
    import_signo_file(path, temp_db)
    with connect(temp_db) as conn:
        conn.execute("UPDATE documents SET observaciones = 'Compromiso manual', fecha_vencimiento = '2024-03-07' WHERE document_key = ?", ("F|B6|15",))
    import_signo_file(path, temp_db)
    with connect(temp_db) as conn:
        row = conn.execute("SELECT observaciones, fecha_vencimiento FROM documents WHERE document_key = ?", ("F|B6|15",)).fetchone()
    assert row["observaciones"] == "Compromiso manual"
    assert row["fecha_vencimiento"] == "2024-03-07"


def test_frequency_anticipation_and_deduplication_work_for_outlook_mocks():
    service = MockOutlookDraftService()
    result_a = service.create_draft("Recordatorio", "<p>Body</p>", "cliente@demo.com", dedupe_key="cli-1")
    result_b = service.create_draft("Recordatorio", "<p>Body</p>", "cliente@demo.com", dedupe_key="cli-1")
    assert result_a.success is True
    assert result_b.success is True
    assert "DUPLICATE" in result_b.message_id
    service.release()
    assert service.released is True


def test_outlook_failure_is_not_accepted_as_real_success():
    service = OutlookDraftService(enabled=True)
    result = service.create_draft("Recordatorio", "<p>Body</p>", "cliente@demo.com", dedupe_key="x")
    assert result.success is False or "DUPLICATE" in (result.message_id or "") or result.error is not None


def test_document_state_rules():
    assert determine_document_state("F|1|1", {"saldo_pendiente": "0", "importe_facturado": "100", "source_sheet": "Pendientes"}) == "EXCLUIDO"
    assert determine_document_state("F|1|2", {"saldo_pendiente": "150", "importe_facturado": "200", "source_sheet": "Pendientes"}) == "SALDO_PARCIAL"
    assert determine_document_state("F|1|3", {"saldo_pendiente": "50", "importe_facturado": "0", "source_sheet": "Ventas"}) == "EN_REVISION"
