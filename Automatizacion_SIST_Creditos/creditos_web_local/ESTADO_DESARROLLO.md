# Créditos Web Local — Estado del desarrollo

## Qué está implementado y probado con mocks

### 1. Arranque de la aplicación ✅
- `app.py` importa `get_recent_imports` correctamente desde `database.py`.
- La carga de archivos usa `tempfile.NamedTemporaryFile` en lugar de `/tmp` → funciona en Windows y Linux.
- El endpoint `/health` responde con estado de la BD y del scheduler.
- El módulo `scheduler.py` arranca un hilo de fondo al iniciar la app.

### 2. Importación de Excel ✅
- Usa `iter_rows(values_only=True)` → no llama a `ws.cell()` en modo `read_only`.
- Detecta fila de título (fila 1) y fila de cabeceras (fila 2 ó la primera que contenga TIPO+SERIE+NUMERO, hasta la fila 40).
- Conserva ceros iniciales de DOI, serie y número (strings de openpyxl se devuelven tal cual; enteros numéricos se convierten sin decimales).
- `fecha_emision`, `fecha_vencimiento`, `fecha_cancelacion` → fecha ISO (`YYYY-MM-DD`).
- `fecha_actualizacion` → timestamp ISO con hora (`YYYY-MM-DDTHH:MM:SS`) para comparar versiones del mismo día.
- `0000-00-00` se trata como fecha ausente (NULL).
- Montos inválidos (texto no numérico) se registran en errors y el comprobante se envía a `EN_REVISION`, sin silenciar el problema con saldo=0.
- Contacts, comentarios y compromisos de pago no son modificados por la importación.
- Observaciones y fecha de vencimiento manual se conservan tras reimportación.

### 3. Reglas de negocio ✅
| Situación | Estado asignado |
|-----------|----------------|
| Saldo ≤ 0 | `EXCLUIDO` |
| Saldo < Importe (ambos > 0) | `SALDO_PARCIAL` |
| Saldo = Importe, en hoja Pendientes | `PENDIENTE` |
| Solo en Ventas/Cancelados con saldo > 0 | `EN_REVISION` |
| Monto inválido | `EN_REVISION` |

- `SALDO_PARCIAL` tiene `revision_required = 0` (es un estado válido, no una inconsistencia).
- Cuando el comprobante aparece en Pendientes y Cancelados, el saldo de Pendientes prevalece.
- Aparece en Cancelados por una detracción → sigue siendo cobrable si tiene saldo.
- Reimportar el mismo archivo es idempotente (upsert por document_key).
- Un reporte más antiguo (fecha_actualizacion menor) no sobreescribe uno más reciente.
- El tiempo del timestamp se usa para desempatar importaciones del mismo día.

### 4. Cobranza con saldos parciales ✅
- `/cartera` filtra por `estado IN ('PENDIENTE', 'SALDO_PARCIAL')`.
- `/cobranza` devuelve los mismos estados con saldo > 0 y `revision_required = 0`.
- El dashboard suma la deuda total de `PENDIENTE` + `SALDO_PARCIAL`.
- `list_documents` acepta lista de estados: `{"estado": ["PENDIENTE", "SALDO_PARCIAL"]}`.

### 5. Integración Outlook ✅ (lógica; COM no probado en CI)
- `OutlookDraftService(enabled=False)` devuelve `success=False` con mensaje claro ("desactivada"). **No reporta éxito falso.**
- `MockOutlookDraftService` solo se usa en pruebas; no aparece en código de producción.
- `/draft` requiere email válido (con `@`); el nombre del cliente nunca se usa como destinatario.
- Deduplicación persistente: usa la tabla `draft_history` (SQLite), sobrevive reinicios.
- Deduplicación en sesión: `_created_keys` evita doble clic.
- `OutlookDraftService.release()` libera referencias COM.
- No hay envío automático de correos; solo creación de borradores.
- Se registra en `draft_history` solo cuando el borrador real se crea con éxito.
- Gestión de contactos: endpoints `GET /contacts/{doc_key}` y `POST /contacts`.

### 6. Scheduler mínimo ✅
- Hilo daemon único; usa `threading.Lock` para evitar múltiples instancias.
- Lee `signo_folder`, `check_frequency_minutes` y `days_advance` de la configuración.
- Importa archivos Excel nuevos/modificados en la carpeta configurada.
- Cuenta documentos que requieren recordatorio según `days_advance`.
- Estado disponible en `/health` y vía `get_scheduler_status()`.
- No inicia si `signo_folder` está vacío o no existe.
- Reimportación evitada con huella de archivo (path + mtime).

---

## Pruebas automatizadas (25 tests, 25 ✅)

| Test | Cobertura |
|------|-----------|
| `test_import_new_format_keeps_leading_zeroes` | Ceros iniciales en serie/número |
| `test_title_row_plus_header_row` | Fila de título + fila de cabeceras |
| `test_partial_balance_and_document_present_in_pending_and_cancelled` | SALDO_PARCIAL en Pendientes+Cancelados |
| `test_partial_balance_in_all_three_sheets` | F\|FE01\|0030712 en las tres hojas |
| `test_saldo_parcial_has_no_revision_required` | SALDO_PARCIAL no bloquea cobranza |
| `test_partial_balance_included_in_cobranza_list` | list_documents incluye SALDO_PARCIAL |
| `test_zero_balance_document_is_excluded` | Saldo cero → EXCLUIDO |
| `test_missing_document_is_not_marked_as_paid` | Solo en Ventas → EN_REVISION |
| `test_repeated_import_is_idempotent_...` | Reimportación sin duplicados |
| `test_reimport_no_duplicates` | COUNT = 2 tras doble importación |
| `test_older_report_does_not_overwrite_newer` | Reporte antiguo no sobreescribe |
| `test_old_report_does_not_overwrite_newer_data` | Importación de archivo viejo no daña datos recientes |
| `test_date_normalized_to_iso_and_timestamp_preserves_time` | Fechas ISO + timestamp con hora |
| `test_same_day_update_different_hours_newer_wins` | Desempate por hora del día |
| `test_absent_date_0000_treated_as_null` | 0000-00-00 → NULL |
| `test_manual_information_is_preserved_when_same_document_is_updated` | observaciones preservadas |
| `test_manual_contacts_preserved_on_reimport` | Contactos manuales sobreviven reimportación |
| `test_invalid_amount_sends_to_revision` | Monto inválido → EN_REVISION |
| `test_cancelled_sheet_alone_does_not_pay_document` | Cancelados con saldo > 0 no es EXCLUIDO |
| `test_document_state_rules` | Reglas directas de `determine_document_state` |
| `test_outlook_disabled_returns_no_success` | Outlook desactivado → success=False |
| `test_mock_outlook_dedup_in_session` | Dedup en sesión con MockOutlookDraftService |
| `test_persistent_draft_deduplication` | Dedup persistente via draft_history |
| `test_frequency_anticipation_and_deduplication_work_for_outlook_mocks` | Mock release + dedup |
| `test_outlook_failure_is_not_accepted_as_real_success` | Entorno automatizado bloquea COM real |

---

## Archivo real (20261021133448BWv9Yl1k.xlsx)

**No estuvo disponible** en el workspace durante esta sesión. Los resultados esperados
(978 ventas, 28 pendientes, 12 cancelados, etc.) **no pudieron ser verificados aquí**.
Los fixtures de prueba replican la estructura del archivo real y los casos de borde descritos.

---

## Pendientes para validación en la computadora de la usuaria

### Prioridad 1 — Verificar con el archivo real
1. Importar `20261021133448BWv9Yl1k.xlsx` y confirmar:
   - 978 comprobantes únicos, 28 pendientes, 12 cancelados, 1018 filas leídas.
   - Cartera cobrable: S/ 28,180 en 28 comprobantes (27 PENDIENTE + 1 SALDO_PARCIAL).
   - `F|FE01|0030712` → `SALDO_PARCIAL`, saldo S/ 1,320, visible en `/cobranza`.

### Prioridad 2 — Verificar en Windows con Outlook instalado
2. Activar la integración en Configuración → cuenta de Outlook configurada.
3. Crear un borrador real desde `/cobranza` → confirmar que aparece en Borradores de Outlook.
4. Verificar que reimportar o recargar la página no crea borrador duplicado.
5. Verificar que `release()` libera la conexión COM al terminar.

### Prioridad 3 — Scheduler con carpeta real
6. Configurar `signo_folder` apuntando a la carpeta donde llegan los archivos de SIGNO.
7. Dejar correr el servidor y confirmar que la carpeta se revisa y los archivos se importan automáticamente.
8. Verificar que el conteo de recordatorios pendientes en `/health` es correcto.

### Prioridad 4 — Interfaz de contactos
9. Registrar contactos desde la UI (endpoint `/contacts`).
10. Confirmar que al crear un borrador, el correo del contacto se completa automáticamente.

### Prioridad 5 — Templates
11. Revisar que las vistas HTML (`cobranza.html`, `dashboard.html`) muestran correctamente
    documentos SALDO_PARCIAL en la lista cobrable y los totales del dashboard.
