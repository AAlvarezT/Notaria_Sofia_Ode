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
- `/draft` requiere email válido con regex (más estricto que validar solo `@`); el nombre del cliente nunca se usa como destinatario.
- Deduplicación persistente: usa la tabla `draft_history` (SQLite), sobrevive reinicios.
- Deduplicación concurrente: reserva de slot `PROCESSING` en DB antes de crear el borrador para evitar duplicados simultáneos.
- Deduplicación en sesión: `_created_keys` evita doble clic.
- `OutlookDraftService.release()` libera referencias COM.
- No hay envío automático de correos; solo creación de borradores.
- Se registra en `draft_history` solo cuando el borrador real se crea con éxito.
- Si no encuentra la cuenta de Outlook configurada, devuelve error claro.
- Gestión de contactos: endpoints `GET /contacts/{doc_key}` y `POST /contacts` + formulario web en `/cliente/{document_key}` para alta/edición.

### 6. Scheduler mínimo ✅
- Hilo daemon único por proceso.
- Protección entre procesos con lock file (`scheduler.lock` + `msvcrt.locking`): si otro proceso ya está revisando, este proceso no inicia scheduler y lo informa.
- Lee `signo_folder`, `check_frequency_minutes` y `days_advance` de la configuración.
- Importa archivos Excel nuevos/modificados en la carpeta configurada.
- Ignora archivos temporales de Excel (`~$*.xlsx`).
- Cuenta documentos que requieren recordatorio según `days_advance`.
- Estado disponible en `/health` y vía `get_scheduler_status()`.
- No inicia si `signo_folder` está vacío o no existe.
- Reimportación evitada con huella de archivo (path + mtime).
- Un archivo defectuoso no detiene la revisión del resto; se acumulan errores recientes.

---

## Pruebas automatizadas (28 tests, 28 ✅)

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
| `test_email_validation_is_stricter_than_at_only` | Validación estricta de email |
| `test_reserve_draft_slot_prevents_duplicates_and_allows_retry_after_failure` | Dedupe concurrente y reintento |
| `test_scheduler_ignores_excel_temp_files_and_continues_after_faulty` | Scheduler ignora `~$` y sigue con errores |

---

## Archivo real (20261021133448BWv9Yl1k.xlsx)

Validado en esta sesión sobre base temporal vacía (sin modificar el archivo):

- Filas leídas: **1,018**
- Comprobantes únicos: **978**
- Ventas: **978**
- Pendientes: **28**
- Cancelados: **12**
- Sin saldo cobrable: **950**
- `PENDIENTE`: **27** por **S/ 26,860**
- `SALDO_PARCIAL`: **1** por **S/ 1,320**
- Total cobrable: **S/ 28,180** en **28** comprobantes
- `F|FE01|0030712`: estado `SALDO_PARCIAL`, saldo **1320**, `revision_required = 0`
- Reimportación: no duplica documentos (sigue en **978**)
- Tiempo de importación medido: ~**0.50 s** (segunda importación ~**0.49 s**)

---

## Pendientes para validación en la computadora de la usuaria

### Prioridad 1 — Verificar en Windows con Outlook instalado
1. Activar la integración en Configuración con una cuenta de Outlook real.
2. Crear un borrador desde `/cobranza` y confirmar que aparece en Borradores de Outlook.
3. Verificar error claro cuando la cuenta configurada no exista.
4. Verificar deduplicación persistente tras reinicio de aplicación.

### Prioridad 2 — Scheduler con carpeta real de SIGNO
5. Configurar `signo_folder` real y comprobar importación automática periódica.
6. Validar que los archivos `~$*.xlsx` se ignoran.
7. Confirmar que un archivo defectuoso no detiene la revisión del resto.
8. Verificar en `/configuracion` y `/health`: última revisión, última importación, errores y recordatorios.

### Prioridad 3 — Flujo de contactos y cobranza en UI
9. Registrar y editar contactos en `/cliente/{document_key}` y confirmar persistencia.
10. Confirmar que la lista de `/cobranza` incluye `SALDO_PARCIAL` cobrable y excluye revisión/saldo cero.s totales del dashboard.
