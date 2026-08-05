from __future__ import annotations

import threading
import traceback
from datetime import datetime
from pathlib import Path
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import actualizar_creditos_vista_operativa_con_correos as core
import correos_outlook
from openpyxl import load_workbook


class CreditosApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Actualización de Créditos desde SIGNO")
        self.geometry("780x480")
        self.minsize(720, 450)

        self.signo_var = tk.StringVar()
        self.creditos_var = tk.StringVar()
        self.salida_var = tk.StringVar()
        self.estado_var = tk.StringVar(value="Seleccione los archivos para comenzar.")

        self._crear_interfaz()

    def _crear_interfaz(self) -> None:
        cont = ttk.Frame(self, padding=22)
        cont.pack(fill="both", expand=True)

        ttk.Label(
            cont,
            text="Conciliación de créditos",
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            cont,
            text="Seleccione el reporte de SIGNO y el archivo operativo de Créditos actuales.",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(2, 18))

        self._fila_archivo(
            cont,
            "1. Reporte exportado de SIGNO",
            self.signo_var,
            self._seleccionar_signo,
        )
        self._fila_archivo(
            cont,
            "2. Archivo Créditos actuales",
            self.creditos_var,
            self._seleccionar_creditos,
        )
        self._fila_archivo(
            cont,
            "3. Archivo de salida",
            self.salida_var,
            self._seleccionar_salida,
            guardar=True,
        )

        ttk.Separator(cont).pack(fill="x", pady=18)

        self.progreso = ttk.Progressbar(cont, mode="indeterminate")
        self.progreso.pack(fill="x")
        ttk.Label(cont, textvariable=self.estado_var, wraplength=660).pack(
            anchor="w", pady=(8, 14)
        )

        acciones = ttk.Frame(cont)
        acciones.pack(fill="x")
        ttk.Button(acciones, text="Cerrar", command=self.destroy).pack(side="right")
        self.boton_correos = ttk.Button(
            acciones,
            text="Generar borradores Outlook",
            command=self._iniciar_correos,
        )
        self.boton_correos.pack(side="right", padx=(0, 10))
        self.boton_ejecutar = ttk.Button(
            acciones,
            text="Actualizar créditos",
            command=self._iniciar_proceso,
        )
        self.boton_ejecutar.pack(side="right", padx=(0, 10))

    def _fila_archivo(self, parent, texto, variable, comando, guardar=False) -> None:
        bloque = ttk.Frame(parent)
        bloque.pack(fill="x", pady=6)
        ttk.Label(bloque, text=texto, width=28).pack(side="left")
        ttk.Entry(bloque, textvariable=variable).pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(
            bloque,
            text="Guardar como..." if guardar else "Examinar...",
            command=comando,
            width=15,
        ).pack(side="right")

    def _seleccionar_signo(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccione el reporte de SIGNO",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.signo_var.set(ruta)
            self._sugerir_salida()

    def _seleccionar_creditos(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccione Créditos actuales",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self.creditos_var.set(ruta)
            self._sugerir_salida()

    def _seleccionar_salida(self) -> None:
        inicial = self.salida_var.get() or f"CREDITOS_ACTUALIZADOS_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        ruta = filedialog.asksaveasfilename(
            title="Guardar archivo actualizado",
            defaultextension=".xlsx",
            initialfile=Path(inicial).name,
            initialdir=str(Path(inicial).parent) if Path(inicial).parent.exists() else None,
            filetypes=[("Archivos Excel", "*.xlsx")],
        )
        if ruta:
            self.salida_var.set(ruta)

    def _sugerir_salida(self) -> None:
        if self.salida_var.get():
            return
        base = self.creditos_var.get() or self.signo_var.get()
        if base:
            carpeta = Path(base).parent
            self.salida_var.set(
                str(carpeta / f"CREDITOS_ACTUALIZADOS_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
            )

    def _iniciar_proceso(self) -> None:
        signo = Path(self.signo_var.get().strip())
        creditos = Path(self.creditos_var.get().strip())
        salida = Path(self.salida_var.get().strip())

        if not signo.is_file():
            messagebox.showerror("Archivo faltante", "Seleccione un reporte válido de SIGNO.")
            return
        if not creditos.is_file():
            messagebox.showerror("Archivo faltante", "Seleccione un archivo válido de Créditos actuales.")
            return
        if not salida.name:
            messagebox.showerror("Salida faltante", "Indique dónde guardar el resultado.")
            return
        if salida.suffix.lower() != ".xlsx":
            salida = salida.with_suffix(".xlsx")
            self.salida_var.set(str(salida))
        if salida.resolve() in (signo.resolve(), creditos.resolve()):
            messagebox.showerror(
                "Ruta no permitida",
                "El archivo de salida no puede sobrescribir ninguno de los archivos originales.",
            )
            return

        self.boton_ejecutar.config(state="disabled")
        self.boton_correos.config(state="disabled")
        self.progreso.start(12)
        self.estado_var.set("Procesando archivos. No cierre esta ventana...")
        threading.Thread(
            target=self._ejecutar,
            args=(signo, creditos, salida),
            daemon=True,
        ).start()

    def _ejecutar(self, signo_path: Path, creditos_path: Path, output: Path) -> None:
        try:
            signo = core.read_signo(signo_path)
            old, section_order, _section_rank, client_order = core.read_current_view(creditos_path)
            merged, changes, exceptions = core.reconcile(
                signo, old, section_order, client_order
            )

            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(creditos_path, output)
            wb = load_workbook(output)
            core.write_operational_view(
                core.replace_or_create(wb, "CREDITOS"),
                merged,
                section_order,
                client_order,
            )
            core.write_summary(core.replace_or_create(wb, "RESUMEN"), merged)
            core.write_review(
                core.replace_or_create(wb, "REVISAR_CAMBIOS"), changes, exceptions
            )
            core.write_base(core.replace_or_create(wb, "BASE_DATOS"), merged)
            core.write_reminders(core.replace_or_create(wb, "RECORDATORIOS"), merged)
            core.ensure_history(wb)
            correos_outlook.ensure_email_sheets(wb)
            correos_outlook.sync_master_clients(wb, merged)
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.save(output)
            wb.close()

            nuevos = sum(1 for x in changes if x[0] == "NUEVO")
            pagos = sum(1 for x in changes if x[0] == "POSIBLE PAGO/ANULACIÓN")
            saldos = sum(1 for x in changes if x[0] == "SALDO MODIFICADO")
            resumen = (
                f"Proceso terminado.\n\n"
                f"Comprobantes en SIGNO: {len(signo)}\n"
                f"Registros anteriores: {len(old)}\n"
                f"Nuevos: {nuevos}\n"
                f"Saldos modificados: {saldos}\n"
                f"Posibles pagos/anulaciones: {pagos}\n\n"
                f"Archivo generado:\n{output}"
            )
            self.after(0, self._finalizar_exito, resumen, output)
        except Exception as exc:
            detalle = f"{exc}\n\n{traceback.format_exc()}"
            self.after(0, self._finalizar_error, detalle)


    def _archivo_para_correos(self) -> Path | None:
        salida = Path(self.salida_var.get().strip()) if self.salida_var.get().strip() else None
        creditos = Path(self.creditos_var.get().strip()) if self.creditos_var.get().strip() else None
        if salida and salida.is_file():
            return salida
        if creditos and creditos.is_file():
            return creditos
        ruta = filedialog.askopenfilename(
            title="Seleccione el archivo actualizado para generar borradores",
            filetypes=[("Archivos Excel", "*.xlsx")],
        )
        return Path(ruta) if ruta else None

    def _iniciar_correos(self) -> None:
        archivo = self._archivo_para_correos()
        if not archivo:
            return
        respuesta = messagebox.askyesno(
            "Generar borradores",
            "Se crearán borradores en Outlook usando la hoja MAESTRO_CLIENTES. "
            "No se enviará ningún correo automáticamente.\n\n"
            "Cierre el archivo de Excel antes de continuar. ¿Desea proceder?",
        )
        if not respuesta:
            return
        self.boton_ejecutar.config(state="disabled")
        self.boton_correos.config(state="disabled")
        self.progreso.start(12)
        self.estado_var.set("Generando borradores en Outlook...")
        threading.Thread(target=self._ejecutar_correos, args=(archivo,), daemon=True).start()

    def _ejecutar_correos(self, archivo: Path) -> None:
        try:
            result = correos_outlook.generate_outlook_drafts(archivo)
            lines = [
                "Proceso de borradores terminado.",
                "",
                f"Borradores creados: {result['created']}",
                f"Omitidos por duplicado del día: {result['duplicates']}",
                f"Clientes sin correo: {len(result['missing_email'])}",
                f"Clientes bloqueados o desactivados: {len(result['blocked'])}",
                f"Errores: {len(result['errors'])}",
            ]
            if result['missing_email']:
                lines += ["", "Sin correo (primeros 10):"] + result['missing_email'][:10]
            if result['errors']:
                lines += ["", "Errores (primeros 5):"] + result['errors'][:5]
            self.after(0, self._finalizar_correos, "\n".join(lines))
        except Exception as exc:
            detalle = f"{exc}\n\n{traceback.format_exc()}"
            self.after(0, self._finalizar_error, detalle)

    def _finalizar_correos(self, resumen: str) -> None:
        self.progreso.stop()
        self.boton_ejecutar.config(state="normal")
        self.boton_correos.config(state="normal")
        self.estado_var.set("Borradores creados en Outlook. Revise la carpeta Borradores.")
        messagebox.showinfo("Borradores de Outlook", resumen)

    def _finalizar_exito(self, resumen: str, output: Path) -> None:
        self.progreso.stop()
        self.boton_ejecutar.config(state="normal")
        self.boton_correos.config(state="normal")
        self.estado_var.set(f"Proceso terminado: {output.name}")
        messagebox.showinfo("Actualización completada", resumen)

    def _finalizar_error(self, detalle: str) -> None:
        self.progreso.stop()
        self.boton_ejecutar.config(state="normal")
        self.boton_correos.config(state="normal")
        self.estado_var.set("No se pudo completar el proceso.")
        messagebox.showerror("Error durante la actualización", detalle)


if __name__ == "__main__":
    app = CreditosApp()
    app.mainloop()
