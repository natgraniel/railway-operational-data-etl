"""Tkinter application for generating the daily Programa workbook."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from src.loaders.excel_loader import TestTrainInput
from src.pipeline import run
from src.utils.execution_lock import RunAlreadyInProgressError
from src.utils.shared_workspace import SharedWorkspace
from src.utils.source_date import extract_fleet_forecast_date, extract_operations_report_date

DEFAULT_DESTINATION = Path(
    r"\\tmmertru09\Seguridad_Circulacion\Direccion de Circulacion Ferroviaria\Previsión de Trenes_2026\Agosto_2026"
)


class ProgramaApp(ttk.Frame):
    """Simple operator-facing desktop interface for the ETL pipeline."""

    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=18)
        self.master = master
        self.master.title("Programa ETL")
        self.master.minsize(960, 510)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        self.destination_root = tk.StringVar(value=str(DEFAULT_DESTINATION))
        self.pdf_path = tk.StringVar()
        self.word_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.program_date = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.has_test_train = tk.BooleanVar(value=False)
        self.test_rows: list[tuple[tk.StringVar, tk.StringVar, tk.StringVar, tk.StringVar, tk.StringVar, list[ttk.Entry]]] = []
        self.status = tk.StringVar(value="Selecciona los tres archivos de origen.")
        self._build_form()

    def _build_form(self) -> None:
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text="Generar Programa diario", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 14)
        )
        self._path_row(1, "Carpeta Destino:", self.destination_root, self._select_destination, "Carpeta")
        self._path_row(2, "Previsión de flota (PDF):", self.pdf_path, self._select_pdf, "Examinar")
        self._path_row(3, "Parte de operaciones (Word):", self.word_path, self._select_word, "Examinar")
        self._path_row(4, "Plantilla Programa (Excel):", self.template_path, self._select_template, "Examinar")
        ttk.Label(self, text="Fecha del Programa (dd/mm/aaaa):").grid(row=5, column=0, sticky="w", pady=(10, 4))
        ttk.Entry(self, textvariable=self.program_date, width=18).grid(row=5, column=1, sticky="w", pady=(10, 4))

        self.tests_frame = ttk.LabelFrame(self, text="Pruebas", padding=8)
        self.tests_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        for column in range(5):
            self.tests_frame.columnconfigure(column, weight=1)
        ttk.Checkbutton(
            self.tests_frame,
            text="Hay trenes de pruebas",
            variable=self.has_test_train,
            command=self._toggle_test_fields,
        ).grid(row=0, column=0, columnspan=5, sticky="w")
        for column, label in enumerate(("Tren", "P.K.'s", "MR", "Hora Inicio (UTC-6)", "Hora Final (UTC-6)")):
            ttk.Label(self.tests_frame, text=label).grid(row=1, column=column, sticky="w", padx=4, pady=(6, 0))
        self._add_test_row()
        ttk.Label(
            self.tests_frame,
            text="N/A se agrega automáticamente; Tiempo total se calcula automáticamente.",
        ).grid(row=98, column=0, columnspan=5, sticky="w", pady=(8, 0))
        self.add_test_button = ttk.Button(self.tests_frame, text="+ Agregar otro tren", command=self._add_test_row)
        self.add_test_button.grid(row=99, column=0, columnspan=2, sticky="w", padx=4, pady=(4, 0))
        self.remove_test_button = ttk.Button(self.tests_frame, text="Quitar último", command=self._remove_last_test_row)
        self.remove_test_button.grid(row=99, column=2, columnspan=2, sticky="w", padx=4, pady=(4, 0))
        self._toggle_test_fields()

        ttk.Button(self, text="Generar Programa", command=self._generate).grid(
            row=8, column=0, columnspan=3, sticky="ew", pady=(18, 8)
        )
        ttk.Label(self, textvariable=self.status, wraplength=900).grid(row=9, column=0, columnspan=3, sticky="w")

    def _path_row(self, row: int, label: str, variable: tk.StringVar, command: object, button_text: str) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(self, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=4, padx=(8, 8))
        ttk.Button(self, text=button_text, command=command).grid(row=row, column=2, sticky="ew", pady=4)

    def _select_destination(self) -> None:
        selected = filedialog.askdirectory(title="Selecciona la carpeta destino")
        if selected:
            self.destination_root.set(selected)

    def _select_pdf(self) -> None:
        self._select_file(self.pdf_path, [("PDF", "*.pdf")])

    def _select_word(self) -> None:
        self._select_file(self.word_path, [("Word", "*.docx")])

    def _select_template(self) -> None:
        self._select_file(self.template_path, [("Excel", "*.xlsx")])

    @staticmethod
    def _select_file(variable: tk.StringVar, filetypes: list[tuple[str, str]]) -> None:
        selected = filedialog.askopenfilename(filetypes=filetypes)
        if selected:
            variable.set(selected)

    def _toggle_test_fields(self) -> None:
        state = "normal" if self.has_test_train.get() else "disabled"
        for *_, entries in self.test_rows:
            for entry in entries:
                entry.configure(state=state)
        self.add_test_button.configure(state=state)
        self.remove_test_button.configure(
            state="normal" if self.has_test_train.get() and len(self.test_rows) > 1 else "disabled"
        )

    def _add_test_row(self) -> None:
        row = len(self.test_rows) + 2
        variables = [tk.StringVar() for _ in range(5)]
        entries: list[ttk.Entry] = []
        for column, variable in enumerate(variables):
            entry = ttk.Entry(self.tests_frame, textvariable=variable, width=18)
            entry.grid(row=row, column=column, sticky="ew", padx=4, pady=2)
            entries.append(entry)
        self.test_rows.append((*variables, entries))
        if hasattr(self, "add_test_button"):
            self._toggle_test_fields()

    def _remove_last_test_row(self) -> None:
        if len(self.test_rows) <= 1:
            return
        *_, entries = self.test_rows.pop()
        for entry in entries:
            entry.destroy()
        self._toggle_test_fields()

    def _generate(self) -> None:
        paths = [Path(self.pdf_path.get()), Path(self.word_path.get()), Path(self.template_path.get())]
        if not all(path.is_file() for path in paths):
            messagebox.showerror("Archivos requeridos", "Selecciona un PDF, un Word y una plantilla Excel válidos.")
            return
        try:
            selected_date = datetime.strptime(self.program_date.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            messagebox.showerror("Fecha no válida", "Usa el formato dd/mm/aaaa, por ejemplo 10/07/2026.")
            return

        source_dates = (
            ("Previsión de flota (PDF)", extract_fleet_forecast_date(paths[0])),
            ("Parte de operaciones (Word)", extract_operations_report_date(paths[1])),
        )
        mismatches = [
            (label, source_date)
            for label, source_date in source_dates
            if source_date is not None and source_date != selected_date
        ]
        if mismatches:
            details = "\n".join(
                f"- {label}: {source_date:%d/%m/%Y}"
                for label, source_date in mismatches
            )
            messagebox.showerror(
                "Las fechas no coinciden",
                f"La fecha del Programa es {selected_date:%d/%m/%Y}, pero se detectó:\n\n"
                f"{details}\n\nSelecciona los archivos correspondientes antes de continuar.",
            )
            return

        unidentified = [label for label, source_date in source_dates if source_date is None]
        if unidentified and not messagebox.askyesno(
            "Fecha no identificada",
            "No fue posible identificar la fecha en el nombre de:\n\n"
            + "\n".join(f"- {label}" for label in unidentified)
            + f"\n\nFecha capturada: {selected_date:%d/%m/%Y}.\n"
            "Verifica manualmente los archivos. ¿Deseas continuar?",
        ):
            return

        test_trains = None
        if self.has_test_train.get():
            values = [(train.get().strip(), pks.get().strip(), mr.get().strip(), start.get().strip(), end.get().strip()) for train, pks, mr, start, end, _ in self.test_rows]
            if any(not all(row) for row in values):
                messagebox.showerror("Datos de pruebas", "Completa Tren, P.K.'s, MR, Hora Inicio y Hora Final en cada fila.")
                return
            test_trains = [TestTrainInput(*row) for row in values]

        workspace = SharedWorkspace.from_root(Path(self.destination_root.get()).expanduser())
        workspace.ensure_exists()
        output_path = workspace.output_dir / f"Programa_{selected_date:%Y-%m-%d}.xlsx"
        if output_path.exists() and not messagebox.askyesno("Reemplazar archivo", f"Ya existe {output_path.name}. ¿Deseas reemplazarlo?"):
            return
        self.status.set("Procesando archivos; espera un momento...")
        self.master.update_idletasks()
        try:
            result = run(
                pdf_path=paths[0], word_path=paths[1], workbook_path=paths[2], output_path=output_path,
                program_date=selected_date, test_trains=test_trains,
                report_path=workspace.output_dir / f"Reporte_calidad_{selected_date:%Y-%m-%d}.json",
                log_dir=workspace.logs_dir, lock_dir=workspace.output_dir, archive_dir=workspace.archive_dir,
            )
        except RunAlreadyInProgressError:
            messagebox.showwarning("Proceso en curso", "Otra persona ya está generando un Programa. Intenta nuevamente al terminar.")
            self.status.set("No se generó archivo: hay otro proceso en curso.")
        except PermissionError:
            messagebox.showwarning("Archivo en uso", "No se pudo guardar el Programa porque está abierto en Excel. Ciérralo y vuelve a intentar.")
            self.status.set("No se generó archivo: el Programa está abierto en Excel.")
        except Exception as error:
            messagebox.showerror("No se pudo generar", str(error))
            self.status.set("No se generó archivo. Consulta el reporte de calidad y los registros.")
        else:
            messagebox.showinfo("Programa generado", f"Archivo creado:\n{result.output_path}")
            self.status.set(f"Programa generado correctamente: {result.output_path.name}")


def main() -> None:
    root = tk.Tk()
    ProgramaApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
