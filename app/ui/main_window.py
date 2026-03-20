from __future__ import annotations

import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app.models.options import Action, ProcessingOptions, RoundingMode
from app.services.excel_processor import process_column_percentage
from app.utils.validators import normalize_column, parse_percentage

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

ACTION_LABEL_TO_VALUE = {
    "Aumentar": Action.INCREASE,
    "Diminuir": Action.DECREASE,
}

ROUNDING_LABEL_TO_VALUE = {
    "Nenhum": RoundingMode.NONE,
    "Cima": RoundingMode.UP,
    "Baixo": RoundingMode.DOWN,
}


class MainWindow(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Licita Facil - Ajuste de Colunas Excel")
        self.geometry("760x560")
        self.minsize(720, 520)

        self.input_path_var = ctk.StringVar()
        self.output_path_var = ctk.StringVar()
        self.column_var = ctk.StringVar(value="A")
        self.percentage_var = ctk.StringVar(value="15")
        self.action_var = ctk.StringVar(value="Aumentar")
        self.rounding_var = ctk.StringVar(value="Nenhum")

        self._build_layout()
        self._toggle_rounding(self.action_var.get())

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(self, corner_radius=16, fg_color="#f7f8fa")
        container.grid(row=0, column=0, padx=22, pady=22, sticky="nsew")

        container.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, padx=20, pady=(20, 6), sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)

        title = ctk.CTkLabel(
            header,
            text="Ajuste de Coluna por Porcentagem",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#1f2937",
        )
        title.grid(row=0, column=0, sticky="w")

        author = ctk.CTkLabel(
            header,
            text="Criado por Ronalyson Medeiros",
            text_color="#2563eb",
            cursor="hand2",
            font=ctk.CTkFont(size=14, underline=True),
        )
        author.grid(row=0, column=1, padx=(16, 0), sticky="e")
        author.bind("<Button-1>", lambda _event: webbrowser.open_new_tab("https://ronalyson.dev"))

        subtitle = ctk.CTkLabel(
            container,
            text="Evolução dos scripts VBA para uma interface simples e rapida.",
            font=ctk.CTkFont(size=14),
            text_color="#5f6b7a",
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")

        form = ctk.CTkFrame(container, corner_radius=12, fg_color="#ffffff")
        form.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")
        form.grid_columnconfigure(0, weight=1)
        form.grid_columnconfigure(1, weight=0)

        self._build_file_picker(
            parent=form,
            row=0,
            label="Arquivo de entrada (.xlsx)",
            variable=self.input_path_var,
            button_text="Selecionar",
            callback=self._pick_input,
        )
        self._build_file_picker(
            parent=form,
            row=2,
            label="Salvar como (novo arquivo)",
            variable=self.output_path_var,
            button_text="Salvar em",
            callback=self._pick_output,
        )

        fields = ctk.CTkFrame(form, fg_color="transparent")
        fields.grid(row=4, column=0, columnspan=2, padx=14, pady=(10, 14), sticky="ew")
        fields.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._build_entry(fields, "Coluna", self.column_var, 0)
        self._build_entry(fields, "Porcentagem", self.percentage_var, 1)

        action_label = ctk.CTkLabel(fields, text="Ação", anchor="w")
        action_label.grid(row=0, column=2, padx=8, pady=(6, 2), sticky="w")
        action_segmented = ctk.CTkSegmentedButton(
            fields,
            values=list(ACTION_LABEL_TO_VALUE.keys()),
            variable=self.action_var,
            command=self._toggle_rounding,
            height=36,
        )
        action_segmented.grid(row=1, column=2, padx=8, pady=(0, 6), sticky="ew")
        action_segmented.set("Aumentar")

        rounding_label = ctk.CTkLabel(fields, text="Arredondamento", anchor="w")
        rounding_label.grid(row=0, column=3, padx=8, pady=(6, 2), sticky="w")

        self.rounding_menu = ctk.CTkOptionMenu(
            fields,
            values=list(ROUNDING_LABEL_TO_VALUE.keys()),
            variable=self.rounding_var,
            height=36,
        )
        self.rounding_menu.grid(row=1, column=3, padx=8, pady=(0, 6), sticky="ew")

        labels_help = ctk.CTkLabel(
            form,
            text="Escolha coluna e porcentagem. Arredondamento só vale ao diminuir.",
            text_color="#6b7280",
        )
        labels_help.grid(row=5, column=0, columnspan=2, padx=14, pady=(0, 10), sticky="w")

        self.process_button = ctk.CTkButton(
            container,
            text="Aplicar e Salvar",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=44,
            corner_radius=10,
            command=self._start_processing,
        )
        self.process_button.grid(row=3, column=0, padx=20, pady=(16, 10), sticky="ew")

        self.status_box = ctk.CTkTextbox(container, height=140, corner_radius=10)
        self.status_box.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.status_box.insert("1.0", "Pronto para processar sua planilha.\n")
        self.status_box.configure(state="disabled")


    def _build_file_picker(self, parent, row, label, variable, button_text, callback) -> None:
        label_widget = ctk.CTkLabel(parent, text=label, anchor="w")
        label_widget.grid(row=row, column=0, columnspan=2, padx=14, pady=(14, 4), sticky="w")

        entry = ctk.CTkEntry(parent, textvariable=variable, height=38)
        entry.grid(row=row + 1, column=0, padx=(14, 8), pady=(0, 8), sticky="ew")

        button = ctk.CTkButton(parent, text=button_text, width=120, command=callback)
        button.grid(row=row + 1, column=1, padx=(0, 14), pady=(0, 8))

    def _build_entry(self, parent, label, variable, column) -> None:
        label_widget = ctk.CTkLabel(parent, text=label, anchor="w")
        label_widget.grid(row=0, column=column, padx=8, pady=(6, 2), sticky="w")

        entry = ctk.CTkEntry(parent, textvariable=variable, height=36)
        entry.grid(row=1, column=column, padx=8, pady=(0, 6), sticky="ew")

    def _pick_input(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Escolha o arquivo Excel",
            filetypes=[("Excel", "*.xlsx")],
        )
        if file_path:
            self.input_path_var.set(file_path)
            default_output = str(Path(file_path).with_stem(f"{Path(file_path).stem}-processado"))
            self.output_path_var.set(default_output)

    def _pick_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Salvar novo arquivo",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if file_path:
            self.output_path_var.set(file_path)

    def _toggle_rounding(self, action_value: str) -> None:
        if action_value == "Diminuir":
            self.rounding_menu.configure(state="normal")
        else:
            self.rounding_var.set("Nenhum")
            self.rounding_menu.configure(state="disabled")

    def _start_processing(self) -> None:
        try:
            options = self._collect_options()
        except ValueError as exc:
            messagebox.showerror("Validação", str(exc))
            return

        self.process_button.configure(state="disabled", text="Processando...")
        self._log("Iniciando processamento...")

        thread = threading.Thread(target=self._process, args=(options,), daemon=True)
        thread.start()

    def _collect_options(self) -> ProcessingOptions:
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_path:
            raise ValueError("Selecione o arquivo de entrada.")
        if not output_path:
            raise ValueError("Informe onde salvar o novo arquivo.")
        if Path(input_path).resolve() == Path(output_path).resolve():
            raise ValueError("Salve em um novo arquivo para não sobrescrever o original.")

        action = ACTION_LABEL_TO_VALUE[self.action_var.get()]
        rounding = ROUNDING_LABEL_TO_VALUE[self.rounding_var.get()]

        return ProcessingOptions(
            input_path=input_path,
            output_path=output_path,
            column=normalize_column(self.column_var.get()),
            percentage=parse_percentage(self.percentage_var.get()),
            action=action,
            rounding=rounding,
        )

    def _process(self, options: ProcessingOptions) -> None:
        try:
            result = process_column_percentage(options)
            self.after(0, self._on_success, result, options)
        except Exception as exc:
            self.after(0, self._on_error, exc)

    def _on_success(self, result, options: ProcessingOptions) -> None:
        self._log(f"Planilha ativa: {result.sheet_name}")
        self._log(f"Coluna processada: {options.column}")
        self._log(f"Celulas alteradas: {result.processed_cells}")
        self._log(f"Ignoradas (vazias): {result.skipped_empty}")
        self._log(f"Ignoradas (formulas): {result.skipped_formula}")
        self._log(f"Ignoradas (texto/outros): {result.skipped_non_numeric}")
        self._log(f"Arquivo salvo: {options.output_path}")

        self.process_button.configure(state="normal", text="Aplicar e Salvar")
        messagebox.showinfo("Concluido", "Processamento finalizado com sucesso.")

    def _on_error(self, error: Exception) -> None:
        self.process_button.configure(state="normal", text="Aplicar e Salvar")
        self._log(f"Erro: {error}")
        messagebox.showerror("Erro", str(error))

    def _log(self, message: str) -> None:
        self.status_box.configure(state="normal")
        self.status_box.insert("end", f"{message}\n")
        self.status_box.see("end")
        self.status_box.configure(state="disabled")
