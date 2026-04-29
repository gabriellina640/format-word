from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.config import AppConfig, ConfigStore, FormatSettings
from app.formatter import FormatterError, format_document


class FormatWordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Format Word")
        self.geometry("980x680")
        self.minsize(860, 620)

        self.store = ConfigStore()
        self.config_model = self.store.load()
        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents"))
        self.status_text = tk.StringVar(value="Pronto para formatar.")
        self.settings_summary = tk.StringVar()
        self.last_output_path: Path | None = None

        self._build_theme()
        self._build_layout()
        self._load_settings_into_form(self.config_model.settings)

    def _build_theme(self) -> None:
        self.configure(bg="#eef2f5")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#eef2f5")
        style.configure("Surface.TFrame", background="#ffffff")
        style.configure("Panel.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("Action.TFrame", background="#ffffff")
        style.configure("Toolbar.TFrame", background="#102a43")
        style.configure("TLabel", background="#eef2f5", foreground="#1f2933", font=("Segoe UI", 10))
        style.configure("Panel.TLabel", background="#ffffff", foreground="#1f2933", font=("Segoe UI", 10))
        style.configure("PanelTitle.TLabel", background="#ffffff", foreground="#102a43", font=("Segoe UI", 13, "bold"))
        style.configure("Title.TLabel", background="#102a43", foreground="#ffffff", font=("Segoe UI", 21, "bold"))
        style.configure("HeaderMuted.TLabel", background="#102a43", foreground="#d9e2ec", font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#eef2f5", foreground="#52606d", font=("Segoe UI", 10))
        style.configure("PanelMuted.TLabel", background="#ffffff", foreground="#52606d", font=("Segoe UI", 9))
        style.configure("Status.TLabel", background="#ffffff", foreground="#243b53", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=(13, 8), background="#d9e2ec", foreground="#102a43")
        style.map(
            "TButton",
            background=[("active", "#bcccdc"), ("disabled", "#e5e7eb")],
            foreground=[("active", "#102a43"), ("disabled", "#9aa5b1")],
        )
        style.configure("Accent.TButton", background="#0f766e", foreground="#ffffff", font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#115e59"), ("disabled", "#99bdb7")])
        style.configure("TNotebook", background="#eef2f5", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(18, 10), background="#d9e2ec")
        style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#f0f4f8")])
        style.configure("TCheckbutton", background="#ffffff", foreground="#1f2933", font=("Segoe UI", 10))
        style.configure("TEntry", padding=(8, 6))
        style.configure("TSpinbox", padding=(8, 6))
        style.configure("Horizontal.TProgressbar", background="#0f766e", troughcolor="#d9e2ec", bordercolor="#d9e2ec")

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=22)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root, style="Toolbar.TFrame", padding=(22, 18))
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="Format Word", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Padronização local de documentos DOCX e PDF com preferências persistentes.",
            style="HeaderMuted.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        self.format_tab = ttk.Frame(notebook, padding=18)
        self.settings_tab = ttk.Frame(notebook, padding=18)
        notebook.add(self.format_tab, text="Subir arquivo")
        notebook.add(self.settings_tab, text="Configurações")

        self._build_format_tab()
        self._build_settings_tab()

        footer = ttk.Frame(root, style="Surface.TFrame", padding=(14, 10))
        footer.pack(fill="x", pady=(14, 0))
        ttk.Label(footer, textvariable=self.status_text, style="Status.TLabel").pack(side="left")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=180)
        self.progress.pack(side="right")

    def _build_format_tab(self) -> None:
        container = ttk.Frame(self.format_tab, style="Panel.TFrame", padding=20)
        container.pack(fill="x", anchor="n")

        ttk.Label(container, text="Gerar documento", style="PanelTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        ttk.Label(container, text="Arquivo de entrada", style="Panel.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.input_path).grid(row=2, column=0, sticky="ew", pady=(6, 14))
        ttk.Button(container, text="Selecionar", command=self._choose_input_file).grid(row=2, column=1, padx=(12, 0), pady=(6, 14))

        ttk.Label(container, text="Pasta de saída", style="Panel.TLabel").grid(row=3, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.output_dir).grid(row=4, column=0, sticky="ew", pady=(6, 14))
        ttk.Button(container, text="Escolher", command=self._choose_output_dir).grid(row=4, column=1, padx=(12, 0), pady=(6, 14))

        self.quick_header_var = tk.BooleanVar()
        self.quick_footer_var = tk.BooleanVar()
        ttk.Checkbutton(container, text="Aplicar imagem de cabeçalho salva", variable=self.quick_header_var).grid(row=5, column=0, sticky="w", pady=(4, 2))
        ttk.Checkbutton(container, text="Aplicar imagem de rodapé salva", variable=self.quick_footer_var).grid(row=6, column=0, sticky="w", pady=(2, 14))

        self.format_button = ttk.Button(container, text="Gerar Word formatado", style="Accent.TButton", command=self._start_formatting)
        self.format_button.grid(
            row=7, column=0, sticky="w", pady=(8, 0)
        )
        container.columnconfigure(0, weight=1)

        summary = ttk.Frame(self.format_tab, style="Panel.TFrame", padding=20)
        summary.pack(fill="x", anchor="n", pady=(16, 0))
        ttk.Label(summary, text="Configuração ativa", style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(
            summary,
            textvariable=self.settings_summary,
            style="PanelMuted.TLabel",
            wraplength=820,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

    def _build_settings_tab(self) -> None:
        form = ttk.Frame(self.settings_tab, style="Panel.TFrame", padding=20)
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        self.font_name_var = tk.StringVar()
        self.font_size_var = tk.IntVar()
        self.line_spacing_var = tk.DoubleVar()
        self.space_after_var = tk.IntVar()
        self.indent_var = tk.DoubleVar()
        self.margin_top_var = tk.DoubleVar()
        self.margin_bottom_var = tk.DoubleVar()
        self.margin_left_var = tk.DoubleVar()
        self.margin_right_var = tk.DoubleVar()
        self.justify_var = tk.BooleanVar()
        self.include_header_var = tk.BooleanVar()
        self.include_footer_var = tk.BooleanVar()
        self.output_suffix_var = tk.StringVar()
        self.max_input_mb_var = tk.IntVar()
        self.header_path_var = tk.StringVar()
        self.footer_path_var = tk.StringVar()

        self._label(form, "Fonte", 0, 0)
        ttk.Entry(form, textvariable=self.font_name_var).grid(row=0, column=1, sticky="ew", padx=(10, 18), pady=8)
        self._label(form, "Tamanho", 0, 2)
        ttk.Spinbox(form, from_=8, to=32, textvariable=self.font_size_var, width=8).grid(row=0, column=3, sticky="w", pady=8)

        self._label(form, "Espaçamento linhas", 1, 0)
        ttk.Spinbox(form, from_=1.0, to=3.0, increment=0.1, textvariable=self.line_spacing_var, width=8).grid(row=1, column=1, sticky="w", padx=(10, 18), pady=8)
        self._label(form, "Espaço após parágrafo", 1, 2)
        ttk.Spinbox(form, from_=0, to=36, textvariable=self.space_after_var, width=8).grid(row=1, column=3, sticky="w", pady=8)

        self._label(form, "Recuo primeira linha cm", 2, 0)
        ttk.Spinbox(form, from_=0.0, to=5.0, increment=0.25, textvariable=self.indent_var, width=8).grid(row=2, column=1, sticky="w", padx=(10, 18), pady=8)
        ttk.Checkbutton(form, text="Justificar texto", variable=self.justify_var).grid(row=2, column=2, columnspan=2, sticky="w", pady=8)

        self._label(form, "Margem superior cm", 3, 0)
        ttk.Spinbox(form, from_=0.5, to=6.0, increment=0.25, textvariable=self.margin_top_var, width=8).grid(row=3, column=1, sticky="w", padx=(10, 18), pady=8)
        self._label(form, "Margem inferior cm", 3, 2)
        ttk.Spinbox(form, from_=0.5, to=6.0, increment=0.25, textvariable=self.margin_bottom_var, width=8).grid(row=3, column=3, sticky="w", pady=8)

        self._label(form, "Margem esquerda cm", 4, 0)
        ttk.Spinbox(form, from_=0.5, to=6.0, increment=0.25, textvariable=self.margin_left_var, width=8).grid(row=4, column=1, sticky="w", padx=(10, 18), pady=8)
        self._label(form, "Margem direita cm", 4, 2)
        ttk.Spinbox(form, from_=0.5, to=6.0, increment=0.25, textvariable=self.margin_right_var, width=8).grid(row=4, column=3, sticky="w", pady=8)

        self._label(form, "Sufixo do arquivo", 5, 0)
        ttk.Entry(form, textvariable=self.output_suffix_var).grid(row=5, column=1, sticky="ew", padx=(10, 18), pady=8)
        self._label(form, "Limite MB", 5, 2)
        ttk.Spinbox(form, from_=1, to=300, textvariable=self.max_input_mb_var, width=8).grid(row=5, column=3, sticky="w", pady=8)

        ttk.Separator(form).grid(row=6, column=0, columnspan=4, sticky="ew", pady=18)

        ttk.Checkbutton(form, text="Usar cabeçalho por padrão", variable=self.include_header_var).grid(row=7, column=0, columnspan=2, sticky="w", pady=6)
        ttk.Button(form, text="Enviar cabeçalho", command=self._choose_header_image).grid(row=7, column=2, sticky="w", pady=6)
        ttk.Label(form, textvariable=self.header_path_var, style="PanelMuted.TLabel").grid(row=7, column=3, sticky="ew", pady=6)

        ttk.Checkbutton(form, text="Usar rodapé por padrão", variable=self.include_footer_var).grid(row=8, column=0, columnspan=2, sticky="w", pady=6)
        ttk.Button(form, text="Enviar rodapé", command=self._choose_footer_image).grid(row=8, column=2, sticky="w", pady=6)
        ttk.Label(form, textvariable=self.footer_path_var, style="PanelMuted.TLabel").grid(row=8, column=3, sticky="ew", pady=6)

        actions = ttk.Frame(form, style="Action.TFrame")
        actions.grid(row=9, column=0, columnspan=4, sticky="ew", pady=(22, 0))
        ttk.Button(actions, text="Salvar configurações", style="Accent.TButton", command=self._save_settings).pack(side="left")
        ttk.Button(actions, text="Restaurar padrão", command=self._reset_settings).pack(side="left", padx=(12, 0))

    def _label(self, parent: ttk.Frame, text: str, row: int, column: int) -> None:
        ttk.Label(parent, text=text, style="Panel.TLabel").grid(row=row, column=column, sticky="w", pady=8)

    def _choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Documentos", "*.docx *.pdf"), ("Word", "*.docx"), ("PDF", "*.pdf")])
        if file_path:
            self.input_path.set(file_path)

    def _choose_output_dir(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir.set(folder)

    def _choose_header_image(self) -> None:
        self._choose_image("header")

    def _choose_footer_image(self) -> None:
        self._choose_image("footer")

    def _choose_image(self, image_type: str) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if not file_path:
            return
        try:
            stored = self.store.store_image(Path(file_path), image_type)
        except ValueError as exc:
            messagebox.showerror("Imagem inválida", str(exc))
            return

        if image_type == "header":
            self.header_path_var.set(stored)
        else:
            self.footer_path_var.set(stored)
        self._update_settings_summary(self._collect_settings())
        self.status_text.set("Imagem salva nas configurações locais.")

    def _load_settings_into_form(self, settings: FormatSettings) -> None:
        self.font_name_var.set(settings.font_name)
        self.font_size_var.set(settings.font_size)
        self.line_spacing_var.set(settings.line_spacing)
        self.space_after_var.set(settings.paragraph_spacing_after)
        self.indent_var.set(settings.first_line_indent_cm)
        self.margin_top_var.set(settings.margin_top_cm)
        self.margin_bottom_var.set(settings.margin_bottom_cm)
        self.margin_left_var.set(settings.margin_left_cm)
        self.margin_right_var.set(settings.margin_right_cm)
        self.justify_var.set(settings.justify_text)
        self.include_header_var.set(settings.include_header)
        self.include_footer_var.set(settings.include_footer)
        self.quick_header_var.set(settings.include_header)
        self.quick_footer_var.set(settings.include_footer)
        self.output_suffix_var.set(settings.output_suffix)
        self.max_input_mb_var.set(settings.max_input_mb)
        self.header_path_var.set(settings.header_image_path)
        self.footer_path_var.set(settings.footer_image_path)
        self._update_settings_summary(settings)

    def _collect_settings(self) -> FormatSettings:
        return FormatSettings(
            font_name=self.font_name_var.get().strip() or "Arial",
            font_size=self._int_from_var(self.font_size_var, 12, 8, 32),
            line_spacing=self._float_from_var(self.line_spacing_var, 1.5, 1.0, 3.0),
            paragraph_spacing_after=self._int_from_var(self.space_after_var, 6, 0, 36),
            first_line_indent_cm=self._float_from_var(self.indent_var, 1.25, 0.0, 5.0),
            margin_top_cm=self._float_from_var(self.margin_top_var, 3.0, 0.5, 6.0),
            margin_bottom_cm=self._float_from_var(self.margin_bottom_var, 2.0, 0.5, 6.0),
            margin_left_cm=self._float_from_var(self.margin_left_var, 3.0, 0.5, 6.0),
            margin_right_cm=self._float_from_var(self.margin_right_var, 2.0, 0.5, 6.0),
            justify_text=bool(self.justify_var.get()),
            include_header=bool(self.include_header_var.get()),
            include_footer=bool(self.include_footer_var.get()),
            header_image_path=self.header_path_var.get(),
            footer_image_path=self.footer_path_var.get(),
            output_suffix=self.output_suffix_var.get().strip() or "_formatado",
            max_input_mb=self._int_from_var(self.max_input_mb_var, 50, 1, 300),
        )

    def _save_settings(self) -> None:
        settings = self._collect_settings()
        self.config_model = AppConfig(settings=settings)
        self.store.save(self.config_model)
        self.quick_header_var.set(settings.include_header)
        self.quick_footer_var.set(settings.include_footer)
        self._update_settings_summary(settings)
        self.status_text.set("Configurações salvas.")
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")

    def _reset_settings(self) -> None:
        self._load_settings_into_form(FormatSettings())
        self.status_text.set("Configurações restauradas para o padrão. Salve para tornar permanente.")

    def _start_formatting(self) -> None:
        settings = self._collect_settings()
        settings.include_header = bool(self.quick_header_var.get())
        settings.include_footer = bool(self.quick_footer_var.get())
        input_path = Path(self.input_path.get())
        output_dir = Path(self.output_dir.get())

        self._set_busy(True)
        self.status_text.set("Formatando documento...")
        thread = threading.Thread(target=self._format_worker, args=(input_path, output_dir, settings), daemon=True)
        thread.start()

    def _format_worker(self, input_path: Path, output_dir: Path, settings: FormatSettings) -> None:
        try:
            result = format_document(input_path, output_dir, settings)
        except FormatterError as exc:
            self.after(0, self._show_error, str(exc))
        except Exception as exc:
            self.after(0, self._show_error, f"Erro inesperado: {exc}")
        else:
            self.after(0, self._show_success, str(result.output_path), result.paragraphs)

    def _show_error(self, message: str) -> None:
        self._set_busy(False)
        self.status_text.set("Falha ao formatar.")
        messagebox.showerror("Não foi possível formatar", message)

    def _show_success(self, output_path: str, paragraphs: int) -> None:
        self._set_busy(False)
        self.last_output_path = Path(output_path)
        self.status_text.set(f"Arquivo gerado: {output_path}")
        messagebox.showinfo("Documento gerado", f"Arquivo gerado com {paragraphs} parágrafos:\n{output_path}")

    def _set_busy(self, is_busy: bool) -> None:
        state = "disabled" if is_busy else "normal"
        self.format_button.configure(state=state)
        if is_busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _update_settings_summary(self, settings: FormatSettings) -> None:
        alignment = "justificado" if settings.justify_text else "alinhado à esquerda"
        header = "cabeçalho ativo" if settings.include_header else "sem cabeçalho"
        footer = "rodapé ativo" if settings.include_footer else "sem rodapé"
        self.settings_summary.set(
            f"{settings.font_name} {settings.font_size} pt, texto {alignment}, "
            f"linhas {settings.line_spacing:.1f}, recuo {settings.first_line_indent_cm:.2f} cm, "
            f"margens {settings.margin_top_cm:.2f}/{settings.margin_right_cm:.2f}/"
            f"{settings.margin_bottom_cm:.2f}/{settings.margin_left_cm:.2f} cm, {header}, {footer}."
        )

    @staticmethod
    def _int_from_var(var: tk.Variable, default: int, minimum: int, maximum: int) -> int:
        try:
            value = int(var.get())
        except (tk.TclError, ValueError):
            value = default
        return max(minimum, min(maximum, value))

    @staticmethod
    def _float_from_var(var: tk.Variable, default: float, minimum: float, maximum: float) -> float:
        try:
            value = float(var.get())
        except (tk.TclError, ValueError):
            value = default
        return max(minimum, min(maximum, value))
