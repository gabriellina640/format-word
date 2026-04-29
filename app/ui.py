from __future__ import annotations

import threading
import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Callable

import customtkinter as ctk
from PIL import Image, ImageTk

from app.config import ConfigStore, FONT_OPTIONS, FormatSettings
from app.formatter import FormatterError, format_paragraphs, read_document_paragraphs, validate_input


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BG = "#f3f6fb"
SURFACE = "#ffffff"
INK = "#101828"
MUTED = "#667085"
PRIMARY = "#2563eb"
SUCCESS = "#047857"


class FormatWordApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Format Word")
        self.geometry("1180x760")
        self.minsize(980, 680)
        self.configure(fg_color=BG)

        self.store = ConfigStore()
        self.config_model = self.store.load()
        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents"))
        self.selected_stack_var = tk.StringVar(value="Configuração atual")
        self.manage_stack_var = tk.StringVar(value="Selecione uma stack")
        self.stack_name_var = tk.StringVar()
        self.status_text = tk.StringVar(value="Pronto para formatar.")
        self.settings_summary = tk.StringVar()
        self.header_status = tk.StringVar(value="Nenhuma imagem carregada")
        self.footer_status = tk.StringVar(value="Nenhuma imagem carregada")
        self.header_preview_image: ctk.CTkImage | None = None
        self.footer_preview_image: ctk.CTkImage | None = None
        self.last_output_path: Path | None = None

        self._build_layout()
        initial_settings = self.config_model.stacks.get(self.config_model.active_stack, self.config_model.settings)
        initial_name = self.config_model.active_stack or "Configuração atual"
        self.selected_stack_var.set(initial_name)
        self.manage_stack_var.set(self.config_model.active_stack or "Selecione uma stack")
        self.stack_name_var.set(self.config_model.active_stack)
        self._load_settings_into_form(initial_settings)
        self._refresh_stack_dropdowns()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="#111827", corner_radius=18)
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 16))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Format Word", text_color="#ffffff", font=("Segoe UI", 28, "bold")).grid(
            row=0, column=0, sticky="w", padx=24, pady=(20, 2)
        )
        ctk.CTkLabel(
            header,
            text="Padronize documentos com prévia editável, cabeçalho, rodapé e preferências persistentes.",
            text_color="#d1d5db",
            font=("Segoe UI", 13),
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 20))

        self.tabs = ctk.CTkTabview(self, fg_color=BG, segmented_button_fg_color="#e5e7eb")
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=24)
        self.upload_tab = self.tabs.add("Documento")
        self.settings_tab = self.tabs.add("Configurações")
        self.upload_tab.configure(fg_color=BG)
        self.settings_tab.configure(fg_color=BG)
        self.upload_tab.grid_columnconfigure(0, weight=1)
        self.upload_tab.grid_rowconfigure(0, weight=1)
        self.settings_tab.grid_columnconfigure(0, weight=1)
        self.settings_tab.grid_rowconfigure(0, weight=1)

        self._build_upload_tab()
        self._build_settings_tab()

        footer = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=14, border_width=1, border_color="#e5e7eb")
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(16, 22))
        footer.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(footer, textvariable=self.status_text, text_color=INK, font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=12
        )
        self.progress = ctk.CTkProgressBar(footer, mode="indeterminate", width=190, progress_color=PRIMARY)
        self.progress.grid(row=0, column=1, sticky="e", padx=16, pady=12)
        self.progress.set(0)

    def _build_upload_tab(self) -> None:
        content = ctk.CTkScrollableFrame(self.upload_tab, fg_color=BG)
        content.grid(row=0, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)

        card = self._card(content)
        card.grid(row=0, column=0, sticky="ew", pady=(6, 16))
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text="Preparar documento", text_color=INK, font=("Segoe UI", 19, "bold")).grid(
            row=0, column=0, sticky="w", padx=22, pady=(20, 4)
        )
        ctk.CTkLabel(
            card,
            text="Selecione o arquivo, revise em uma prévia editável e exporte somente quando estiver pronto.",
            text_color=MUTED,
            font=("Segoe UI", 13),
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 18))

        self._path_picker(card, "Arquivo de entrada", self.input_path, self._choose_input_file, 2)
        self._path_picker(card, "Pasta de saída", self.output_dir, self._choose_output_dir, 4)

        stack_row = ctk.CTkFrame(card, fg_color="transparent")
        stack_row.grid(row=6, column=0, sticky="ew", padx=22, pady=(0, 16))
        stack_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(stack_row, text="Stack de configuração", text_color=INK, font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        self.stack_dropdown = ctk.CTkOptionMenu(
            stack_row,
            values=self._stack_options(),
            variable=self.selected_stack_var,
            command=self._select_stack,
            height=40,
            corner_radius=12,
        )
        self.stack_dropdown.grid(row=1, column=0, sticky="ew")

        switches = ctk.CTkFrame(card, fg_color="transparent")
        switches.grid(row=7, column=0, sticky="w", padx=22, pady=(4, 14))
        self.quick_header_var = tk.BooleanVar()
        self.quick_footer_var = tk.BooleanVar()
        ctk.CTkCheckBox(switches, text="Aplicar cabeçalho salvo", variable=self.quick_header_var).grid(
            row=0, column=0, sticky="w", padx=(0, 18)
        )
        ctk.CTkCheckBox(switches, text="Aplicar rodapé salvo", variable=self.quick_footer_var).grid(row=0, column=1, sticky="w")

        self.format_button = ctk.CTkButton(
            card,
            text="Pré-visualizar e exportar",
            command=self._start_formatting,
            height=42,
            corner_radius=12,
            fg_color=PRIMARY,
            hover_color="#1d4ed8",
            font=("Segoe UI", 14, "bold"),
        )
        self.format_button.grid(row=8, column=0, sticky="w", padx=22, pady=(0, 22))

        summary = self._card(content)
        summary.grid(row=1, column=0, sticky="ew")
        summary.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(summary, text="Configuração ativa", text_color=INK, font=("Segoe UI", 17, "bold")).grid(
            row=0, column=0, sticky="w", padx=22, pady=(18, 4)
        )
        ctk.CTkLabel(
            summary,
            textvariable=self.settings_summary,
            text_color=MUTED,
            font=("Segoe UI", 13),
            wraplength=940,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 18))

    def _build_settings_tab(self) -> None:
        content = ctk.CTkScrollableFrame(self.settings_tab, fg_color=BG)
        content.grid(row=0, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)

        card = self._card(content)
        card.grid(row=0, column=0, sticky="nsew", pady=(6, 0))
        card.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(card, text="Preferências de formatação", text_color=INK, font=("Segoe UI", 19, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=22, pady=(20, 16)
        )

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
        self.header_offset_x_var = tk.DoubleVar()
        self.header_offset_y_var = tk.DoubleVar()
        self.footer_offset_x_var = tk.DoubleVar()
        self.footer_offset_y_var = tk.DoubleVar()

        self._field(card, "Fonte", ctk.CTkComboBox(card, values=list(FONT_OPTIONS), variable=self.font_name_var), 1, 0)
        self._field(card, "Tamanho", ctk.CTkEntry(card, textvariable=self.font_size_var), 1, 1)
        self._field(card, "Espaçamento entre linhas", ctk.CTkEntry(card, textvariable=self.line_spacing_var), 1, 2)
        self._field(card, "Espaço após parágrafo", ctk.CTkEntry(card, textvariable=self.space_after_var), 1, 3)
        self._field(card, "Recuo primeira linha cm", ctk.CTkEntry(card, textvariable=self.indent_var), 3, 0)
        self._field(card, "Margem superior cm", ctk.CTkEntry(card, textvariable=self.margin_top_var), 3, 1)
        self._field(card, "Margem inferior cm", ctk.CTkEntry(card, textvariable=self.margin_bottom_var), 3, 2)
        self._field(card, "Limite do arquivo MB", ctk.CTkEntry(card, textvariable=self.max_input_mb_var), 3, 3)
        self._field(card, "Margem esquerda cm", ctk.CTkEntry(card, textvariable=self.margin_left_var), 5, 0)
        self._field(card, "Margem direita cm", ctk.CTkEntry(card, textvariable=self.margin_right_var), 5, 1)
        self._field(card, "Sufixo do arquivo", ctk.CTkEntry(card, textvariable=self.output_suffix_var), 5, 2)

        ctk.CTkCheckBox(card, text="Justificar texto", variable=self.justify_var).grid(
            row=6, column=0, sticky="w", padx=22, pady=(6, 18)
        )

        media = ctk.CTkFrame(card, fg_color="#f8fafc", corner_radius=16, border_width=1, border_color="#e5e7eb")
        media.grid(row=7, column=0, columnspan=4, sticky="ew", padx=22, pady=(4, 16))
        media.grid_columnconfigure((0, 1), weight=1)

        self._image_panel(media, "Cabeçalho", self.include_header_var, self.header_status, self._choose_header_image, "header", 0)
        self._image_panel(media, "Rodapé", self.include_footer_var, self.footer_status, self._choose_footer_image, "footer", 1)

        stack_box = ctk.CTkFrame(card, fg_color="#f8fafc", corner_radius=16, border_width=1, border_color="#e5e7eb")
        stack_box.grid(row=8, column=0, columnspan=4, sticky="ew", padx=22, pady=(0, 16))
        stack_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(stack_box, text="Stacks salvas", text_color=INK, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 4)
        )
        ctk.CTkLabel(
            stack_box,
            text="Selecione uma stack para editar/excluir ou informe um novo nome para criar outra.",
            text_color=MUTED,
            font=("Segoe UI", 12),
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 10))
        self.manage_stack_dropdown = ctk.CTkOptionMenu(
            stack_box,
            values=self._manage_stack_options(),
            variable=self.manage_stack_var,
            command=self._select_stack_for_edit,
            height=38,
            corner_radius=11,
        )
        self.manage_stack_dropdown.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 10))
        stack_actions = ctk.CTkFrame(stack_box, fg_color="transparent")
        stack_actions.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))
        stack_actions.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(
            stack_actions,
            textvariable=self.stack_name_var,
            placeholder_text="Ex: Documentos para PGJ",
            height=38,
            corner_radius=11,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ctk.CTkButton(stack_actions, text="Salvar stack", command=self._save_stack, height=38, corner_radius=11).grid(
            row=0, column=1, padx=(0, 10)
        )
        ctk.CTkButton(
            stack_actions,
            text="Excluir stack",
            command=self._delete_stack,
            height=38,
            corner_radius=11,
            fg_color="#fee2e2",
            hover_color="#fecaca",
            text_color="#991b1b",
        ).grid(row=0, column=2)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=9, column=0, columnspan=4, sticky="ew", padx=22, pady=(0, 22))
        ctk.CTkButton(actions, text="Salvar configurações", command=self._save_settings, height=40, corner_radius=12).pack(side="left")
        ctk.CTkButton(
            actions,
            text="Restaurar padrão",
            command=self._reset_settings,
            height=40,
            corner_radius=12,
            fg_color="#e5e7eb",
            hover_color="#d1d5db",
            text_color=INK,
        ).pack(side="left", padx=(10, 0))

    def _card(self, parent) -> ctk.CTkFrame:
        return ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=18, border_width=1, border_color="#e5e7eb")

    def _path_picker(self, parent, label: str, variable: tk.StringVar, command: Callable[[], None], row: int) -> None:
        ctk.CTkLabel(parent, text=label, text_color=INK, font=("Segoe UI", 13, "bold")).grid(
            row=row, column=0, sticky="w", padx=22, pady=(0, 6)
        )
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.grid(row=row + 1, column=0, sticky="ew", padx=22, pady=(0, 16))
        line.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(line, textvariable=variable, height=40, corner_radius=12).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ctk.CTkButton(line, text="Selecionar", command=command, width=120, height=40, corner_radius=12).grid(row=0, column=1)

    def _field(self, parent, label: str, widget, row: int, column: int) -> None:
        ctk.CTkLabel(parent, text=label, text_color=INK, font=("Segoe UI", 12, "bold")).grid(
            row=row, column=column, sticky="w", padx=22, pady=(0, 6)
        )
        widget.configure(height=38, corner_radius=11)
        widget.grid(row=row + 1, column=column, sticky="ew", padx=22, pady=(0, 14))

    def _stack_options(self) -> list[str]:
        return ["Configuração atual", *sorted(self.config_model.stacks)]

    def _manage_stack_options(self) -> list[str]:
        return ["Selecione uma stack", *sorted(self.config_model.stacks)]

    def _refresh_stack_dropdowns(self) -> None:
        if hasattr(self, "stack_dropdown"):
            self.stack_dropdown.configure(values=self._stack_options())
        if hasattr(self, "manage_stack_dropdown"):
            self.manage_stack_dropdown.configure(values=self._manage_stack_options())

    def _select_stack(self, stack_name: str) -> None:
        if stack_name == "Configuração atual":
            settings = self.config_model.settings
            self.config_model.active_stack = ""
            self.stack_name_var.set("")
            self.manage_stack_var.set("Selecione uma stack")
        else:
            settings = self.config_model.stacks.get(stack_name)
            if settings is None:
                return
            self.config_model.active_stack = stack_name
            self.stack_name_var.set(stack_name)
            self.manage_stack_var.set(stack_name)

        self._load_settings_into_form(settings)
        self.store.save(self.config_model)
        self.status_text.set(f"Stack selecionada: {stack_name}.")

    def _select_stack_for_edit(self, stack_name: str) -> None:
        if stack_name == "Selecione uma stack":
            self.stack_name_var.set("")
            return
        self.selected_stack_var.set(stack_name)
        self._select_stack(stack_name)

    def _save_stack(self) -> None:
        stack_name = self.stack_name_var.get().strip()
        if not stack_name:
            messagebox.showerror("Nome obrigatório", "Informe um nome para salvar a stack.")
            return

        settings = self._collect_settings()
        self.config_model.settings = settings
        self.config_model.stacks[stack_name] = settings
        self.config_model.active_stack = stack_name
        self.store.save(self.config_model)
        self.selected_stack_var.set(stack_name)
        self.manage_stack_var.set(stack_name)
        self._refresh_stack_dropdowns()
        self._update_settings_summary(settings)
        self.status_text.set(f"Stack '{stack_name}' salva.")
        messagebox.showinfo("Stack salva", f"A stack '{stack_name}' foi salva com sucesso.")

    def _delete_stack(self) -> None:
        selected_for_management = self.manage_stack_var.get()
        stack_name = selected_for_management if selected_for_management != "Selecione uma stack" else self.stack_name_var.get().strip()
        if not stack_name or stack_name == "Configuração atual" or stack_name not in self.config_model.stacks:
            messagebox.showerror("Stack não encontrada", "Selecione uma stack salva para excluir.")
            return
        if not messagebox.askyesno("Excluir stack", f"Excluir a stack '{stack_name}'?"):
            return

        del self.config_model.stacks[stack_name]
        self.config_model.active_stack = ""
        self.selected_stack_var.set("Configuração atual")
        self.manage_stack_var.set("Selecione uma stack")
        self.stack_name_var.set("")
        self._load_settings_into_form(self.config_model.settings)
        self.store.save(self.config_model)
        self._refresh_stack_dropdowns()
        self.status_text.set(f"Stack '{stack_name}' excluída.")

    def _image_panel(
        self,
        parent,
        title: str,
        enabled_var: tk.BooleanVar,
        status_var: tk.StringVar,
        command: Callable[[], None],
        image_type: str,
        column: int,
    ) -> None:
        panel = ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=14, border_width=1, border_color="#e5e7eb")
        panel.grid(row=0, column=column, sticky="nsew", padx=12, pady=12)
        panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(panel, text=title, text_color=INK, font=("Segoe UI", 15, "bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 4)
        )
        ctk.CTkCheckBox(panel, text=f"Usar {title.lower()}", variable=enabled_var).grid(
            row=1, column=0, sticky="w", padx=16, pady=(0, 8)
        )
        ctk.CTkButton(panel, text=f"Importar {title.lower()}", command=command, height=36, corner_radius=11).grid(
            row=2, column=0, sticky="w", padx=16, pady=(0, 8)
        )
        ctk.CTkLabel(panel, textvariable=status_var, text_color=SUCCESS, font=("Segoe UI", 12, "bold")).grid(
            row=3, column=0, sticky="w", padx=16, pady=(0, 8)
        )
        preview = ctk.CTkLabel(panel, text="", fg_color="#f8fafc", corner_radius=10, width=330, height=60)
        preview.grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 16))
        if image_type == "header":
            self.header_preview_label = preview
        else:
            self.footer_preview_label = preview

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
            self.include_header_var.set(True)
            self.quick_header_var.set(True)
        else:
            self.footer_path_var.set(stored)
            self.include_footer_var.set(True)
            self.quick_footer_var.set(True)
        self._refresh_image_feedback(image_type)
        self._update_settings_summary(self._collect_settings())
        label = "Cabeçalho" if image_type == "header" else "Rodapé"
        self.status_text.set(f"{label} importado e adaptado ao tamanho profissional.")

    def _refresh_image_feedback(self, image_type: str) -> None:
        path_var = self.header_path_var if image_type == "header" else self.footer_path_var
        status_var = self.header_status if image_type == "header" else self.footer_status
        preview_label = self.header_preview_label if image_type == "header" else self.footer_preview_label
        image_path = Path(path_var.get()) if path_var.get() else Path()
        if not image_path.is_file():
            status_var.set("Nenhuma imagem carregada")
            preview_label.configure(image=None, text="")
            return

        status_var.set(f"Imagem pronta: {image_path.name}")
        try:
            preview = Image.open(image_path).convert("RGBA")
            preview.thumbnail((330, 60), Image.Resampling.LANCZOS)
            photo = ctk.CTkImage(light_image=preview, dark_image=preview, size=preview.size)
        except OSError:
            preview_label.configure(image=None, text="Prévia indisponível")
            return

        if image_type == "header":
            self.header_preview_image = photo
        else:
            self.footer_preview_image = photo
        preview_label.configure(image=photo, text="")

    def _load_settings_into_form(self, settings: FormatSettings) -> None:
        self.font_name_var.set(settings.font_name if settings.font_name in FONT_OPTIONS else "Arial")
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
        self.header_offset_x_var.set(settings.header_offset_x_cm)
        self.header_offset_y_var.set(settings.header_offset_y_cm)
        self.footer_offset_x_var.set(settings.footer_offset_x_cm)
        self.footer_offset_y_var.set(settings.footer_offset_y_cm)
        self._refresh_image_feedback("header")
        self._refresh_image_feedback("footer")
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
            header_offset_x_cm=self._float_from_var(self.header_offset_x_var, 0.0, -4.0, 4.0),
            header_offset_y_cm=self._float_from_var(self.header_offset_y_var, 0.0, -0.5, 1.5),
            footer_offset_x_cm=self._float_from_var(self.footer_offset_x_var, 0.0, -4.0, 4.0),
            footer_offset_y_cm=self._float_from_var(self.footer_offset_y_var, 0.0, -0.5, 1.5),
            output_suffix=self.output_suffix_var.get().strip() or "_formatado",
            max_input_mb=self._int_from_var(self.max_input_mb_var, 50, 1, 300),
        )

    def _save_settings(self) -> None:
        settings = self._collect_settings()
        self.config_model.settings = settings
        selected_stack = self.selected_stack_var.get()
        if selected_stack in self.config_model.stacks:
            self.config_model.stacks[selected_stack] = settings
            self.config_model.active_stack = selected_stack
        self.store.save(self.config_model)
        self.quick_header_var.set(settings.include_header)
        self.quick_footer_var.set(settings.include_footer)
        self._refresh_stack_dropdowns()
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
        self.status_text.set("Preparando pré-visualização...")
        thread = threading.Thread(target=self._preview_worker, args=(input_path, output_dir, settings), daemon=True)
        thread.start()

    def _preview_worker(self, input_path: Path, output_dir: Path, settings: FormatSettings) -> None:
        try:
            validate_input(input_path.expanduser().resolve(), settings.max_input_mb)
            paragraphs = read_document_paragraphs(input_path)
            if not paragraphs:
                raise FormatterError("Não foi possível extrair texto do arquivo informado.")
        except FormatterError as exc:
            self.after(0, self._show_error, str(exc))
        except Exception as exc:
            self.after(0, self._show_error, f"Erro inesperado: {exc}")
        else:
            self.after(0, self._open_preview_modal, paragraphs, input_path.stem, output_dir, settings)

    def _open_preview_modal(self, paragraphs: list[str], input_stem: str, output_dir: Path, settings: FormatSettings) -> None:
        self._set_busy(False)
        self.status_text.set("Pré-visualização pronta.")
        PreviewDraftModal(self, paragraphs, input_stem, output_dir, settings, self._show_success)

    def _show_error(self, message: str) -> None:
        self._set_busy(False)
        self.status_text.set("Falha ao preparar documento.")
        messagebox.showerror("Não foi possível continuar", message)

    def _show_success(self, output_path: str, paragraphs: int) -> None:
        self._set_busy(False)
        self.last_output_path = Path(output_path)
        self.status_text.set(f"Arquivo gerado: {output_path}")
        messagebox.showinfo("Documento gerado", f"Arquivo gerado com {paragraphs} parágrafos:\n{output_path}")

    def _set_busy(self, is_busy: bool) -> None:
        self.format_button.configure(state="disabled" if is_busy else "normal")
        if is_busy:
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.set(0)

    def _update_settings_summary(self, settings: FormatSettings) -> None:
        alignment = "justificado" if settings.justify_text else "alinhado à esquerda"
        header = "cabeçalho ativo" if settings.include_header else "sem cabeçalho"
        footer = "rodapé ativo" if settings.include_footer else "sem rodapé"
        self.settings_summary.set(
            f"{settings.font_name} {settings.font_size} pt, texto {alignment}, "
            f"linhas {settings.line_spacing:.1f}, recuo {settings.first_line_indent_cm:.2f} cm, "
            f"margens {settings.margin_top_cm:.2f}/{settings.margin_right_cm:.2f}/"
            f"{settings.margin_bottom_cm:.2f}/{settings.margin_left_cm:.2f} cm, {header}, {footer}. "
            "Imagens são adaptadas automaticamente a uma área profissional fixa."
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


class PreviewDraftModal(ctk.CTkToplevel):
    def __init__(
        self,
        parent: FormatWordApp,
        paragraphs: list[str],
        input_stem: str,
        output_dir: Path,
        settings: FormatSettings,
        on_export: Callable[[str, int], None],
    ) -> None:
        super().__init__(parent)
        self.parent = parent
        self.input_stem = input_stem
        self.output_dir = output_dir
        self.settings = settings
        self.on_export = on_export
        self.preview_images: list[ImageTk.PhotoImage] = []

        self.title("Pré-visualização")
        self.geometry("1220x780")
        self.minsize(920, 640)
        self.configure(fg_color=BG)
        self.transient(parent)

        self.font_name_var = tk.StringVar(value=settings.font_name)
        self.font_size_var = tk.IntVar(value=settings.font_size)
        self.line_spacing_var = tk.DoubleVar(value=settings.line_spacing)
        self.space_after_var = tk.IntVar(value=settings.paragraph_spacing_after)
        self.justify_var = tk.BooleanVar(value=settings.justify_text)
        self.header_x_var = tk.DoubleVar(value=settings.header_offset_x_cm)
        self.header_y_var = tk.DoubleVar(value=settings.header_offset_y_cm)
        self.footer_x_var = tk.DoubleVar(value=settings.footer_offset_x_cm)
        self.footer_y_var = tk.DoubleVar(value=settings.footer_offset_y_cm)

        self._build(paragraphs)
        self._redraw_preview()
        self.grab_set()

    def _build(self, paragraphs: list[str]) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="#111827", corner_radius=18)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(20, 14))
        ctk.CTkLabel(header, text="Revisar antes de exportar", text_color="#ffffff", font=("Segoe UI", 24, "bold")).pack(
            anchor="w", padx=22, pady=(18, 2)
        )
        ctk.CTkLabel(
            header,
            text="Ajuste texto, espaçamento e posição das imagens. O Word só será criado ao confirmar.",
            text_color="#d1d5db",
            font=("Segoe UI", 13),
        ).pack(anchor="w", padx=22, pady=(0, 18))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=22)
        self.body = body
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.preview_card = ctk.CTkFrame(body, fg_color=SURFACE, corner_radius=18, border_width=1, border_color="#e5e7eb")
        self.preview_card.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        ctk.CTkLabel(self.preview_card, text="Prévia visual", text_color=INK, font=("Segoe UI", 17, "bold")).pack(
            anchor="w", padx=18, pady=(18, 10)
        )
        self.canvas = tk.Canvas(self.preview_card, width=390, height=550, bg="#eef2f7", highlightthickness=0)
        self.canvas.pack(padx=18, pady=(0, 18))

        self.editor = ctk.CTkFrame(body, fg_color=SURFACE, corner_radius=18, border_width=1, border_color="#e5e7eb")
        self.editor.grid(row=0, column=1, sticky="nsew")
        self.editor.grid_columnconfigure(0, weight=1)
        self.editor.grid_rowconfigure(3, weight=1)

        controls = ctk.CTkFrame(self.editor, fg_color="#f8fafc", corner_radius=16)
        controls.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 12))
        controls.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self._modal_field(controls, "Fonte", ctk.CTkComboBox(controls, values=list(FONT_OPTIONS), variable=self.font_name_var), 0, 0)
        self._modal_field(controls, "Tamanho", ctk.CTkEntry(controls, textvariable=self.font_size_var), 0, 1)
        self._modal_field(controls, "Linhas", ctk.CTkEntry(controls, textvariable=self.line_spacing_var), 0, 2)
        self._modal_field(controls, "Após parágrafo", ctk.CTkEntry(controls, textvariable=self.space_after_var), 0, 3)
        ctk.CTkCheckBox(controls, text="Justificar", variable=self.justify_var, command=self._redraw_preview).grid(
            row=1, column=4, sticky="w", padx=12, pady=(0, 14)
        )

        sliders = ctk.CTkFrame(self.editor, fg_color="#f8fafc", corner_radius=16)
        sliders.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        sliders.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self._slider(sliders, "Cabeçalho X", self.header_x_var, 0, -4.0, 4.0)
        self._slider(sliders, "Cabeçalho Y", self.header_y_var, 1, -0.5, 1.5)
        self._slider(sliders, "Rodapé X", self.footer_x_var, 2, -4.0, 4.0)
        self._slider(sliders, "Rodapé Y", self.footer_y_var, 3, -0.5, 1.5)

        ctk.CTkLabel(self.editor, text="Texto do documento", text_color=INK, font=("Segoe UI", 17, "bold")).grid(
            row=2, column=0, sticky="w", padx=18, pady=(0, 8)
        )
        text_frame = ctk.CTkFrame(self.editor, fg_color="#f8fafc", corner_radius=16)
        text_frame.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 18))
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        self.text = tk.Text(
            text_frame,
            wrap="word",
            undo=True,
            font=("Segoe UI", 11),
            relief="flat",
            padx=14,
            pady=14,
            bg="#f8fafc",
            fg=INK,
            insertbackground=INK,
        )
        self.text.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        scrollbar = ctk.CTkScrollbar(text_frame, command=self.text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=8)
        self.text.configure(yscrollcommand=scrollbar.set)
        self.text.insert("1.0", "\n\n".join(paragraphs))
        self.text.bind("<KeyRelease>", lambda _event: self._redraw_preview())

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=22, pady=(14, 20))
        ctk.CTkButton(
            actions,
            text="Cancelar",
            command=self.destroy,
            height=40,
            corner_radius=12,
            fg_color="#e5e7eb",
            hover_color="#d1d5db",
            text_color=INK,
        ).pack(side="right")
        ctk.CTkButton(actions, text="Exportar Word", command=self._export, height=40, corner_radius=12).pack(
            side="right", padx=(0, 10)
        )

        for var in (
            self.font_name_var,
            self.font_size_var,
            self.line_spacing_var,
            self.space_after_var,
            self.header_x_var,
            self.header_y_var,
            self.footer_x_var,
            self.footer_y_var,
        ):
            var.trace_add("write", lambda *_args: self._redraw_preview())
        self.bind("<Configure>", self._apply_responsive_layout)

    def _apply_responsive_layout(self, _event=None) -> None:
        if not hasattr(self, "preview_card") or not hasattr(self, "editor"):
            return
        try:
            if not self.winfo_exists() or not self.body.winfo_exists():
                return
            if self.winfo_width() < 1080:
                self.body.grid_columnconfigure(0, weight=1)
                self.body.grid_columnconfigure(1, weight=0)
                self.body.grid_rowconfigure(0, weight=0)
                self.body.grid_rowconfigure(1, weight=1)
                self.preview_card.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 14))
                self.editor.grid(row=1, column=0, sticky="nsew")
                return
            self.body.grid_columnconfigure(0, weight=0)
            self.body.grid_columnconfigure(1, weight=1)
            self.body.grid_rowconfigure(0, weight=1)
            self.body.grid_rowconfigure(1, weight=0)
            self.preview_card.grid(row=0, column=0, sticky="ns", padx=(0, 14), pady=0)
            self.editor.grid(row=0, column=1, sticky="nsew")
        except tk.TclError:
            return

    def destroy(self) -> None:
        try:
            self.unbind("<Configure>")
        except tk.TclError:
            pass
        super().destroy()

    def _modal_field(self, parent, label: str, widget, row: int, column: int) -> None:
        ctk.CTkLabel(parent, text=label, text_color=INK, font=("Segoe UI", 12, "bold")).grid(
            row=row, column=column, sticky="w", padx=12, pady=(12, 6)
        )
        widget.configure(height=36, corner_radius=10)
        widget.grid(row=row + 1, column=column, sticky="ew", padx=12, pady=(0, 14))

    def _slider(self, parent, label: str, variable: tk.DoubleVar, column: int, minimum: float, maximum: float) -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=0, column=column, sticky="ew", padx=12, pady=12)
        ctk.CTkLabel(frame, text=label, text_color=INK, font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ctk.CTkSlider(frame, from_=minimum, to=maximum, variable=variable, command=lambda _value: self._redraw_preview()).pack(
            fill="x", pady=(8, 0)
        )

    def _current_settings(self) -> FormatSettings:
        settings = replace(self.settings)
        settings.font_name = self.font_name_var.get()
        settings.font_size = self._int_value(self.font_size_var, settings.font_size, 8, 32)
        settings.line_spacing = self._float_value(self.line_spacing_var, settings.line_spacing, 1.0, 3.0)
        settings.paragraph_spacing_after = self._int_value(self.space_after_var, settings.paragraph_spacing_after, 0, 36)
        settings.justify_text = bool(self.justify_var.get())
        settings.header_offset_x_cm = self._float_value(self.header_x_var, 0.0, -4.0, 4.0)
        settings.header_offset_y_cm = self._float_value(self.header_y_var, 0.0, -0.5, 1.5)
        settings.footer_offset_x_cm = self._float_value(self.footer_x_var, 0.0, -4.0, 4.0)
        settings.footer_offset_y_cm = self._float_value(self.footer_y_var, 0.0, -0.5, 1.5)
        return settings

    def _draft_paragraphs(self) -> list[str]:
        raw = self.text.get("1.0", "end").strip()
        return [" ".join(block.split()) for block in split_paragraphs(raw)]

    def _redraw_preview(self) -> None:
        if not hasattr(self, "canvas"):
            return
        settings = self._current_settings()
        self.canvas.delete("all")
        self.preview_images.clear()

        x0, y0, width, height = 30, 18, 325, 500
        self.canvas.create_rectangle(x0 + 7, y0 + 8, x0 + width + 7, y0 + height + 8, fill="#cbd5e1", outline="")
        self.canvas.create_rectangle(x0, y0, x0 + width, y0 + height, fill="#ffffff", outline="#d0d5dd")

        if settings.include_header:
            self._draw_image(settings.header_image_path, x0, y0 + 28 + settings.header_offset_y_cm * 16, width, 42, settings.header_offset_x_cm)
        else:
            self.canvas.create_text(x0 + width / 2, y0 + 45, text="Sem cabeçalho", fill="#98a2b3", font=("Segoe UI", 8))

        text_y = y0 + 98
        font_size = max(7, min(12, int(settings.font_size * 0.75)))
        for line in self._preview_lines()[:10]:
            self.canvas.create_text(
                x0 + 34,
                text_y,
                text=line[:78],
                anchor="w",
                fill="#1f2937",
                font=(settings.font_name, font_size),
            )
            text_y += 15 + int((settings.line_spacing - 1.0) * 7)

        footer_y = y0 + height - 62 + settings.footer_offset_y_cm * 16
        if settings.include_footer:
            self._draw_image(settings.footer_image_path, x0, footer_y, width, 34, settings.footer_offset_x_cm)
        else:
            self.canvas.create_text(x0 + width / 2, y0 + height - 38, text="Sem rodapé", fill="#98a2b3", font=("Segoe UI", 8))

    def _draw_image(self, path_value: str, page_x: int, y: float, page_width: int, max_height: int, offset_x_cm: float) -> None:
        image_path = Path(path_value) if path_value else Path()
        if not image_path.is_file():
            self.canvas.create_text(page_x + page_width / 2, y + 15, text="Imagem não encontrada", fill="#b42318", font=("Segoe UI", 8))
            return
        try:
            image = Image.open(image_path).convert("RGBA")
            image.thumbnail((page_width - 60, max_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
        except OSError:
            self.canvas.create_text(page_x + page_width / 2, y + 15, text="Prévia indisponível", fill="#b42318", font=("Segoe UI", 8))
            return

        self.preview_images.append(photo)
        x = page_x + page_width / 2 - image.width / 2 + offset_x_cm * 18
        self.canvas.create_image(x, y, image=photo, anchor="nw")
        self.canvas.create_rectangle(x, y, x + image.width, y + image.height, outline="#93c5fd")

    def _preview_lines(self) -> list[str]:
        lines: list[str] = []
        for paragraph in self._draft_paragraphs():
            words = paragraph.split()
            current = ""
            for word in words:
                if len(current) + len(word) > 64:
                    lines.append(current)
                    current = word
                else:
                    current = f"{current} {word}".strip()
            if current:
                lines.append(current)
        return lines or ["Digite o texto do documento aqui."]

    def _export(self) -> None:
        try:
            result = format_paragraphs(self._draft_paragraphs(), self.output_dir, self.input_stem, self._current_settings())
        except FormatterError as exc:
            messagebox.showerror("Não foi possível exportar", str(exc), parent=self)
            return
        except Exception as exc:
            messagebox.showerror("Não foi possível exportar", f"Erro inesperado: {exc}", parent=self)
            return
        self.destroy()
        self.on_export(str(result.output_path), result.paragraphs)

    @staticmethod
    def _int_value(var: tk.Variable, default: int, minimum: int, maximum: int) -> int:
        try:
            value = int(var.get())
        except (tk.TclError, ValueError):
            value = default
        return max(minimum, min(maximum, value))

    @staticmethod
    def _float_value(var: tk.Variable, default: float, minimum: float, maximum: float) -> float:
        try:
            value = float(var.get())
        except (tk.TclError, ValueError):
            value = default
        return max(minimum, min(maximum, value))


def split_paragraphs(text: str) -> list[str]:
    blocks: list[str] = []
    current: list[str] = []
    for line in text.splitlines():
        if line.strip():
            current.append(line.strip())
            continue
        if current:
            blocks.append(" ".join(current))
            current = []
    if current:
        blocks.append(" ".join(current))
    return blocks
