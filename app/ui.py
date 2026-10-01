"""Desktop form and batch coordinator. Tk is only accessed by its main thread."""
from __future__ import annotations

import os
from copy import deepcopy
from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, font, messagebox, ttk
from dataclasses import replace

import customtkinter as ctk

from app.batch import BatchItem, format_documents_batch
from app.config import AppConfig, ConfigStore, FONT_OPTIONS, FormatSettings, SettingsError
from app.document_model import DocumentOverrides, inspect_document
from app.review_dialogs import CategoryRulesDialog, DocumentReviewDialog
from app.ui_fields import FIELDS, FIELD_MAP, form_values, settings_from_values, settings_summary, unique_profile_name
from app.word_controls import SpecialIndentControl, spacing_shortcuts

DEFAULT_PROFILE = 'Configuração atual'
BG, SURFACE, INK, MUTED, PRIMARY = '#f3f6fb', '#ffffff', '#101828', '#475467', '#2563eb'
PROFILE_SECTIONS = ('Fonte', 'Parágrafo', 'Página', 'Cabeçalho e rodapé', 'Imagens', 'Opções avançadas')


def field_section(field):
    if field.name in ('keep_with_next', 'keep_together', 'widow_control') or field.group == 'Saída':
        return 'Opções avançadas'
    return {'Texto': 'Fonte', 'Cabeçalho': 'Cabeçalho e rodapé', 'Rodapé': 'Cabeçalho e rodapé'}.get(field.group, field.group)

ctk.set_appearance_mode('light')
ctk.set_default_color_theme('blue')


class FormatWordApp(ctk.CTk):
    def __init__(self, config_dir: Path | None = None) -> None:
        super().__init__()
        self.title('Format Word')
        if sys.platform == 'win32':
            try:
                self.iconbitmap(str(Path(__file__).resolve().parent.parent / 'icone.ico'))
            except tk.TclError:
                pass
        self.geometry('1000x760')
        self.minsize(760, 600)
        self.configure(fg_color=BG)
        self.store = ConfigStore(config_dir)
        self.config_model = self.store.load()
        self.busy = False
        self._destroyed = False
        self._close_pending = False
        self._loading = False
        self._after_ids: set[str] = set()
        self._summary_after = None
        self._form_after = None
        self._form_columns = 0
        self._compact_layout = False
        self._selected = self.config_model.active_stack
        self._draft_base = deepcopy(self.config_model.stacks.get(self._selected, self.config_model.settings))
        self._baseline_values = form_values(self._draft_base)
        self._baseline_base = deepcopy(self._draft_base)
        self._baseline_name = self._selected
        self.field_vars = {name: tk.StringVar(self, value=value) for name, value in self._baseline_values.items()}
        self.field_widgets = {}
        self.field_labels = {}
        self.field_errors = {}
        self.field_cards = []
        self.image_buttons = {}
        self.mutation_widgets = []
        self.paths: list[Path] = []
        self.document_overrides: dict[Path, DocumentOverrides] = {}
        self.results: list[BatchItem] = []
        self.cancel_event = threading.Event()
        self.events: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self._importing = False
        self.import_dialog = None
        self.selected_profile = tk.StringVar(self, value=self._selected or DEFAULT_PROFILE)
        self.profile_name = tk.StringVar(self, value=self._selected)
        self.profile_section = tk.StringVar(self, value=PROFILE_SECTIONS[0])
        self.section_hint = tk.StringVar(self)
        self.output_dir = tk.StringVar(self, value='')
        self.status_text = tk.StringVar(self, value='Selecione arquivos Word e uma pasta de destino.')
        self.error_text = tk.StringVar(self)
        self.dirty_text = tk.StringVar(self)
        self._build()
        for variable in (*self.field_vars.values(), self.profile_name):
            variable.trace_add('write', self._changed)
        self.bind('<Configure>', self._resize)
        for modifier in ('Control', 'Command') if sys.platform == 'darwin' else ('Control',):
            self.bind(f'<{modifier}-s>', lambda _event: self.save_profile())
            self.bind(f'<{modifier}-Return>', lambda _event: self.start_batch())
            self.bind(f'<{modifier}-1>', lambda _event: self.tabs.set('Arquivos e resultados'))
            self.bind(f'<{modifier}-2>', lambda _event: self.tabs.set('Perfis e formatação'))
        self.bind('<Escape>', lambda _event: self.cancel_batch())
        self.protocol('WM_DELETE_WINDOW', self.request_close)
        self._refresh_form()
        self._update_summary()
        self._schedule(80, self._poll)

    @property
    def dirty(self) -> bool:
        return (self._values() != self._baseline_values or self._draft_base != self._baseline_base
                or self.profile_name.get().strip() != self._baseline_name)

    def _schedule(self, milliseconds, callback):
        def run():
            self._after_ids.discard(identifier)
            if not self._destroyed:
                callback()
        identifier = self.after(milliseconds, run)
        self._after_ids.add(identifier)
        return identifier

    def _button(self, parent, text, command, **kwargs):
        widget = ctk.CTkButton(parent, text=text, command=command, fg_color=PRIMARY, height=36, **kwargs)
        self._enable_keyboard(widget)
        self.mutation_widgets.append(widget)
        return widget

    def _enable_keyboard(self, button):
        # CustomTkinter buttons draw on a Canvas, whose default tab focus is off.
        def activate(_event):
            button.invoke()
            return 'break'

        def apply_shortcut(_event):
            self.start_batch()
            return 'break'

        button._canvas.configure(takefocus=1)
        button.bind('<Return>', activate)
        button.bind('<space>', activate)
        for modifier in ('Control', 'Command') if sys.platform == 'darwin' else ('Control',):
            button.bind(f'<{modifier}-Return>', apply_shortcut)
        button.bind('<FocusIn>', lambda _event: button.configure(border_width=2, border_color=INK))
        button.bind('<FocusOut>', lambda _event: button.configure(border_width=0))

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        top = ctk.CTkFrame(self, fg_color=SURFACE)
        top.grid(row=0, column=0, sticky='ew', padx=16, pady=(16, 8))
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(top, text='Format Word', font=ctk.CTkFont(size=24, weight='bold'), text_color=INK).grid(row=0, column=0, padx=16, pady=12)
        self.profile_selector = ctk.CTkComboBox(top, variable=self.selected_profile, values=self._profile_names(), state='readonly', command=self.select_profile, width=230)
        self.profile_selector.grid(row=0, column=1, sticky='e', padx=12)
        self.mutation_widgets.append(self.profile_selector)
        ctk.CTkLabel(top, textvariable=self.dirty_text, text_color=MUTED).grid(row=1, column=0, columnspan=2, sticky='w', padx=16)
        self.tabs = ctk.CTkTabview(self, fg_color=SURFACE)
        self.tabs.grid(row=1, column=0, sticky='nsew', padx=16, pady=4)
        documents = self.tabs.add('Arquivos e resultados')
        profiles = self.tabs.add('Perfis e formatação')
        self._build_documents(documents)
        self._build_profiles(profiles)
        self.status_label = ctk.CTkLabel(self, textvariable=self.status_text, text_color=MUTED, anchor='w', wraplength=900)
        self.status_label.grid(row=2, column=0, sticky='ew', padx=20, pady=(4, 12))

    def _build_documents(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)
        actions = ctk.CTkFrame(parent, fg_color='transparent')
        actions.grid(row=0, column=0, sticky='ew', pady=4)
        for column, (label, command) in enumerate((('Adicionar DOCX', self.choose_files), ('Remover selecionados', self.remove_selected), ('Limpar lista', self.clear_files))):
            self._button(actions, label, command, width=145).grid(row=0, column=column, padx=(0, 8))
        self._button(actions, 'Revisar documento', self.review_document, width=145).grid(row=0, column=3)
        frame = ctk.CTkFrame(parent, fg_color=SURFACE)
        frame.grid(row=1, column=0, sticky='nsew', pady=6)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        self.file_tree = ttk.Treeview(frame, columns=('file', 'status'), show='headings', selectmode='extended', height=5)
        self.file_tree.heading('file', text='Arquivo Word')
        self.file_tree.heading('status', text='Resultado')
        self.file_tree.column('file', width=440, minwidth=160)
        self.file_tree.column('status', width=190, minwidth=120)
        self.file_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.file_tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        self.file_tree.bind('<<TreeviewSelect>>', self._show_result)
        destination = ctk.CTkFrame(parent, fg_color='transparent')
        destination.grid(row=2, column=0, sticky='ew', pady=4)
        destination.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(destination, text='Pasta de destino').grid(row=0, column=0, padx=(0, 8))
        entry = ctk.CTkEntry(destination, textvariable=self.output_dir)
        entry.grid(row=0, column=1, sticky='ew')
        self.mutation_widgets.append(entry)
        self._button(destination, 'Escolher', self.choose_output, width=90).grid(row=0, column=2, padx=(8, 0))
        ctk.CTkLabel(parent, text='Configuração que será aplicada', text_color=INK, anchor='w').grid(row=3, column=0, sticky='ew')
        self.summary_box = ctk.CTkTextbox(parent, height=105, wrap='word')
        self.summary_box.grid(row=4, column=0, sticky='ew')
        run = ctk.CTkFrame(parent, fg_color='transparent')
        run.grid(row=5, column=0, sticky='ew', pady=8)
        run.grid_columnconfigure(2, weight=1)
        self.apply_button = self._button(run, 'Aplicar aos arquivos', self.start_batch, width=170)
        self.apply_button.grid(row=0, column=0, padx=(0, 8))
        self.cancel_button = ctk.CTkButton(run, text='Cancelar lote', command=self.cancel_batch, state='disabled', width=120, height=36)
        self._enable_keyboard(self.cancel_button)
        self.cancel_button.grid(row=0, column=1, padx=(0, 12))
        self.progress = ctk.CTkProgressBar(run)
        self.progress.grid(row=0, column=2, sticky='ew')
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(run, text='0 / 0', width=65)
        self.progress_label.grid(row=0, column=3)
        self.details_box = ctk.CTkTextbox(parent, height=65, wrap='word')
        self.details_box.grid(row=6, column=0, sticky='ew')
        self._set_text(self.details_box, '\n'.join(self.store.warnings) or 'Selecione um resultado para ver os detalhes.')
        opening = ctk.CTkFrame(parent, fg_color='transparent')
        opening.grid(row=7, column=0, sticky='ew', pady=(8, 0))
        for label, command in (('Abrir DOCX selecionado', self.open_selected), ('Abrir pasta de saída', self.open_output)):
            button = ctk.CTkButton(opening, text=label, command=command, height=36)
            self._enable_keyboard(button)
            button.pack(side='left', padx=(0, 8))

    def _build_profiles(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(4, weight=1)
        controls = ctk.CTkFrame(parent, fg_color='transparent')
        controls.grid(row=0, column=0, sticky='ew', pady=4)
        controls.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(controls, text='Nome do perfil').grid(row=0, column=0, padx=(0, 8))
        self.name_entry = ctk.CTkEntry(controls, textvariable=self.profile_name, placeholder_text='Ex.: Promotor João — manifestações')
        self.name_entry.grid(row=0, column=1, sticky='ew')
        self.mutation_widgets.append(self.name_entry)
        buttons = ctk.CTkFrame(parent, fg_color='transparent')
        buttons.grid(row=1, column=0, sticky='ew', pady=4)
        for column, (text, command) in enumerate((('Novo perfil', self.new_profile), ('Salvar', self.save_profile), ('Duplicar', self.duplicate_profile), ('Excluir', self.delete_profile), ('Restaurar', self.restore_profile))):
            self._button(buttons, text, command, width=95).grid(row=0, column=column, padx=(0, 8))
        self.import_button = self._button(buttons, 'Importar de Word', self.import_profile, width=145)
        self.import_button.grid(row=1, column=0, columnspan=2, sticky='w', pady=(8, 0))
        ctk.CTkLabel(buttons, text='1. Importe um Word ou crie um perfil.\n2. Confira as seções abaixo.  3. Salve para reutilizar.',
            text_color=MUTED, justify='left', anchor='w').grid(row=1, column=2, columnspan=3, sticky='w', padx=8, pady=(8, 0))
        notice = ctk.CTkFrame(parent, fg_color='transparent')
        notice.grid(row=2, column=0, sticky='ew')
        notice.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(notice, textvariable=self.error_text, text_color='#b42318', wraplength=340, anchor='w', justify='left').grid(row=0, column=0, sticky='ew')
        self.review_button = self._button(notice, 'Revisar perfil antigo', self.review_legacy, width=155)
        self.review_button.grid(row=0, column=1, padx=6)
        self._button(notice, 'Tipos de texto…', self.edit_category_rules, width=135).grid(row=0, column=2)
        navigation = ctk.CTkFrame(parent, fg_color='transparent')
        navigation.grid(row=3, column=0, sticky='ew', pady=(8, 0))
        navigation.grid_columnconfigure(0, weight=1)
        self.section_selector = ctk.CTkSegmentedButton(navigation, values=list(PROFILE_SECTIONS),
            variable=self.profile_section, command=self.show_profile_section, height=32)
        self.section_selector.grid(row=0, column=0, sticky='ew')
        for button in self.section_selector._buttons_dict.values():
            self._enable_keyboard(button)
        ctk.CTkLabel(navigation, textvariable=self.section_hint, text_color=MUTED,
            anchor='w').grid(row=1, column=0, sticky='ew', pady=3)
        self.form = ctk.CTkScrollableFrame(parent, fg_color=BG)
        self.form.grid(row=4, column=0, sticky='nsew', pady=6)
        try:
            families = sorted(set(font.families(self)) | set(FONT_OPTIONS))
        except tk.TclError:
            families = list(FONT_OPTIONS)
        self.group_labels = {}
        for field in FIELDS:
            if field.group not in self.group_labels:
                self.group_labels[field.group] = ctk.CTkLabel(self.form, text=field.group, text_color=INK, font=ctk.CTkFont(size=18, weight='bold'), anchor='w')
            card = ctk.CTkFrame(self.form, fg_color=SURFACE)
            card.grid_columnconfigure(0, weight=1)
            label = ctk.CTkLabel(card, text=field.label, anchor='w', text_color=INK)
            label.grid(row=0, column=0, sticky='ew', padx=10, pady=(6, 0))
            self.field_labels[field.name] = label
            if field.name == 'first_line_indent_cm':
                widget = SpecialIndentControl(card, self.field_vars[field.name])
            elif field.choices or field.kind == 'font':
                widget = ctk.CTkComboBox(card, variable=self.field_vars[field.name], values=list(field.choices) if field.choices else families, state='readonly' if field.choices else 'normal')
            else:
                widget = ctk.CTkEntry(card, textvariable=self.field_vars[field.name], state='disabled' if field.kind == 'image' else 'normal')
            widget.grid(row=1, column=0, sticky='ew', padx=10, pady=5)
            self.field_widgets[field.name] = widget
            for focus_target in getattr(widget, 'focus_targets', (widget,)):
                focus_target.bind('<FocusIn>', lambda _event, name=field.name, target=focus_target: self._scroll_to_field(name, target), add='+')
            if field.name == 'line_spacing_mode':
                shortcuts, self.spacing_buttons = spacing_shortcuts(card, self.field_vars)
                shortcuts.grid(row=2, column=0, sticky='ew', padx=10, pady=4)
                for button in self.spacing_buttons:
                    button.bind('<FocusIn>', lambda _event, target=button: self._scroll_to_field('line_spacing_mode', target), add='+')
            if field.name == 'font_color':
                colors = ctk.CTkFrame(card, fg_color='transparent')
                colors.grid(row=2, column=0, sticky='ew', padx=10, pady=4)
                self.color_buttons = []
                for text, command in (('Escolher cor…', self.choose_font_color),
                                      ('Manter original', lambda: self.field_vars['font_color'].set(''))):
                    button = self._button(colors, text, command, width=125)
                    button.pack(side='left', padx=(0, 8))
                    button.bind('<FocusIn>', lambda _event, target=button: self._scroll_to_field('font_color', target), add='+')
                    self.color_buttons.append(button)
            if field.kind == 'image':
                slot = field.name.split('_')[0]
                button = self._button(card, 'Escolher imagem PNG/JPEG', lambda slot=slot: self.choose_image(slot))
                button.grid(row=2, column=0, sticky='ew', padx=10, pady=4)
                self.image_buttons[slot] = button
                button.bind('<FocusIn>', lambda _event, name=field.name, target=button: self._scroll_to_field(name, target), add='+')
            if field.help:
                ctk.CTkLabel(card, text=field.help, text_color=MUTED, wraplength=330, justify='left', anchor='w').grid(row=3, column=0, sticky='ew', padx=10, pady=3)
            error = ctk.CTkLabel(card, text='', text_color='#b42318', wraplength=330, justify='left', anchor='w')
            error.grid(row=4, column=0, sticky='ew', padx=10)
            self.field_errors[field.name] = error
            self.field_cards.append((field, card))
        actions = ctk.CTkFrame(parent, fg_color='transparent')
        actions.grid(row=5, column=0, sticky='ew', pady=(0, 4))
        actions.grid_columnconfigure(2, weight=1)
        self.previous_section_button = self._button(actions, 'Anterior', lambda: self.move_profile_section(-1), width=95)
        self.previous_section_button.grid(row=0, column=0, padx=(0, 8))
        self.next_section_button = self._button(actions, 'Próximo', lambda: self.move_profile_section(1), width=95)
        self.next_section_button.grid(row=0, column=1)
        self._button(actions, 'Salvar perfil', self.save_profile, width=135).grid(row=0, column=3)
        self._layout_form(2)

    def show_profile_section(self, section):
        if section not in PROFILE_SECTIONS:
            return
        self.section_selector.set(section)
        self._layout_form(self._form_columns or 2)
        self.form._parent_canvas.yview_moveto(0)

    def move_profile_section(self, direction):
        index = PROFILE_SECTIONS.index(self.profile_section.get())
        self.show_profile_section(PROFILE_SECTIONS[max(0, min(len(PROFILE_SECTIONS) - 1, index + direction))])

    def choose_font_color(self):
        if self.busy:
            return
        raw = self.field_vars['font_color'].get().lstrip('#')
        initial = '#' + raw if len(raw) == 6 and all(c in '0123456789abcdefABCDEF' for c in raw) else '#000000'
        _rgb, selected = colorchooser.askcolor(color=initial, title='Cor da fonte', parent=self)
        if selected:
            self.field_vars['font_color'].set(selected.lstrip('#').upper())

    def _layout_form(self, columns):
        self._form_columns = columns
        section = self.profile_section.get()
        index = PROFILE_SECTIONS.index(section)
        self.section_hint.set(f'{index + 1} de {len(PROFILE_SECTIONS)} — {section}. Você pode abrir qualquer seção ou salvar quando terminar.')
        self.previous_section_button.configure(state='normal' if index and not self.busy else 'disabled')
        self.next_section_button.configure(state='normal' if index < len(PROFILE_SECTIONS) - 1 and not self.busy else 'disabled')
        for label in self.group_labels.values():
            label.grid_remove()
        for col in range(2):
            self.form.grid_columnconfigure(col, weight=1 if col < columns else 0, uniform='fields' if col < columns else '')
        row, previous, position = 0, None, 0
        for field, card in self.field_cards:
            if field_section(field) != section:
                card.grid_remove()
                continue
            if field.group != previous:
                if position:
                    row += 1
                title = ('Quebras de linha e de página' if field.group == 'Parágrafo' and section == 'Opções avançadas'
                         else 'Fonte' if field.group == 'Texto' else field.group)
                self.group_labels[field.group].configure(text=title)
                self.group_labels[field.group].grid(row=row, column=0, columnspan=columns, sticky='ew', padx=8, pady=(6, 4))
                row += 1
                position = 0
                previous = field.group
            card.grid(row=row, column=position, sticky='nsew', padx=5, pady=5)
            position += 1
            if position == columns:
                row += 1
                position = 0

    def _resize(self, event):
        if event.widget is self:
            columns = 1 if event.width < 900 else 2
            if columns != self._form_columns:
                self._layout_form(columns)
            self.status_label.configure(wraplength=max(650, event.width - 50))
            compact = event.height < 700
            if compact != self._compact_layout:
                self._compact_layout = compact
                self.summary_box.configure(height=55 if compact else 105)
                self.details_box.configure(height=45 if compact else 65)

    def _scroll_to_field(self, name, target=None):
        widget = target if target is not None else self.field_widgets[name]
        if self.tabs.get() != 'Perfis e formatação':
            return
        section = field_section(FIELD_MAP[name])
        if self.profile_section.get() != section:
            self.show_profile_section(section)
        self.update_idletasks()
        canvas = self.form._parent_canvas
        top = canvas.canvasy(0)
        bottom = top + canvas.winfo_height()
        y = widget.winfo_rooty() - self.form.winfo_rooty()
        if y < top:
            canvas.yview_moveto(max(0, (y - 4) / max(1, self.form.winfo_height())))
        elif y + widget.winfo_height() > bottom:
            canvas.yview_moveto(max(0, (y + widget.winfo_height() - canvas.winfo_height() + 4)
                                    / max(1, self.form.winfo_height())))

    def _values(self):
        return {name: variable.get() for name, variable in self.field_vars.items()}

    def _changed(self, *_):
        if self._loading or self._destroyed:
            return
        self.dirty_text.set('Modificações não salvas' if self.dirty else 'Perfil salvo')
        # CTkComboBox deletes its Entry text before inserting the selection.
        # A synchronous readonly refresh here interrupts that insert operation.
        if self._form_after is None:
            self._form_after = self._schedule(0, self._refresh_after_change)
        if self._summary_after:
            self.after_cancel(self._summary_after)
            self._after_ids.discard(self._summary_after)
        self._summary_after = self._schedule(180, self._update_summary)

    def _refresh_after_change(self):
        self._form_after = None
        self._refresh_form()

    def _refresh_form(self):
        for field in FIELDS:
            enabled = not self.busy
            if field.name.startswith(('header_', 'footer_')) and field.name.endswith('alignment'):
                enabled = enabled and self.field_vars[field.name.split('_')[0] + '_mode'].get() == 'Imagem do perfil'
            state = 'disabled' if not enabled or field.kind == 'image' else ('readonly' if field.choices else 'normal')
            self.field_widgets[field.name].configure(state=state)
        for button in self.spacing_buttons:
            button.configure(state='disabled' if self.busy else 'normal')
        self.field_labels['line_spacing'].configure(text='Em (múltiplo de linhas)' if self.field_vars['line_spacing_mode'].get() == 'Múltiplo' else 'Em (pt)')
        index = PROFILE_SECTIONS.index(self.profile_section.get())
        self.previous_section_button.configure(state='normal' if index and not self.busy else 'disabled')
        self.next_section_button.configure(state='normal' if index < len(PROFILE_SECTIONS) - 1 and not self.busy else 'disabled')
        for slot, button in self.image_buttons.items():
            button.configure(state='normal' if not self.busy and self.field_vars[slot + '_mode'].get() == 'Imagem do perfil' else 'disabled')
        legacy = bool(self._draft_base.template_path or self._draft_base.migration_warnings)
        self.review_button.configure(state='normal' if legacy and not self.busy else 'disabled')

    @staticmethod
    def _set_text(widget, text):
        widget.configure(state='normal')
        widget.delete('1.0', 'end')
        widget.insert('1.0', text)
        widget.configure(state='disabled')

    def _parse(self, *, check_assets=False, focus=False):
        for label in self.field_errors.values():
            label.configure(text='')
        try:
            settings = settings_from_values(self._values(), self._draft_base, check_assets=check_assets)
        except SettingsError as exc:
            field = FIELD_MAP.get(exc.field)
            message = f'{field.label}: {exc}' if field else str(exc)
            self.error_text.set(message)
            if field:
                self.field_errors[exc.field].configure(text=message)
            if focus:
                self.tabs.set('Perfis e formatação')
                if field:
                    self.show_profile_section(field_section(field))
                    widget = self.field_widgets[exc.field]
                    self.update_idletasks()
                    card = widget.master
                    self.form._parent_canvas.yview_moveto(max(0, card.winfo_y() / max(1, self.form.winfo_height())))
                    widget.focus_set()
                else:
                    self.review_button.focus_set()
            return None
        self.error_text.set('')
        return settings

    def _update_summary(self):
        self._summary_after = None
        settings = self._parse()
        self._set_text(self.summary_box, settings_summary(settings) if settings else 'Configuração inválida. ' + self.error_text.get())
        self.apply_button.configure(state='normal' if settings and not self.busy else 'disabled')
        self.dirty_text.set('Modificações não salvas' if self.dirty else 'Perfil salvo')

    def _profile_names(self):
        return [DEFAULT_PROFILE, *sorted(self.config_model.stacks)]

    def _load_draft(self, settings, selected):
        self._loading = True
        self._selected = selected
        self._draft_base = deepcopy(settings)
        self._baseline_base = deepcopy(settings)
        self._baseline_values = form_values(settings)
        self._baseline_name = selected
        for name, value in self._baseline_values.items():
            self.field_vars[name].set(value)
        self.profile_name.set(selected)
        self.selected_profile.set(selected or DEFAULT_PROFILE)
        self.profile_selector.configure(values=self._profile_names())
        self._loading = False
        self._refresh_form()
        self._update_summary()

    def _resolve_dirty(self, parent=None):
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel('Modificações não salvas', 'Salvar as modificações antes de continuar?\nSim: salvar. Não: descartar. Cancelar: continuar editando.', parent=parent or self)
        return self.save_profile() if answer is True else answer is False

    def select_profile(self, name):
        if self.busy:
            self.selected_profile.set(self._selected or DEFAULT_PROFILE)
            return False
        selected = '' if name == DEFAULT_PROFILE else name
        if selected == self._selected:
            return True
        if selected and selected not in self.config_model.stacks:
            return False
        if not self._resolve_dirty():
            self.selected_profile.set(self._selected or DEFAULT_PROFILE)
            return False
        candidate = replace(self.config_model, active_stack=selected)
        try:
            self.store.save(candidate)
        except (OSError, ValueError) as exc:
            self.status_text.set(f'Não foi possível selecionar o perfil: {exc}')
            self.selected_profile.set(self._selected or DEFAULT_PROFILE)
            return False
        self.config_model = candidate
        self._load_draft(candidate.stacks.get(selected, candidate.settings), selected)
        return True

    def save_profile(self):
        if self.busy:
            return False
        settings = self._parse(check_assets=True, focus=True)
        if settings is None:
            return False
        name = self.profile_name.get().strip()
        if name == DEFAULT_PROFILE or (not name and self._selected):
            self.error_text.set('Informe um nome de perfil; “Configuração atual” é reservado.')
            self.name_entry.focus_set()
            return False
        if name and name in self.config_model.stacks and name != self._selected:
            if not messagebox.askyesno('Substituir perfil', f'Substituir o perfil “{name}”?', parent=self):
                return False
        stacks = dict(self.config_model.stacks)
        if name:
            stacks[name] = settings
        candidate = AppConfig(settings if not name else self.config_model.settings, stacks, name)
        try:
            self.store.save(candidate)
        except (OSError, ValueError) as exc:
            self.status_text.set(f'Não foi possível salvar: {exc}')
            return False
        self.config_model = candidate
        self._load_draft(settings, name)
        self.status_text.set(f'{name or DEFAULT_PROFILE}: salvo.')
        return True

    def new_profile(self):
        if self.busy or not self._resolve_dirty():
            return
        self._load_draft(FormatSettings(formatting_mode='by_category'), '')
        self.profile_name.set(unique_profile_name('Novo perfil', set(self.config_model.stacks)))
        self.show_profile_section('Fonte')
        self.name_entry.focus_set()

    def import_profile(self, path=None):
        if self.busy:
            return False
        if self.import_dialog is not None and self.import_dialog.winfo_exists():
            self.import_dialog.lift()
            return False
        if path is None:
            path = filedialog.askopenfilename(parent=self, title='Importar perfil de Word',
                filetypes=[('Documentos Word', '*.docx')])
        if not path:
            return False
        path = Path(path).expanduser().resolve()
        self.import_dialog = None
        self._importing = True
        self.cancel_event = threading.Event()
        self._set_busy(True)
        self.status_text.set('Lendo configurações do Word. Aguarde a revisão para criar o perfil.')
        events = self.events

        def work():
            try:
                from app.profile_import import inspect_profile
                report = inspect_profile(path)
                events.put(('profile_imported', (path, report)))
            except Exception as exc:
                events.put(('profile_import_error', str(exc)))

        self.worker = threading.Thread(target=work, name='format-word-import', daemon=True)
        self.worker.start()
        return True

    def _finish_profile_import(self, kind, payload):
        self._importing = False
        self._set_busy(False)
        if self._close_pending:
            self._close_pending = False
            if self._resolve_dirty():
                self.destroy()
            return
        if self.cancel_event.is_set():
            self.status_text.set('Importação cancelada. O perfil foi mantido.')
            return
        if kind == 'profile_import_error':
            self.status_text.set('Não foi possível importar: ' + payload)
            return
        from app.profile_import_dialog import ProfileImportDialog
        path, report = payload
        self.import_dialog = ProfileImportDialog(self, path, report,
            lambda settings: self._accept_profile_import(path, settings))
        self.status_text.set('Revise as configurações detectadas antes de criar o perfil.')

    def _accept_profile_import(self, path, settings):
        if self.busy or not self._resolve_dirty(parent=self.import_dialog):
            return False
        name = unique_profile_name(path.stem.strip() or 'Perfil importado',
            set(self.config_model.stacks) | {DEFAULT_PROFILE})
        self._load_draft(settings, '')
        self.profile_name.set(name)
        self.tabs.set('Perfis e formatação')
        self.show_profile_section('Fonte')
        self._update_summary()
        self.status_text.set('Perfil importado como rascunho. Confira os campos e clique em Salvar.')
        return True

    def duplicate_profile(self):
        if self.busy:
            return
        settings = self._parse(check_assets=True, focus=True)
        if settings is None:
            return
        name = unique_profile_name((self.profile_name.get().strip() or 'Perfil') + ' cópia', set(self.config_model.stacks))
        self.profile_name.set(name)
        self.save_profile()

    def delete_profile(self):
        if self.busy or not self._selected:
            return
        if not messagebox.askyesno('Excluir perfil', f'Excluir “{self._selected}” e descartar suas modificações não salvas?\nOs documentos e imagens externos serão mantidos.', parent=self):
            return
        stacks = dict(self.config_model.stacks)
        del stacks[self._selected]
        candidate = replace(self.config_model, stacks=stacks, active_stack='')
        try:
            self.store.save(candidate)
        except (OSError, ValueError) as exc:
            self.status_text.set(f'Não foi possível excluir: {exc}')
            return
        self.config_model = candidate
        self._load_draft(candidate.settings, '')

    def restore_profile(self):
        if self.busy:
            return
        if self.dirty and not messagebox.askyesno('Restaurar perfil', 'Descartar modificações não salvas e restaurar os valores salvos?', parent=self):
            return
        self._load_draft(self.config_model.stacks.get(self._selected, self.config_model.settings), self._selected)

    def review_legacy(self):
        if self.busy:
            return
        details = list(self._draft_base.migration_warnings)
        if self._draft_base.template_path:
            details.insert(0, 'Referência ao template: ' + self._draft_base.template_path)
        if not details:
            return
        if messagebox.askyesno('Revisar perfil antigo', 'Remover deste rascunho as opções antigas abaixo?\n\n' + '\n'.join(details) + '\n\nArquivos externos serão preservados. Revise cabeçalho/rodapé e salve o perfil para confirmar.', parent=self):
            self._draft_base = replace(self._draft_base, template_path='', migration_warnings=())
            self._changed()
            self._update_summary()

    def _accept_category_rules(self, settings):
        self._draft_base.category_rules = deepcopy(settings.category_rules)
        self.field_vars['formatting_mode'].set('Por tipo de texto')
        self._changed()
        self._update_summary()

    def edit_category_rules(self):
        if self.busy:
            return
        settings = self._parse(focus=True)
        if settings is not None:
            return CategoryRulesDialog(self, settings, self._accept_category_rules)

    def _accept_document_review(self, path, overrides, settings):
        self.document_overrides[path] = deepcopy(overrides)
        self._draft_base.style_categories = deepcopy(settings.style_categories)
        if settings.formatting_mode == 'by_category':
            self.field_vars['formatting_mode'].set('Por tipo de texto')
        self._changed()
        self._update_summary()
        if path in self.paths:
            self.file_tree.set(str(self.paths.index(path)), 'status', 'Revisado')
        self.status_text.set('Revisão confirmada. Use Aplicar aos arquivos para gerar o documento.')

    def review_document(self):
        if self.busy:
            return
        selection = self.file_tree.selection()
        if len(selection) != 1:
            self.status_text.set('Selecione um único documento para revisar.')
            return
        settings = self._parse(focus=True)
        if settings is None:
            return
        path = self.paths[int(selection[0])]
        try:
            inspection = inspect_document(path, settings)
        except Exception as exc:
            self.status_text.set(f'Não foi possível revisar: {exc}')
            return
        return DocumentReviewDialog(self, path, inspection, settings, self.document_overrides.get(path),
                                    lambda overrides, reviewed: self._accept_document_review(path, overrides, reviewed))

    def choose_image(self, slot):
        if self.busy:
            return
        path = filedialog.askopenfilename(parent=self, title='Escolher imagem', filetypes=[('PNG/JPEG', '*.png *.jpg *.jpeg')])
        if path:
            try:
                self.field_vars[slot + '_image_path'].set(self.store.store_image(Path(path), slot))
            except (OSError, ValueError) as exc:
                self.error_text.set(str(exc))

    def choose_files(self):
        if not self.busy:
            self.add_files(filedialog.askopenfilenames(parent=self, title='Selecionar documentos Word', filetypes=[('Documentos Word', '*.docx')]))

    def add_files(self, paths):
        if self.busy:
            return
        rejected = []
        for value in paths:
            path = Path(value).expanduser().resolve()
            if path.suffix.lower() != '.docx':
                rejected.append(path.name)
            elif path not in self.paths:
                self.paths.append(path)
        self._render_files()
        self.status_text.set('Somente DOCX é aceito: ' + ', '.join(rejected) if rejected else f'{len(self.paths)} arquivo(s) selecionado(s).')

    def _render_files(self):
        self.document_overrides = {path: value for path, value in self.document_overrides.items() if path in self.paths}
        self.file_tree.delete(*self.file_tree.get_children())
        self.results = []
        self.progress.set(0)
        self.progress_label.configure(text=f'0 / {len(self.paths)}')
        for index, path in enumerate(self.paths):
            self.file_tree.insert('', 'end', iid=str(index), values=(str(path), 'Revisado' if path in self.document_overrides else 'Aguardando'))

    def remove_selected(self):
        if not self.busy:
            selected = set(self.file_tree.selection())
            self.paths = [path for i, path in enumerate(self.paths) if str(i) not in selected]
            self._render_files()

    def clear_files(self):
        if not self.busy:
            self.paths.clear()
            self._render_files()

    def choose_output(self):
        if not self.busy:
            path = filedialog.askdirectory(parent=self, title='Pasta de destino')
            if path:
                self.output_dir.set(path)

    def _set_busy(self, busy):
        self.busy = busy
        for widget in self.mutation_widgets:
            widget.configure(state='disabled' if busy else 'normal')
        self.profile_selector.configure(state='disabled' if busy else 'readonly')
        self.cancel_button.configure(state='normal' if busy else 'disabled')
        self._refresh_form()
        if not busy:
            self._update_summary()

    def start_batch(self):
        if self.busy:
            return False
        if not self.paths:
            self.status_text.set('Selecione ao menos um arquivo Word.')
            return False
        if not self.output_dir.get().strip():
            self.status_text.set('Escolha uma pasta de destino.')
            return False
        settings = self._parse(check_assets=True, focus=True)
        if settings is None:
            return False
        settings = deepcopy(settings)
        overrides = deepcopy(self.document_overrides)
        destination = Path(self.output_dir.get().strip()).expanduser().resolve()
        paths = tuple(self.paths)
        self._render_files()
        self.cancel_event = threading.Event()
        self._set_busy(True)
        self.status_text.set('Processando. O cancelamento ocorre entre arquivos.')
        events, cancel = self.events, self.cancel_event

        def work():
            try:
                format_documents_batch(paths, destination, settings, cancel_event=cancel, overrides_by_path=overrides,
                                       on_result=lambda item: events.put(('result', item)))
            except Exception as exc:
                events.put(('error', str(exc)))
            finally:
                events.put(('done', None))

        self.worker = threading.Thread(target=work, name='format-word-batch', daemon=True)
        self.worker.start()
        return True

    def cancel_batch(self):
        if self.busy:
            self.cancel_event.set()
            self.cancel_button.configure(state='disabled')
            self.status_text.set('Cancelamento solicitado; aguardando a leitura do Word.' if self._importing
                else 'Cancelamento solicitado; aguardando o arquivo em andamento.')

    def _poll(self):
        while not self.events.empty():
            kind, payload = self.events.get_nowait()
            if kind in ('profile_imported', 'profile_import_error'):
                self._finish_profile_import(kind, payload)
                if self._destroyed:
                    return
            elif kind == 'result':
                self.results.append(payload)
                index = len(self.results) - 1
                status = ('Cancelado' if payload.cancelled else 'Erro' if payload.error else
                          'Concluído com ressalvas' if payload.result.warnings else 'Concluído')
                self.file_tree.set(str(index), 'status', status)
                self.progress.set(len(self.results) / max(1, len(self.paths)))
                self.progress_label.configure(text=f'{len(self.results)} / {len(self.paths)}')
                self.file_tree.selection_set(str(index))
                self._show_result()
            elif kind == 'error':
                self.status_text.set('Falha no lote: ' + payload)
                self._set_text(self.details_box, payload)
            elif kind == 'done':
                self._set_busy(False)
                if len(self.results) == len(self.paths):
                    ok = sum(bool(item.result and not item.result.warnings) for item in self.results)
                    warnings = sum(bool(item.result and item.result.warnings) for item in self.results)
                    errors = sum(bool(item.error) for item in self.results)
                    cancelled = sum(item.cancelled for item in self.results)
                    self.status_text.set(f'Lote encerrado: {ok} concluído(s), {warnings} com ressalvas, {errors} erro(s), {cancelled} cancelado(s).')
                if self._close_pending:
                    self._close_pending = False
                    if self._resolve_dirty():
                        self.destroy()
                        return
        self._schedule(80, self._poll)

    def _show_result(self, *_):
        selection = self.file_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        if index >= len(self.results):
            self._set_text(self.details_box, str(self.paths[index]) + '\nAguardando processamento.')
            return
        item = self.results[index]
        text = str(item.input_path) + '\n'
        if item.result:
            text += str(item.result.output_path) + '\n' + '\n'.join(item.result.warnings)
        else:
            text += item.error or 'Cancelado antes de processar este arquivo.'
        self._set_text(self.details_box, text)

    def _open_path(self, path):
        try:
            if not path.exists():
                raise OSError('Caminho não encontrado.')
            if sys.platform == 'win32':
                os.startfile(str(path))
            else:
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(path)])
        except OSError as exc:
            self.status_text.set(f'Não foi possível abrir: {exc}')

    def open_selected(self):
        selection = self.file_tree.selection()
        if selection:
            index = int(selection[0])
            if index < len(self.results) and self.results[index].result:
                self._open_path(self.results[index].result.output_path)

    def open_output(self):
        if self.output_dir.get().strip():
            self._open_path(Path(self.output_dir.get().strip()).expanduser().resolve())

    def request_close(self):
        if self.busy:
            title = 'Importação em andamento' if self._importing else 'Lote em andamento'
            prompt = ('Cancelar a importação e fechar quando a leitura terminar?' if self._importing
                else 'Cancelar os arquivos pendentes e fechar quando o arquivo atual terminar?')
            if messagebox.askyesno(title, prompt, parent=self):
                self._close_pending = True
                self.cancel_batch()
        elif self._resolve_dirty():
            self.destroy()

    def destroy(self):
        if self._destroyed:
            return
        self._destroyed = True
        self.cancel_event.set()
        for identifier in self._after_ids:
            try:
                self.after_cancel(identifier)
            except tk.TclError:
                pass
        self._after_ids.clear()
        super().destroy()
