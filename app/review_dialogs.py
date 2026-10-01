"""Transactional editors for category rules and document-specific adjustments."""
from __future__ import annotations

from copy import deepcopy
import math
import tkinter as tk
from tkinter import colorchooser, ttk

from app.config import CATEGORY_LABELS, TEXT_FIELDS, SettingsError, paper_dimensions, validate_settings
from app.document_model import DocumentOverrides, ImageEdit
from app.ui_fields import FIELDS, form_values, settings_from_values, number_text
from app.word_controls import SpecialIndentControl, spacing_shortcuts

RULE_MODES = {'Seguir corpo': 'inherit', 'Preservar original': 'preserve', 'Personalizar': 'custom'}
IMAGE_ALIGNMENT = {'Preservar': 'preserve', 'Esquerda': 'left', 'Centralizado': 'center', 'Direita': 'right'}


class CategoryRulesDialog(tk.Toplevel):
    def __init__(self, parent, settings, on_confirm):
        super().__init__(parent)
        self.title('Formatação por tipo de texto')
        self.geometry('660x670')
        self.minsize(540, 480)
        self.transient(parent)
        self.settings = deepcopy(settings)
        self.on_confirm = on_confirm
        self.current_category = 'body'
        self.category_var = tk.StringVar(self, CATEGORY_LABELS['body'])
        self.mode_var = tk.StringVar(self)
        self.error_var = tk.StringVar(self)
        self.variables = {}
        self.controls = {}
        self.labels = {}
        self.extra_controls = []
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        ttk.Label(self, text='Escolha o tipo de texto e como formatá-lo. Seguir corpo usa o perfil principal.\nPreservar original mantém o destino; Personalizar permite definir valores abaixo.', wraplength=600).grid(row=0, column=0, sticky='ew', padx=16, pady=12)
        selector = ttk.Combobox(self, textvariable=self.category_var, values=list(CATEGORY_LABELS.values()), state='readonly')
        selector.grid(row=1, column=0, sticky='ew', padx=16)
        selector.bind('<<ComboboxSelected>>', self.switch_category)
        mode = ttk.Combobox(self, textvariable=self.mode_var, values=list(RULE_MODES), state='readonly')
        mode.grid(row=2, column=0, sticky='ew', padx=16, pady=8)
        mode.bind('<<ComboboxSelected>>', lambda _event: self.refresh_controls())
        frame = ttk.Frame(self)
        frame.grid(row=3, column=0, sticky='nsew', padx=16)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        canvas = tk.Canvas(frame, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        canvas.configure(yscrollcommand=scrollbar.set)
        form = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=form, anchor='nw')
        canvas.bind('<Configure>', lambda event: canvas.itemconfigure(window, width=event.width))
        form.bind('<Configure>', lambda _event: canvas.configure(scrollregion=canvas.bbox('all')))
        form.columnconfigure(1, weight=1)
        for row, field in enumerate(f for f in FIELDS if f.name in TEXT_FIELDS):
            variable = tk.StringVar(self)
            self.variables[field.name] = variable
            label = ttk.Label(form, text=field.label, wraplength=180)
            label.grid(row=row, column=0, sticky='w', pady=7, padx=(0, 12))
            self.labels[field.name] = label
            cell = ttk.Frame(form)
            cell.grid(row=row, column=1, sticky='ew', pady=7)
            cell.columnconfigure(0, weight=1)
            extra_focus = []
            if field.name == 'first_line_indent_cm':
                control = SpecialIndentControl(cell, variable)
            else:
                control = (ttk.Combobox(cell, textvariable=variable, values=list(field.choices), state='readonly')
                           if field.choices else ttk.Entry(cell, textvariable=variable))
            control.grid(row=0, column=0, sticky='ew')
            if field.name == 'line_spacing_mode':
                shortcuts, self.spacing_buttons = spacing_shortcuts(cell, self.variables)
                shortcuts.grid(row=1, column=0, sticky='ew', pady=(4, 0))
                self.extra_controls.extend(self.spacing_buttons)
                extra_focus.extend(self.spacing_buttons)
            if field.name == 'font_color':
                colors = ttk.Frame(cell)
                colors.grid(row=1, column=0, sticky='ew', pady=(4, 0))
                for text, command in (('Escolher cor…', self.choose_font_color),
                                      ('Manter original', lambda: self.variables['font_color'].set(''))):
                    button = ttk.Button(colors, text=text, command=command)
                    button.pack(side='left', padx=(0, 4))
                    self.extra_controls.append(button)
                    extra_focus.append(button)
            def reveal(_event, widget=cell):
                self.update_idletasks()
                y = widget.winfo_y()
                if y < canvas.canvasy(0) or y + widget.winfo_height() > canvas.canvasy(canvas.winfo_height()):
                    canvas.yview_moveto(y / max(1, form.winfo_height()))
            for target in (*getattr(control, 'focus_targets', (control,)), *extra_focus):
                target.bind('<FocusIn>', reveal)
            self.controls[field.name] = control
        self.variables['line_spacing_mode'].trace_add('write', self._spacing_unit)
        ttk.Label(self, textvariable=self.error_var, foreground='#b42318', wraplength=600).grid(row=4, column=0, sticky='ew', padx=16, pady=6)
        actions = ttk.Frame(self)
        actions.grid(row=5, column=0, sticky='e', padx=16, pady=12)
        ttk.Button(actions, text='Cancelar', command=self.destroy).pack(side='left', padx=6)
        ttk.Button(actions, text='Confirmar formatação', command=self.confirm).pack(side='left')
        self.bind('<Escape>', lambda _event: self.destroy())
        self.load_category()
        self.grab_set()

    def load_category(self):
        rule = self.settings.category_rules.get(self.current_category, {})
        mode = rule.get('mode', 'inherit' if self.current_category == 'body' else 'preserve')
        self.mode_var.set(next(label for label, value in RULE_MODES.items() if value == mode))
        effective = deepcopy(self.settings)
        for name, value in rule.get('values', {}).items():
            setattr(effective, name, value)
        values = form_values(effective)
        for name, variable in self.variables.items():
            variable.set(values[name])
        self.refresh_controls()

    def refresh_controls(self):
        custom = RULE_MODES[self.mode_var.get()] == 'custom'
        for field in FIELDS:
            if field.name in self.controls:
                self.controls[field.name].configure(state=('readonly' if field.choices else 'normal') if custom else 'disabled')
        for button in self.extra_controls:
            button.configure(state='normal' if custom else 'disabled')
        self._spacing_unit()

    def _spacing_unit(self, *_):
        self.labels['line_spacing'].configure(text='Em (múltiplo de linhas)' if self.variables['line_spacing_mode'].get() == 'Múltiplo' else 'Em (pt)')

    def choose_font_color(self):
        if RULE_MODES[self.mode_var.get()] != 'custom':
            return
        raw = self.variables['font_color'].get().lstrip('#')
        initial = '#' + raw if len(raw) == 6 and all(c in '0123456789abcdefABCDEF' for c in raw) else '#000000'
        _rgb, selected = colorchooser.askcolor(color=initial, title='Cor da fonte', parent=self)
        if selected:
            self.variables['font_color'].set(selected.lstrip('#').upper())

    def save_category(self):
        mode = RULE_MODES[self.mode_var.get()]
        values = {}
        if mode == 'custom':
            raw = form_values(self.settings)
            raw.update({name: variable.get() for name, variable in self.variables.items()})
            try:
                base = deepcopy(self.settings)
                base.category_rules = {}
                parsed = settings_from_values(raw, base)
            except SettingsError as exc:
                self.error_var.set(str(exc))
                return False
            values = {name: getattr(parsed, name) for name in TEXT_FIELDS}
        self.settings.category_rules[self.current_category] = {'mode': mode, 'values': values}
        self.error_var.set('')
        return True

    def switch_category(self, _event=None):
        if not self.save_category():
            self.category_var.set(CATEGORY_LABELS[self.current_category])
            return
        self.current_category = next(key for key, label in CATEGORY_LABELS.items() if label == self.category_var.get())
        self.load_category()

    def confirm(self):
        if self.save_category():
            self.settings.formatting_mode = 'by_category'
            try:
                settings = validate_settings(self.settings)
            except SettingsError as exc:
                self.error_var.set(str(exc))
                return
            self.on_confirm(settings)
            self.destroy()


class DocumentReviewDialog(tk.Toplevel):
    def __init__(self, parent, path, inspection, settings, overrides, on_confirm):
        super().__init__(parent)
        self.title('Revisar documento — ' + path.name)
        self.geometry('840x680')
        self.minsize(640, 520)
        self.transient(parent)
        self.inspection = inspection
        self.settings = deepcopy(settings)
        self.overrides = deepcopy(overrides) if overrides and overrides.fingerprint == inspection.fingerprint else DocumentOverrides(inspection.fingerprint)
        self.on_confirm = on_confirm
        self.image_index = None
        self.category_var = tk.StringVar(self, CATEGORY_LABELS['body'])
        self.same_style = tk.BooleanVar(self, False)
        self.error_var = tk.StringVar(self)
        self.width_var = tk.StringVar(self)
        self.alignment_var = tk.StringVar(self, 'Preservar')
        self.target_var = tk.StringVar(self, 'Manter posição')
        self.placement_var = tk.StringVar(self, 'Depois')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        ttk.Label(self, text='Revisão opcional: selecione trechos e indique o tipo. O estilo Normal não identifica títulos.\nOs ajustes só serão usados ao aplicar o lote; o original permanece intacto.', wraplength=600).grid(row=0, column=0, sticky='ew', padx=14, pady=10)
        tabs = ttk.Notebook(self)
        tabs.grid(row=1, column=0, sticky='nsew', padx=14)
        text_tab, image_tab = ttk.Frame(tabs, padding=8), ttk.Frame(tabs, padding=8)
        tabs.add(text_tab, text='Tipos de texto')
        tabs.add(image_tab, text='Imagens')
        text_tab.columnconfigure(0, weight=1)
        text_tab.rowconfigure(0, weight=1)
        self.paragraph_tree = self.tree(text_tab, ('text', 'style', 'category'), ('Trecho', 'Estilo', 'Tipo'), extended=True)
        for paragraph in inspection.paragraphs:
            self.paragraph_tree.insert('', 'end', iid=str(paragraph.index), values=(f'{paragraph.index + 1}. {paragraph.text[:180] or "[Parágrafo sem texto]"}', paragraph.style_name, self.paragraph_label(paragraph)))
        controls = ttk.Frame(text_tab)
        controls.grid(row=1, column=0, columnspan=2, sticky='ew', pady=8)
        ttk.Combobox(controls, textvariable=self.category_var, values=list(CATEGORY_LABELS.values()), state='readonly', width=17).pack(side='left')
        ttk.Button(controls, text='Classificar selecionados', command=self.classify_selected).pack(side='left', padx=8)
        ttk.Checkbutton(text_tab, text='Usar esse tipo para o mesmo estilo neste perfil (afeta outros documentos)', variable=self.same_style).grid(row=2, column=0, columnspan=2, sticky='w')
        image_tab.columnconfigure(0, weight=1)
        image_tab.rowconfigure(0, weight=1)
        self.image_tree = self.tree(image_tab, ('label', 'size', 'status'), ('Imagem', 'Dimensões (cm)', 'Edição'))
        for item in inspection.images:
            self.image_tree.insert('', 'end', iid=str(item.index), values=(item.label, f'{item.width_cm:.2f} × {item.height_cm:.2f}', 'Em linha' if item.editable else 'Flutuante: preservada' if item.floating else 'Protegida: preservada'))
        self.image_tree.bind('<<TreeviewSelect>>', self.select_image)
        editor = ttk.Frame(image_tab)
        editor.grid(row=1, column=0, columnspan=2, sticky='ew', pady=8)
        editor.columnconfigure(1, weight=1)
        self.targets = {'Manter posição': None}
        self.targets.update({f'{p.index + 1}. {p.text[:65] or "[Sem texto]"}': p.index for p in inspection.paragraphs if not p.protected and p.in_body})
        self.image_controls = []
        for row, (label, variable, choices) in enumerate((('Largura (cm), vazio preserva', self.width_var, None), ('Alinhamento', self.alignment_var, list(IMAGE_ALIGNMENT)), ('Parágrafo de referência (corpo)', self.target_var, list(self.targets)), ('Mover imagem', self.placement_var, ['Antes', 'Depois']))):
            ttk.Label(editor, text=label).grid(row=row, column=0, sticky='w', padx=(0, 10), pady=5)
            control = ttk.Combobox(editor, textvariable=variable, values=choices, state='disabled') if choices else ttk.Entry(editor, textvariable=variable, state='disabled')
            control.grid(row=row, column=1, sticky='ew', pady=5)
            self.image_controls.append((control, bool(choices)))
        ttk.Label(image_tab, text='Os valores iniciais seguem o perfil. Vazio / Preservar mantém o original desta imagem.\nRedimensionar, alinhar ou mover cria um parágrafo próprio; a altura mantém a proporção.\nImagens flutuantes são preservadas; para editá-las, use Em linha no Word.', wraplength=590).grid(row=2, column=0, columnspan=2, sticky='w')
        notices = '\n'.join(inspection.notices)
        if overrides and overrides.fingerprint != inspection.fingerprint:
            notices = 'O arquivo mudou; os ajustes anteriores foram descartados nesta revisão.\n' + notices
        ttk.Label(self, text=notices, wraplength=600).grid(row=2, column=0, sticky='ew', padx=14, pady=6)
        ttk.Label(self, textvariable=self.error_var, foreground='#b42318', wraplength=600).grid(row=3, column=0, sticky='ew', padx=14)
        actions = ttk.Frame(self)
        actions.grid(row=4, column=0, sticky='e', padx=14, pady=12)
        ttk.Button(actions, text='Cancelar', command=self.destroy).pack(side='left', padx=8)
        ttk.Button(actions, text='Confirmar revisão', command=self.confirm).pack(side='left')
        self.bind('<Escape>', lambda _event: self.destroy())
        self.grab_set()

    @staticmethod
    def tree(parent, columns, labels, extended=False):
        tree = ttk.Treeview(parent, columns=columns, show='headings', selectmode='extended' if extended else 'browse', height=6)
        for column, label in zip(columns, labels):
            tree.heading(column, text=label)
            tree.column(column, width=320 if column == columns[0] else 130, minwidth=80)
        tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=scrollbar.set)
        return tree

    def paragraph_label(self, paragraph):
        category = self.overrides.paragraph_categories.get(paragraph.index, self.settings.style_categories.get(paragraph.style_id, self.settings.style_categories.get(paragraph.style_name, paragraph.category)))
        return CATEGORY_LABELS[category] + (' (protegido)' if paragraph.protected else '')

    def classify_selected(self):
        selected = {int(index) for index in self.paragraph_tree.selection()}
        if not selected:
            self.error_var.set('Selecione um ou mais trechos para classificar.')
            return
        category = next(key for key, label in CATEGORY_LABELS.items() if label == self.category_var.get())
        for paragraph in self.inspection.paragraphs:
            if paragraph.index in selected and not paragraph.protected:
                self.overrides.paragraph_categories[paragraph.index] = category
                if self.same_style.get() and paragraph.style_id:
                    self.settings.style_categories[paragraph.style_id] = category
        self.settings.formatting_mode = 'by_category'
        for paragraph in self.inspection.paragraphs:
            self.paragraph_tree.set(str(paragraph.index), 'category', self.paragraph_label(paragraph))
        self.error_var.set('Classificação registrada. Trechos protegidos permanecem preservados.')

    def save_image(self):
        if self.image_index is None:
            return True
        item = next(image for image in self.inspection.images if image.index == self.image_index)
        if not item.editable:
            return True
        try:
            width = float(self.width_var.get().strip().replace(',', '.')) if self.width_var.get().strip() else None
            page_width, _ = paper_dimensions(self.settings)
            available_width = page_width - self.settings.margin_left_cm - self.settings.margin_right_cm
            if width is not None and (not math.isfinite(width) or not .01 <= width <= available_width):
                raise ValueError
        except ValueError:
            self.error_var.set('Informe uma largura de pelo menos 0,01 cm que caiba entre as margens, ou deixe vazio.')
            return False
        edit = ImageEdit(width, IMAGE_ALIGNMENT[self.alignment_var.get()], self.targets[self.target_var.get()], 'before' if self.placement_var.get() == 'Antes' else 'after')
        if edit == ImageEdit(self.settings.body_image_width_cm, self.settings.body_image_alignment):
            self.overrides.image_edits.pop(self.image_index, None)
        else:
            self.overrides.image_edits[self.image_index] = edit
        self.error_var.set('')
        return True

    def select_image(self, _event=None):
        selection = self.image_tree.selection()
        if not selection or int(selection[0]) == self.image_index:
            return
        if not self.save_image():
            self.image_tree.selection_set(str(self.image_index))
            return
        self.image_index = int(selection[0])
        item = next(image for image in self.inspection.images if image.index == self.image_index)
        edit = self.overrides.image_edits.get(self.image_index, ImageEdit(self.settings.body_image_width_cm, self.settings.body_image_alignment))
        self.width_var.set('' if edit.width_cm is None else number_text(edit.width_cm))
        self.alignment_var.set(next(label for label, value in IMAGE_ALIGNMENT.items() if value == edit.alignment))
        self.target_var.set(next((label for label, value in self.targets.items() if value == edit.target_paragraph), 'Manter posição'))
        self.placement_var.set('Antes' if edit.placement == 'before' else 'Depois')
        for control, choices in self.image_controls:
            control.configure(state=('readonly' if choices else 'normal') if item.editable else 'disabled')

    def confirm(self):
        if self.save_image():
            self.on_confirm(deepcopy(self.overrides), deepcopy(self.settings))
            self.destroy()
