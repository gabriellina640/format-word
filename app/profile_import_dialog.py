"""Review detected DOCX settings before creating a reusable profile draft."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from app.config import CATEGORY_LABELS, SettingsError
from app.ui_fields import FIELD_MAP, settings_summary


class ProfileImportDialog(tk.Toplevel):
    def __init__(self, parent, path, report, on_confirm):
        super().__init__(parent)
        self.title('Criar perfil a partir de um Word')
        self.geometry('740x650')
        self.minsize(560, 500)
        self.transient(parent)
        self.report = report
        self.on_confirm = on_confirm
        self.selections = {category: 0 for category in report.category_options}
        self.categories = list(report.category_options)
        self.error_var = tk.StringVar(self)
        self.sample_var = tk.StringVar(self)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        header = ttk.Frame(self, padding=(16, 12, 16, 4))
        header.grid(row=0, column=0, sticky='ew')
        header.columnconfigure(0, weight=1)
        self.heading = ttk.Label(header, text=f'Referência: {path.name}', wraplength=690)
        self.heading.grid(row=0, column=0, sticky='w')
        self.intro = ttk.Label(header, text='Confira as configurações detectadas. Quando houver diferenças, '
            'escolha a página e a formatação de cada tipo de texto. Depois, confira o nome e salve o perfil.', wraplength=690)
        self.intro.grid(row=1, column=0, sticky='ew', pady=(6, 0))

        choices = ttk.Frame(self, padding=(16, 8))
        choices.grid(row=1, column=0, sticky='ew')
        choices.columnconfigure(1, weight=1)
        ttk.Label(choices, text='Configuração de página').grid(row=0, column=0, sticky='w', padx=(0, 12))
        self.page_selector = ttk.Combobox(choices, state='readonly',
            values=[option.label for option in report.page_options])
        self.page_selector.grid(row=0, column=1, sticky='ew', pady=4)
        self.page_selector.current(0)
        self.page_selector.bind('<<ComboboxSelected>>', self.update_preview)
        ttk.Label(choices, text='Tipo de texto').grid(row=1, column=0, sticky='w', padx=(0, 12))
        self.category_selector = ttk.Combobox(choices, state='readonly',
            values=[CATEGORY_LABELS[category] for category in self.categories])
        self.category_selector.grid(row=1, column=1, sticky='ew', pady=4)
        self.category_selector.bind('<<ComboboxSelected>>', self.switch_category)
        ttk.Label(choices, text='Formatação encontrada').grid(row=2, column=0, sticky='w', padx=(0, 12))
        self.option_selector = ttk.Combobox(choices, state='readonly')
        self.option_selector.grid(row=2, column=1, sticky='ew', pady=4)
        self.option_selector.bind('<<ComboboxSelected>>', self.select_option)
        self.sample = ttk.Label(self, textvariable=self.sample_var, wraplength=690)
        self.sample.grid(row=2, column=0, sticky='ew', padx=16, pady=(0, 8))

        notebook = ttk.Notebook(self)
        notebook.grid(row=3, column=0, sticky='nsew', padx=16)
        self.summary = self._text_tab(notebook, 'Configurações do perfil')
        notices = self._text_tab(notebook, f'Avisos e limites ({len(report.notices)})')
        self._write(notices, '\n\n'.join(report.notices) or 'Nenhum aviso adicional.')
        self.notice = ttk.Label(self, text='Cabeçalho, rodapé e imagens da referência não são copiados. '
            'O perfil preserva esses elementos nos documentos de destino. Leia a aba Avisos e limites.',
            wraplength=690)
        self.notice.grid(row=4, column=0, sticky='ew', padx=16, pady=(8, 0))
        self.error = ttk.Label(self, textvariable=self.error_var, foreground='#b42318', wraplength=690)
        self.error.grid(row=5, column=0, sticky='ew', padx=16, pady=4)
        actions = ttk.Frame(self, padding=(16, 4, 16, 12))
        actions.grid(row=6, column=0, sticky='e')
        ttk.Button(actions, text='Cancelar', command=self.destroy).pack(side='left', padx=8)
        self.confirm_button = ttk.Button(actions, text='Usar estas configurações', command=self.confirm)
        self.confirm_button.pack(side='left')
        self.bind('<Escape>', lambda _event: self.destroy())
        self.bind('<Configure>', self._resize)
        if self.categories:
            self.category_selector.current(0)
            self.switch_category()
        else:
            self.category_selector.configure(state='disabled')
            self.option_selector.configure(state='disabled')
            self.update_preview()
        self.grab_set()
        self.page_selector.focus_set()

    @staticmethod
    def _text_tab(notebook, title):
        frame = ttk.Frame(notebook)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        text = tk.Text(frame, wrap='word', state='disabled', width=1, height=8,
                       padx=10, pady=8, relief='flat', background='white', foreground='#101828')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        notebook.add(frame, text=title)
        return text

    @staticmethod
    def _write(widget, value):
        widget.configure(state='normal')
        widget.delete('1.0', 'end')
        widget.insert('1.0', value)
        widget.configure(state='disabled')

    def _resize(self, event):
        if event.widget is self:
            for label in (self.heading, self.intro, self.sample, self.notice, self.error):
                label.configure(wraplength=max(300, event.width - 40))

    def switch_category(self, _event=None):
        category = self.categories[self.category_selector.current()]
        options = self.report.category_options[category]
        self.option_selector.configure(values=[option.label for option in options])
        self.option_selector.current(self.selections[category])
        self.update_preview()

    def select_option(self, _event=None):
        category = self.categories[self.category_selector.current()]
        self.selections[category] = self.option_selector.current()
        self.update_preview()

    def _settings(self):
        return self.report.build_settings(page_index=self.page_selector.current(), selections=self.selections)

    def update_preview(self, _event=None):
        if self.categories:
            category = self.categories[self.category_selector.current()]
            option = self.report.category_options[category][self.selections[category]]
            self.sample_var.set('Exemplo: ' + (option.sample or 'Sem trecho de exemplo.'))
        try:
            settings = self._settings()
        except (SettingsError, ValueError) as exc:
            field = FIELD_MAP.get(getattr(exc, 'field', ''))
            self.error_var.set((field.label + ': ' if field else '') + str(exc))
            self._write(self.summary, 'Este padrão contém valores que o perfil não suporta. '
                'Escolha outra alternativa ou ajuste o documento de referência.\n\n' + self.error_var.get())
            self.confirm_button.configure(state='disabled')
            return
        self.error_var.set('')
        self._write(self.summary, settings_summary(settings))
        self.confirm_button.configure(state='normal')

    def confirm(self):
        try:
            settings = self._settings()
        except (SettingsError, ValueError):
            self.update_preview()
            return
        if self.on_confirm(settings):
            self.destroy()
