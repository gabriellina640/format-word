"""Word-like presentation controls without changing persisted profile values."""
from __future__ import annotations

import math
import tkinter as tk
from tkinter import ttk


class SpecialIndentControl(ttk.Frame):
    """Present a signed first-line indent as Special + positive By, like Word."""
    def __init__(self, parent, variable):
        super().__init__(parent)
        self.variable = variable
        self.kind = tk.StringVar(self, 'Primeira linha')
        self.amount = tk.StringVar(self)
        self._syncing = False
        self._state = 'normal'
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text='Especial').grid(row=0, column=0, sticky='w')
        ttk.Label(self, text='Por (cm)').grid(row=0, column=1, sticky='w', padx=(8, 0))
        self.selector = ttk.Combobox(self, textvariable=self.kind, state='readonly', width=14,
            values=('Nenhum', 'Primeira linha', 'Deslocado'))
        self.selector.grid(row=1, column=0, sticky='ew')
        self.entry = ttk.Entry(self, textvariable=self.amount, width=9)
        self.entry.grid(row=1, column=1, sticky='ew', padx=(8, 0))
        self.focus_targets = (self.selector, self.entry)
        self.selector.bind('<<ComboboxSelected>>', self._from_controls)
        self._amount_trace = self.amount.trace_add('write', self._from_controls)
        self._source_trace = variable.trace_add('write', self._from_source)
        self.bind('<Destroy>', self._cleanup, add='+')
        self._from_source()

    def _from_source(self, *_):
        if self._syncing:
            return
        self._syncing = True
        try:
            raw = self.variable.get().strip()
            try:
                value = float(raw.replace(',', '.'))
                if not math.isfinite(value):
                    raise ValueError()
            except ValueError:
                self.kind.set('Primeira linha')
                self.amount.set(raw)
            else:
                self.kind.set('Nenhum' if value == 0 else 'Deslocado' if value < 0 else 'Primeira linha')
                self.amount.set(raw.lstrip('+-') if value else '0')
            self._refresh_state()
        finally:
            self._syncing = False

    def _from_controls(self, *_):
        if self._syncing or self._state == 'disabled':
            return
        self._syncing = True
        try:
            raw = self.amount.get().strip()
            # Prefix the sign instead of abs()/float(): invalid input, including
            # a negative magnitude, remains invalid for the shared validator.
            value = '0' if self.kind.get() == 'Nenhum' else ('-' if self.kind.get() == 'Deslocado' else '+') + raw
            self.variable.set(value)
            self._refresh_state()
        finally:
            self._syncing = False

    def _refresh_state(self):
        self.selector.configure(state='disabled' if self._state == 'disabled' else 'readonly')
        self.entry.configure(state='disabled' if self._state == 'disabled' or self.kind.get() == 'Nenhum' else 'normal')

    def configure(self, cnf=None, **kwargs):
        state = kwargs.pop('state', None)
        if state is not None:
            self._state = state
            self._refresh_state()
        return super().configure(cnf, **kwargs)

    def cget(self, key):
        return self._state if key == 'state' else super().cget(key)

    def focus_set(self):
        (self.entry if str(self.entry.cget('state')) == 'normal' else self.selector).focus_set()

    def _cleanup(self, event):
        if event.widget is self:
            self.variable.trace_remove('write', self._source_trace)
            self.amount.trace_remove('write', self._amount_trace)


def spacing_shortcuts(parent, variables):
    """Return the preset button row; callers control whether editing is enabled."""
    frame = ttk.Frame(parent)
    buttons = []
    def apply(value):
        variables['line_spacing_mode'].set('Múltiplo')
        variables['line_spacing'].set(value)
    for index, (label, value) in enumerate((('Simples', '1'), ('1,5 linhas', '1,5'), ('Duplo', '2'))):
        button = ttk.Button(frame, text=label, command=lambda value=value: apply(value), width=9)
        button.grid(row=0, column=index, padx=(0, 4), sticky='ew')
        frame.columnconfigure(index, weight=1)
        buttons.append(button)
    return frame, buttons
