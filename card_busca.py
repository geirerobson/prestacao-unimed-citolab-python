"""
Card de busca para pesquisa de fichas.
"""

import tkinter as tk
from tkinter import ttk

from config import COLORS, FONT_NORMAL, FONT_INPUT, OPCOES_FILTRO_BUSCA


class CardBusca(tk.Frame):
    """Card com campo de pesquisa, combobox de filtro e botão Limpar."""

    def __init__(self, parent, on_busca=None, on_limpar=None, opcoes_campo=None, **kwargs):
        kwargs.setdefault("bg", COLORS["card_bg"])
        super().__init__(parent, **kwargs)
        self._on_busca = on_busca
        self._on_limpar = on_limpar
        self._opcoes_campo = opcoes_campo if opcoes_campo is not None else OPCOES_FILTRO_BUSCA
        self._construir()

    def _construir(self):
        row_search = tk.Frame(self, bg=COLORS["card_bg"], padx=20, pady=16)
        row_search.pack(fill=tk.X)

        tk.Label(
            row_search,
            text="Pesquisar:",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
        ).pack(side=tk.LEFT, padx=(0, 8), pady=4)

        self.var_termo = tk.StringVar()
        self.var_termo.trace_add("write", self._ao_digitar)
        self.entry_busca = tk.Entry(
            row_search,
            textvariable=self.var_termo,
            font=FONT_INPUT,
            width=35,
            relief=tk.FLAT,
            bg=COLORS["bg"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
        )
        self.entry_busca.pack(side=tk.LEFT, padx=(0, 12), ipady=8, ipadx=10)

        tk.Label(
            row_search,
            text="Campo:",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
        ).pack(side=tk.LEFT, padx=(12, 8), pady=4)

        self.var_campo = tk.StringVar(value="todos")
        self.combo_campo = ttk.Combobox(
            row_search,
            textvariable=self.var_campo,
            values=self._opcoes_campo,
            state="readonly",
            width=16,
            font=FONT_NORMAL,
        )
        self.combo_campo.pack(side=tk.LEFT, padx=(0, 12), pady=4)
        self.combo_campo.bind("<<ComboboxSelected>>", self._ao_trocar_campo)

        self.btn_limpar = tk.Button(
            row_search,
            text="Limpar",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
            activebackground=COLORS["border"],
            relief=tk.FLAT,
            cursor="hand2",
            command=self._ao_limpar,
        )
        self.btn_limpar.pack(side=tk.LEFT, padx=(8, 0), pady=4)

    def _ao_digitar(self, *args):
        if self._on_busca:
            self._on_busca()

    def _ao_trocar_campo(self, event=None):
        if self._on_busca:
            self._on_busca()

    def _ao_limpar(self):
        self.var_termo.set("")
        self.var_campo.set("todos")
        self.entry_busca.focus_set()
        if self._on_limpar:
            self._on_limpar()

    def focus_busca(self):
        """Coloca o foco no campo de pesquisa."""
        self.entry_busca.focus_set()

    def get_termo(self):
        return self.var_termo.get().strip()

    def get_campo(self):
        return self.var_campo.get()
