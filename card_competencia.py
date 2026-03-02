"""
Card de filtro por competência (mês) e exercício (ano).
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime

from config import COLORS, FONT_NORMAL, MESES_NOMES


class CardCompetencia(tk.Frame):
    """Card com filtro de Exercício (ano) e Competência (mês)."""

    def __init__(self, parent, on_mudar=None, exercicio_como_input=False, **kwargs):
        kwargs.setdefault("bg", COLORS["card_bg"])
        super().__init__(parent, **kwargs)
        self._on_mudar = on_mudar
        self._exercicio_como_input = exercicio_como_input
        self._construir()

    def _construir(self):
        row = tk.Frame(self, bg=COLORS["card_bg"], padx=20, pady=12)
        row.pack(fill=tk.X)

        tk.Label(
            row,
            text="Exercício (ano):",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
        ).pack(side=tk.LEFT, padx=(0, 8), pady=4)

        ano_atual = datetime.now().year
        self.var_exercicio = tk.StringVar(value=str(ano_atual) if self._exercicio_como_input else "")
        if self._exercicio_como_input:
            self.entry_exercicio = tk.Entry(
                row,
                textvariable=self.var_exercicio,
                width=8,
                font=FONT_NORMAL,
            )
            self.entry_exercicio.pack(side=tk.LEFT, padx=(0, 16), pady=4)
            self.entry_exercicio.bind("<KeyRelease>", self._ao_mudar)
            self.entry_exercicio.bind("<FocusOut>", self._ao_mudar)
        else:
            anos = [""] + [str(a) for a in range(ano_atual - 2, ano_atual + 2)]
            self.combo_exercicio = ttk.Combobox(
                row,
                textvariable=self.var_exercicio,
                values=anos,
                state="readonly",
                width=8,
                font=FONT_NORMAL,
            )
            self.combo_exercicio.pack(side=tk.LEFT, padx=(0, 16), pady=4)
            self.combo_exercicio.bind("<<ComboboxSelected>>", self._ao_mudar)

        tk.Label(
            row,
            text="Competência (mês):",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
        ).pack(side=tk.LEFT, padx=(16, 8), pady=4)

        # Opções: vazio (todos) + 01 a 12 com nome do mês
        opcoes_mes = [""] + [f"{i:02d} - {nome}" for i, nome in enumerate(MESES_NOMES, 1)]
        self.var_competencia = tk.StringVar(value="")
        self.combo_competencia = ttk.Combobox(
            row,
            textvariable=self.var_competencia,
            values=opcoes_mes,
            state="readonly",
            width=18,
            font=FONT_NORMAL,
        )
        self.combo_competencia.pack(side=tk.LEFT, padx=(0, 12), pady=4)
        self.combo_competencia.bind("<<ComboboxSelected>>", self._ao_mudar)

        self.btn_limpar = tk.Button(
            row,
            text="Limpar filtro",
            font=FONT_NORMAL,
            fg=COLORS["text_secondary"],
            bg=COLORS["card_bg"],
            activebackground=COLORS["border"],
            relief=tk.FLAT,
            cursor="hand2",
            command=self._ao_limpar,
        )
        self.btn_limpar.pack(side=tk.LEFT, padx=(12, 0), pady=4)

    def _ao_mudar(self, event=None):
        if self._on_mudar:
            self._on_mudar()

    def _ao_limpar(self):
        self.var_exercicio.set("")
        self.var_competencia.set("")
        if self._on_mudar:
            self._on_mudar()

    def get_exercicio(self):
        """Retorna o ano selecionado (ex: '2025') ou '' para todos."""
        return self.var_exercicio.get().strip()

    def get_competencia(self):
        """Retorna o número do mês (01 a 12) ou '' para todos."""
        val = self.var_competencia.get().strip()
        if not val:
            return ""
        # Formato "01 - Janeiro" -> "01"
        if " - " in val:
            return val.split(" - ")[0].strip()
        return val
