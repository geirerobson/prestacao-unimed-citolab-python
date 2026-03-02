"""
Tabela de fichas (Treeview) para exibição dos dados da tabela labplus.
"""

import tkinter as tk
from tkinter import ttk

import pandas as pd

from config import (
    COLORS,
    CAMPOS_TABELA,
    CABECALHOS_TABELA,
    LARGURAS_TABELA,
    COLUNAS_ALINHAMENTO_DIREITA,
    CHECK_SIMBOLO,
)


class TabelaLabPlus(tk.Frame):
    """Frame com Treeview e scrollbars para exibir dados da tabela labplus."""

    def __init__(self, parent, **kwargs):
        kwargs.setdefault("bg", COLORS["card_bg"])
        super().__init__(parent, **kwargs)
        self._construir()

    def _construir(self):
        scroll_y = ttk.Scrollbar(self, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(self, orient=tk.HORIZONTAL)

        colunas = tuple(CAMPOS_TABELA)
        self.tree = ttk.Treeview(
            self,
            columns=colunas,
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
        )

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        for campo in CAMPOS_TABELA:
            self.tree.heading(campo, text=CABECALHOS_TABELA.get(campo, campo))
            largura = LARGURAS_TABELA.get(campo, 100)
            if campo == "checked":
                anchor = tk.CENTER
            elif campo in COLUNAS_ALINHAMENTO_DIREITA:
                anchor = tk.E
            else:
                anchor = tk.W
            self.tree.column(campo, width=largura, anchor=anchor)

        # Grid: linha 0 = tree + scroll vertical; linha 1 = scroll horizontal
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

    def _valor_checked(self, valor) -> str:
        """Exibe o símbolo de check quando checked é verdadeiro."""
        if valor is None or (isinstance(valor, str) and not valor.strip()):
            return ""
        if isinstance(valor, str):
            if valor.strip().lower() in ("1", "true", "sim", "s", "x", "✓", "✔"):
                return CHECK_SIMBOLO
            return ""
        if valor:
            return CHECK_SIMBOLO
        return ""

    def atualizar(self, df: pd.DataFrame):
        """Remove todos os itens da árvore e preenche com as linhas do DataFrame."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in df.iterrows():
            valores = []
            for campo in CAMPOS_TABELA:
                if campo == "checked":
                    valores.append(self._valor_checked(row.get(campo)))
                else:
                    valores.append(row.get(campo, ""))
            self.tree.insert("", tk.END, values=tuple(valores))
