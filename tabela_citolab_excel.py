"""
Tabela (Treeview) para exibição dos dados da tabela citolab_excel (relatório Excel CITOLAB).
"""

import tkinter as tk
from tkinter import ttk

import pandas as pd

from config import (
    COLORS,
    COLUNAS_EXIBICAO_CITOLAB,
    CABECALHOS_CITOLAB_EXCEL,
    LARGURAS_CITOLAB_EXCEL,
    COLUNAS_ALINHAMENTO_DIREITA_CITOLAB,
)


class TabelaCitolabExcel(tk.Frame):
    """Frame com Treeview e scrollbars para exibir dados da tabela citolab_excel."""

    def __init__(self, parent, **kwargs):
        kwargs.setdefault("bg", COLORS["card_bg"])
        super().__init__(parent, **kwargs)
        self._construir()

    def _construir(self):
        scroll_y = ttk.Scrollbar(self, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(self, orient=tk.HORIZONTAL)

        colunas = COLUNAS_EXIBICAO_CITOLAB
        self.tree = ttk.Treeview(
            self,
            columns=colunas,
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
        )

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        for campo in colunas:
            self.tree.heading(
                campo, text=CABECALHOS_CITOLAB_EXCEL.get(campo, campo)
            )
            largura = LARGURAS_CITOLAB_EXCEL.get(campo, 100)
            anchor = (
                tk.E
                if campo in COLUNAS_ALINHAMENTO_DIREITA_CITOLAB
                else tk.W
            )
            self.tree.column(campo, width=largura, anchor=anchor)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

    def atualizar(self, df: pd.DataFrame):
        """Remove todos os itens e preenche com as linhas do DataFrame."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in df.iterrows():
            valores = tuple(row.get(c, "") for c in COLUNAS_EXIBICAO_CITOLAB)
            self.tree.insert("", tk.END, values=valores)
