"""
Prestação de Contas Unimed - Citolab - Interface para prestação de contas.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import queue
import threading
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer

from thefuzz import fuzz, process
import unicodedata
import re

from config import (
    COLORS,
    FONT_TITLE,
    FONT_FAMILY,
    FONT_NORMAL,
    FONT_SMALL,
    OPCOES_FILTRO_BUSCA_CITOLAB,
    CAMPOS_TABELA,
    COLUNAS_EXIBICAO_CITOLAB,
    CABECALHOS_CITOLAB_EXCEL,
    LARGURAS_CITOLAB_EXCEL,
    COLUNAS_ALINHAMENTO_DIREITA_CITOLAB,
    PERCENTUAL_SIMILARIDADE_MINIMA,
)
from card_busca import CardBusca
from card_competencia import CardCompetencia
from tabela_labplus import TabelaLabPlus
from tabela_citolab_excel import TabelaCitolabExcel
from init import (
    carregar_dados_db,
    carregar_citolab_excel_db,
    carregar_citolab_excel_db_com_id,
    atualizar_proximidade_labplus,
    marcar_checked_labplus,
    desmarcar_checked_labplus,
    inicializar_banco,
)
from importar_labplus import importar_xml_labplus
from importar_excel_citolab import importar_excel_citolab


def _apenas_data(valor, campo: str) -> str:
    """Para campos de data (data_hora_atendimento, data_atendimento), retorna apenas a data (YYYY-MM-DD)."""
    if campo not in ("data_hora_atendimento", "data_atendimento"):
        return str(valor).strip() if pd.notna(valor) else ""
    s = str(valor).strip()
    if not s or s.lower() in ("nan", "nat"):
        return ""
    # YYYY-MM-DD são os primeiros 10 caracteres (com ou sem hora)
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    return s


def filtrar_por_competencia(
    df: pd.DataFrame, exercicio: str, competencia: str
) -> pd.DataFrame:
    """Filtra pelo exercício (ano) e competência (mês) usando a coluna de data de atendimento."""
    if not exercicio and not competencia:
        return df
    if df.empty:
        return df
    # LabPlus usa data_hora_atendimento; Citolab Excel usa data_atendimento
    if "data_hora_atendimento" in df.columns:
        col_data = "data_hora_atendimento"
    else:
        col_data = "data_atendimento"
    if col_data not in df.columns:
        return df

    # Extrair ano e mês da string (YYYY-MM-DD ou YYYY-MM-DDTHH:MM:SS)
    datas = df[col_data].astype(str).str.strip()
    anos_str = datas.str[:4]
    meses_str = datas.str[5:7]

    mask = pd.Series(True, index=df.index)
    if exercicio:
        try:
            ano_ref = int(exercicio)
            mask &= anos_str.eq(str(ano_ref))
        except ValueError:
            pass
    if competencia:
        try:
            mes_ref = int(competencia)
            mask &= meses_str.eq(f"{mes_ref:02d}")
        except ValueError:
            pass
    return df[mask]


def filtrar_dados(df: pd.DataFrame, termo: str, campo: str) -> pd.DataFrame:
    """
    Filtra linhas onde o termo aparece no(s) campo(s) indicado(s).
    """
    termo = (termo or "").strip().lower()
    if not termo:
        return df
    if df.empty:
        return df

    colunas_busca = [
        "numero",
        "numero_carteira",
        "nome_beneficiario",
        "dados_solicitante_nome",
        "prestador_executante_nome",
        "data_hora_atendimento",
        "procedimento_codigo",
        "procedimento_descricao",
    ]

    if campo == "todos":
        mask = pd.Series(False, index=df.index)
        for col in colunas_busca:
            if col in df.columns:
                mask |= df[col].astype(str).str.lower().str.contains(termo, na=False)
    else:
        # Filtro por campo específico: garantir que a coluna existe
        if campo not in df.columns:
            return df
        mask = df[campo].astype(str).str.lower().str.contains(termo, na=False)

    return df[mask]


def filtrar_dados_excel(df: pd.DataFrame, termo: str, campo: str) -> pd.DataFrame:
    """
    Filtra linhas da tabela Citolab Excel onde o termo aparece no(s) campo(s) indicado(s).
    Colunas: id, data_atendimento, nome_prestador, servico, num_nota, soma, descricao, nome_beneficiario.
    """
    termo = (termo or "").strip().lower()
    if not termo:
        return df
    if df.empty:
        return df

    colunas_busca = [
        "id",
        "data_atendimento",
        "nome_prestador",
        "servico",
        "num_nota",
        "soma",
        "descricao",
        "nome_beneficiario",
    ]

    if campo == "todos":
        mask = pd.Series(False, index=df.index)
        for col in colunas_busca:
            if col in df.columns:
                mask |= df[col].astype(str).str.lower().str.contains(termo, na=False)
    else:
        if campo not in df.columns:
            return df
        mask = df[campo].astype(str).str.lower().str.contains(termo, na=False)

    return df[mask]


def filtrar_excel_por_data_atendimento(df: pd.DataFrame, data: str) -> pd.DataFrame:
    """
    Filtra os registros do DataFrame onde a data (parte YYYY-MM-DD) de 'data_atendimento'
    é igual à data fornecida. Compara apenas o dia, ignorando hora.

    Parâmetros:
        df (pd.DataFrame): DataFrame a ser filtrado (espera-se citolab_excel).
        data (str): Data no formato YYYY-MM-DD.

    Retorna:
        pd.DataFrame: DataFrame filtrado (ou vazio se data vazia / sem coluna).
    """
    if df.empty:
        return pd.DataFrame()
    if "data_atendimento" not in df.columns:
        return pd.DataFrame()
    data = (data or "").strip()
    if not data:
        return pd.DataFrame()
    datas_normalizadas = (
        df["data_atendimento"]
        .astype(str)
        .str.strip()
        .apply(lambda s: _apenas_data(s, "data_atendimento"))
    )
    return df[datas_normalizadas == data]


def normalizar_nome(nome):
    # Remove acentos, normaliza case, remove pontuação/espaços extras
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("ASCII")
    nome = re.sub(r"[^\w\s]", "", nome)  # Remove pontuação
    return " ".join(nome.split()).upper()  # Espaços únicos, maiúscula


def normalizar_valor(valor_str):
    # Remove R$, vírgulas, espaços; converte vírgula em ponto
    valor_limpo = re.sub(r"[^\d,.]", "", valor_str)
    if "," in valor_limpo:
        valor_limpo = valor_limpo.replace(",", ".")
    return valor_limpo


def _similaridade_uma_linha(
    nome_labplus: str,
    desc_labplus: str,
    nome_excel: str,
    desc_excel: str,
) -> float:
    """Calcula similaridade (0-100) entre um registro LabPlus e uma linha Excel (nome + descrição)."""
    nome_l = normalizar_nome(str(nome_labplus or ""))
    nome_e = normalizar_nome(str(nome_excel or ""))
    desc_l = normalizar_nome(str(desc_labplus or ""))
    desc_e = normalizar_nome(str(desc_excel or ""))
    sim_nome = fuzz.ratio(nome_l, nome_e)
    sim_desc = fuzz.ratio(desc_l, desc_e)
    return (sim_nome + sim_desc) / 2.0


def calcular_proximidade(
    nome_beneficiario: str,
    procedimento_descricao: str,
    procedimento_valor: float,
    df_citolab_excel: pd.DataFrame,
) -> float:
    """Calcula a proximidade entre um registro LabPlus e o primeiro registro Excel (legado)."""
    if df_citolab_excel.empty:
        return 0.0
    return _similaridade_uma_linha(
        nome_beneficiario,
        procedimento_descricao,
        df_citolab_excel["nome_beneficiario"].iloc[0],
        df_citolab_excel["descricao"].iloc[0],
    )


def calcular_melhor_proximidade(
    nome_beneficiario: str,
    procedimento_descricao: str,
    procedimento_valor: float,
    df_citolab_excel: pd.DataFrame,
) -> tuple[float, int | None]:
    """
    Para um registro LabPlus, percorre todos os registros do Excel (mesma data) e retorna
    a maior similaridade e o id do registro Citolab Excel que obteve essa similaridade.
    Retorno: (similaridade 0-100, id_citolab ou None se df vazio).
    """
    if df_citolab_excel.empty:
        return (0.0, None)
    melhor_sim = 0.0
    id_melhor = None
    for _, ex_row in df_citolab_excel.iterrows():
        sim = _similaridade_uma_linha(
            nome_beneficiario,
            procedimento_descricao,
            ex_row.get("nome_beneficiario", "") or "",
            ex_row.get("descricao", "") or "",
        )
        if sim > melhor_sim:
            melhor_sim = sim
            try:
                id_melhor = int(ex_row["id"])
            except (TypeError, ValueError, KeyError):
                id_melhor = None
    return (melhor_sim, id_melhor)


class AppPrestacaoContas(tk.Tk):
    def __init__(self):
        super().__init__()
        self.df = carregar_dados_db()
        self.df_citolab_excel = carregar_citolab_excel_db()
        self.configurar_janela()
        self.criar_menu()
        self.criar_estilos()
        self.criar_interface()

    def configurar_janela(self):
        self.title("Prestação de Contas Unimed — Citolab")
        self.geometry("960x620")
        self.minsize(720, 480)
        self.configure(bg=COLORS["bg"])

    def criar_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        menu_arquivo = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)
        menu_arquivo.add_command(
            label="Importar XML (LabPlus)...", command=self._importar_xml
        )
        menu_arquivo.add_command(
            label="Importar Excel (CITOLAB)...", command=self._importar_excel_citolab
        )
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.quit)

    def criar_estilos(self):
        self.option_add("*Font", (FONT_FAMILY, 10))
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("Card.TFrame", background=COLORS["card_bg"])
        self.style.configure(
            "Treeview",
            background=COLORS["card_bg"],
            foreground=COLORS["text"],
            fieldbackground=COLORS["card_bg"],
            rowheight=28,
            font=FONT_NORMAL,
        )
        self.style.configure(
            "Treeview.Heading",
            background=COLORS["header_bg"],
            foreground="white",
            font=(FONT_FAMILY, 10, "bold"),
        )
        self.style.map("Treeview", background=[("selected", COLORS["accent"])])
        self.style.map(
            "Treeview.Heading", background=[("active", COLORS["primary_light"])]
        )

    def _desenhar_sombra(self, widget):
        """Aplica borda sutil para efeito de card."""
        widget.configure(highlightbackground=COLORS["border"], highlightthickness=1)

    def criar_interface(self):
        main = tk.Frame(self, bg=COLORS["bg"], padx=24, pady=20)
        main.pack(fill=tk.BOTH, expand=True)

        titulo = tk.Label(
            main,
            text="Prestação de Contas Unimed - Citolab",
            font=FONT_TITLE,
            fg=COLORS["primary"],
            bg=COLORS["bg"],
        )
        titulo.pack(anchor=tk.W, pady=(0, 16))

        # Abas: LabPlus e Citolab Excel
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Botão Verificar (proximidade labplus x citolab_excel) e gauge de andamento
        frame_verificar = tk.Frame(main, bg=COLORS["bg"])
        frame_verificar.pack(fill=tk.X, pady=(12, 0))
        self._frm_gauge_verificar = tk.Frame(frame_verificar, bg=COLORS["bg"])
        self._lbl_verificando = tk.Label(
            self._frm_gauge_verificar,
            text="Verificando...",
            font=FONT_SMALL,
            fg=COLORS["text_secondary"],
            bg=COLORS["bg"],
        )
        self._gauge_verificar = ttk.Progressbar(
            self._frm_gauge_verificar,
            mode="determinate",
            length=200,
        )
        self._lbl_verificando.pack(side=tk.LEFT, padx=(0, 6))
        self._gauge_verificar.pack(side=tk.LEFT)
        self._frm_gauge_verificar.pack(side=tk.LEFT, padx=(0, 8))
        self._frm_gauge_verificar.pack_forget()
        self.btn_verificar = tk.Button(
            frame_verificar,
            text="Verificar",
            font=FONT_NORMAL,
            bg=COLORS["accent"],
            fg="black",
            activebackground=COLORS["accent_hover"],
            activeforeground="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._verificar_proximidade,
        )
        self.btn_verificar.pack(side=tk.LEFT)

        # ---- Aba LabPlus ----
        tab_labplus = tk.Frame(self.notebook, bg=COLORS["bg"])
        self.notebook.add(tab_labplus, text="  LabPlus (XML)  ")

        self.card_busca = CardBusca(
            tab_labplus,
            on_busca=self.executar_busca,
            on_limpar=self.executar_busca,
        )
        self.card_busca.pack(fill=tk.X, pady=(0, 8))
        self._desenhar_sombra(self.card_busca)
        self.card_busca.focus_busca()

        self.card_competencia = CardCompetencia(
            tab_labplus,
            on_mudar=self.executar_busca,
            exercicio_como_input=True,
        )
        self.card_competencia.pack(fill=tk.X, pady=(0, 16))
        self._desenhar_sombra(self.card_competencia)

        self.tabela_labplus = TabelaLabPlus(tab_labplus)
        self.tabela_labplus.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self._desenhar_sombra(self.tabela_labplus)
        self.tabela_labplus.tree.bind("<Double-1>", self._ao_double_click_labplus)
        self.tabela_labplus.tree.bind("<Button-3>", self._ao_clique_direito_labplus)

        self.status = tk.Label(
            tab_labplus,
            text="",
            font=FONT_SMALL,
            fg=COLORS["text_secondary"],
            bg=COLORS["bg"],
        )
        self.status.pack(anchor=tk.W)

        # ---- Aba Citolab Excel ----
        tab_excel = tk.Frame(self.notebook, bg=COLORS["bg"])
        self.notebook.add(tab_excel, text="  Citolab Excel  ")

        self.card_busca_excel = CardBusca(
            tab_excel,
            on_busca=self.executar_busca_excel,
            on_limpar=self.executar_busca_excel,
            opcoes_campo=OPCOES_FILTRO_BUSCA_CITOLAB,
        )
        self.card_busca_excel.pack(fill=tk.X, pady=(0, 8))
        self._desenhar_sombra(self.card_busca_excel)

        self.card_competencia_excel = CardCompetencia(
            tab_excel,
            on_mudar=self.executar_busca_excel,
        )
        self.card_competencia_excel.pack(fill=tk.X, pady=(0, 16))
        self._desenhar_sombra(self.card_competencia_excel)

        lbl_excel = tk.Label(
            tab_excel,
            text="Relatório Excel (CITOLAB) — importe via Arquivo → Importar Excel (CITOLAB)...",
            font=FONT_SMALL,
            fg=COLORS["text_secondary"],
            bg=COLORS["bg"],
        )
        lbl_excel.pack(anchor=tk.W, pady=(0, 8))

        self.tabela_citolab_excel = TabelaCitolabExcel(tab_excel)
        self.tabela_citolab_excel.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self._desenhar_sombra(self.tabela_citolab_excel)

        self.status_excel = tk.Label(
            tab_excel,
            text="",
            font=FONT_SMALL,
            fg=COLORS["text_secondary"],
            bg=COLORS["bg"],
        )
        self.status_excel.pack(anchor=tk.W)

        self.executar_busca()
        self.executar_busca_excel()

    def executar_busca(self):
        """Executa a busca e atualiza a tabela LabPlus (texto + competência/exercício).
        Só exibe registros quando Exercício e Competência estiverem definidos."""
        termo = self.card_busca.get_termo()
        campo = self.card_busca.get_campo()
        exercicio = self.card_competencia.get_exercicio()
        competencia = self.card_competencia.get_competencia()

        if not exercicio or not competencia:
            df_filtrado = self.df.iloc[0:0].copy()
            self.tabela_labplus.atualizar(df_filtrado)
            self.status.config(
                text="Defina Exercício e Competência para exibir registros."
            )
            return

        df_filtrado = filtrar_dados(self.df, termo, campo)
        df_filtrado = filtrar_por_competencia(df_filtrado, exercicio, competencia)
        self.tabela_labplus.atualizar(df_filtrado)

        n = len(df_filtrado)
        total_periodo = len(filtrar_por_competencia(self.df, exercicio, competencia))
        if termo:
            self.status.config(
                text=f"{n} registro(s) encontrado(s) (total no período: {total_periodo})"
            )
        else:
            self.status.config(
                text=f"{n} registro(s) no período (exercício {exercicio}, competência {competencia})"
            )

    def executar_busca_excel(self):
        """Aplica filtros opcionais (termo, campo, exercício, competência) e atualiza a tabela Citolab Excel.
        Exibe todos os registros por padrão; filtros são opcionais."""
        termo = self.card_busca_excel.get_termo()
        campo = self.card_busca_excel.get_campo()
        exercicio = self.card_competencia_excel.get_exercicio()
        competencia = self.card_competencia_excel.get_competencia()

        df_filtrado = filtrar_dados_excel(self.df_citolab_excel, termo, campo)
        if exercicio and competencia:
            df_filtrado = filtrar_por_competencia(df_filtrado, exercicio, competencia)
        self.tabela_citolab_excel.atualizar(df_filtrado)

        n = len(df_filtrado)
        total = len(self.df_citolab_excel)
        if termo or (exercicio and competencia):
            self.status_excel.config(
                text=f"{n} registro(s) encontrado(s) (total: {total})"
            )
        else:
            self.status_excel.config(text=f"Exibindo todos os registros: {total}")

    def atualizar_tab_citolab_excel(self):
        """Recarrega os dados do banco e aplica os filtros atuais na aba Citolab Excel."""
        self.df_citolab_excel = carregar_citolab_excel_db()
        self.executar_busca_excel()

    def _ao_double_click_labplus(self, event):
        """Ao duplo clique em uma linha da tabela LabPlus, abre modal com Citolab Excel filtrado pela data."""
        sel = self.tabela_labplus.tree.selection()
        if not sel:
            return
        item = sel[0]
        values = self.tabela_labplus.tree.item(item, "values")
        if not values:
            return
        idx_data = next(
            (i for i, c in enumerate(CAMPOS_TABELA) if c == "data_hora_atendimento"),
            None,
        )
        if idx_data is None or idx_data >= len(values):
            return
        data_str = (values[idx_data] or "").strip()
        data_yyyy_mm_dd = _apenas_data(data_str, "data_hora_atendimento")
        if not data_yyyy_mm_dd:
            messagebox.showinfo(
                "Data",
                "Registro sem data de atendimento.",
                parent=self,
            )
            return
        idx_id = next((i for i, c in enumerate(CAMPOS_TABELA) if c == "id"), None)
        if idx_id is None or idx_id >= len(values):
            return
        try:
            id_labplus = int(values[idx_id])
        except (ValueError, TypeError):
            return
        proximidade_id = None
        idx_proximidade = next(
            (i for i, c in enumerate(CAMPOS_TABELA) if c == "proximidade_id"), None
        )
        if idx_proximidade is not None and idx_proximidade < len(values):
            raw = values[idx_proximidade]
            if raw is not None and str(raw).strip():
                try:
                    proximidade_id = int(raw)
                except (ValueError, TypeError):
                    pass
        self._abrir_modal_citolab_por_data(
            data_yyyy_mm_dd, id_labplus, proximidade_id=proximidade_id
        )

    def _id_labplus_da_selecao(self):
        """Retorna o id da linha LabPlus atualmente selecionada, ou None."""
        sel = self.tabela_labplus.tree.selection()
        if not sel:
            return None
        values = self.tabela_labplus.tree.item(sel[0], "values")
        if not values:
            return None
        idx_id = next((i for i, c in enumerate(CAMPOS_TABELA) if c == "id"), None)
        if idx_id is None or idx_id >= len(values):
            return None
        try:
            return int(values[idx_id])
        except (ValueError, TypeError):
            return None

    def _ao_clique_direito_labplus(self, event):
        """Exibe menu de contexto (Marcar / Desmarcar) na tabela LabPlus."""
        tree = self.tabela_labplus.tree
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
        id_lab = self._id_labplus_da_selecao()
        if id_lab is None:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(
            label="Marcar (checked)",
            command=lambda: self._menu_labplus_marcar(id_lab),
        )
        menu.add_command(
            label="Desmarcar (checked e proximidade_id)",
            command=lambda: self._menu_labplus_desmarcar(id_lab),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _menu_labplus_marcar(self, id_labplus: int):
        """Ação do menu: marcar registro como checked."""
        if marcar_checked_labplus("prestacao_contas.db", id_labplus):
            self.df = carregar_dados_db()
            self.executar_busca()

    def _menu_labplus_desmarcar(self, id_labplus: int):
        """Ação do menu: desmarcar (checked False e proximidade_id = NULL)."""
        if desmarcar_checked_labplus("prestacao_contas.db", id_labplus):
            self.df = carregar_dados_db()
            self.executar_busca()

    def _abrir_modal_citolab_por_data(
        self,
        data_yyyy_mm_dd: str,
        id_labplus: int,
        proximidade_id: int | None = None,
    ):
        """Abre janela modal com a tabela citolab_excel filtrada pela data (apenas YYYY-MM-DD).
        Ao duplo clique em uma row do modal, marca a row de origem (id_labplus) como checked na LabPlus.
        Se proximidade_id for informado, a linha do Citolab Excel com esse id é destacada em amarelo.
        """
        df_excel = carregar_citolab_excel_db()
        if df_excel.empty or "data_atendimento" not in df_excel.columns:
            messagebox.showinfo(
                "Citolab Excel",
                "Nenhum dado na tabela Citolab Excel.",
                parent=self,
            )
            return
        # Filtrar pela data (só a parte YYYY-MM-DD)
        datas_excel = df_excel["data_atendimento"].astype(str).str.strip()
        mask = datas_excel.apply(
            lambda s: _apenas_data(s, "data_atendimento") == data_yyyy_mm_dd
        )
        df_filtrado = df_excel.loc[mask].copy()
        # Ordenar para exibição
        if not df_filtrado.empty and "nome_beneficiario" in df_filtrado.columns:
            df_filtrado = df_filtrado.sort_values("nome_beneficiario").reset_index(
                drop=True
            )

        modal = tk.Toplevel(self)
        modal.title(f"Citolab Excel — Atendimentos em {data_yyyy_mm_dd}")
        modal.geometry("900x400")
        modal.transient(self)
        modal.grab_set()
        modal.configure(bg=COLORS["bg"])

        frm = tk.Frame(modal, bg=COLORS["bg"], padx=8, pady=8)
        frm.pack(fill=tk.BOTH, expand=True)

        lbl = tk.Label(
            frm,
            text=f"Registros com data_atendimento = {data_yyyy_mm_dd} ({len(df_filtrado)} registro(s))",
            font=FONT_NORMAL,
            fg=COLORS["text"],
            bg=COLORS["bg"],
        )
        lbl.pack(anchor=tk.W, pady=(0, 4))

        # Filtro por ID, Paciente e Procedimento
        frm_filtro = tk.Frame(frm, bg=COLORS["bg"])
        frm_filtro.pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            frm_filtro,
            text="Filtrar (ID, Paciente, Procedimento):",
            font=FONT_NORMAL,
            fg=COLORS["text"],
            bg=COLORS["bg"],
        ).pack(side=tk.LEFT, padx=(0, 8))
        var_filtro = tk.StringVar()

        def _aplicar_filtro_modal(*_args):
            termo = (var_filtro.get() or "").strip().lower()
            if not termo:
                df_view = df_filtrado
            else:
                col_id = "id" if "id" in df_filtrado.columns else None
                col_pac = (
                    "nome_beneficiario"
                    if "nome_beneficiario" in df_filtrado.columns
                    else None
                )
                col_proc = "descricao" if "descricao" in df_filtrado.columns else None
                mask = pd.Series(False, index=df_filtrado.index)
                if col_id:
                    mask |= (
                        df_filtrado[col_id]
                        .astype(str)
                        .str.lower()
                        .str.contains(termo, na=False)
                    )
                if col_pac:
                    mask |= (
                        df_filtrado[col_pac]
                        .astype(str)
                        .str.lower()
                        .str.contains(termo, na=False)
                    )
                if col_proc:
                    mask |= (
                        df_filtrado[col_proc]
                        .astype(str)
                        .str.lower()
                        .str.contains(termo, na=False)
                    )
                df_view = df_filtrado.loc[mask]
            for item in tree.get_children():
                tree.delete(item)
            for _, row in df_view.iterrows():
                vals = tuple(
                    str(row[c]) if c in row.index and pd.notna(row[c]) else ""
                    for c in COLUNAS_EXIBICAO_CITOLAB
                )
                row_id = None
                if "id" in row.index and pd.notna(row.get("id")):
                    try:
                        row_id = int(row["id"])
                    except (ValueError, TypeError):
                        pass
                tags = (
                    ("proximidade",)
                    if (proximidade_id is not None and row_id == proximidade_id)
                    else ()
                )
                tree.insert("", tk.END, values=vals, tags=tags)
            n_view, n_total = len(df_view), len(df_filtrado)
            if termo and n_total > 0:
                lbl.config(
                    text=f"Registros com data_atendimento = {data_yyyy_mm_dd} — {n_view} de {n_total} registro(s)"
                )
            else:
                lbl.config(
                    text=f"Registros com data_atendimento = {data_yyyy_mm_dd} ({n_total} registro(s))"
                )

        entry_filtro = tk.Entry(
            frm_filtro,
            textvariable=var_filtro,
            font=FONT_NORMAL,
            width=40,
        )
        entry_filtro.pack(side=tk.LEFT, fill=tk.X, expand=True)
        var_filtro.trace_add("write", _aplicar_filtro_modal)

        # Frame só para a tabela (grid), igual à aba Citolab Excel
        frm_tree = tk.Frame(frm, bg=COLORS["card_bg"])
        frm_tree.pack(fill=tk.BOTH, expand=True)

        scroll_y = ttk.Scrollbar(frm_tree, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(frm_tree, orient=tk.HORIZONTAL)
        colunas = tuple(COLUNAS_EXIBICAO_CITOLAB)
        tree = ttk.Treeview(
            frm_tree,
            columns=colunas,
            show="headings",
            height=12,
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
        )
        scroll_y.config(command=tree.yview)
        scroll_x.config(command=tree.xview)
        for c in colunas:
            tree.heading(c, text=CABECALHOS_CITOLAB_EXCEL.get(c, c))
            w = LARGURAS_CITOLAB_EXCEL.get(c, 100)
            anchor = tk.E if c in COLUNAS_ALINHAMENTO_DIREITA_CITOLAB else tk.W
            tree.column(c, width=w, anchor=anchor)
        tree.tag_configure("proximidade", background="#fff3cd")
        _aplicar_filtro_modal()
        frm_tree.grid_rowconfigure(0, weight=1)
        frm_tree.grid_columnconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        def _ao_double_click_modal(_event, id_origem=id_labplus, win=modal, tr=tree):
            sel = tr.selection()
            id_citolab = None
            if sel:
                vals = tr.item(sel[0], "values")
                if vals and len(vals) > 0 and str(vals[0]).strip():
                    try:
                        id_citolab = int(vals[0])  # primeira coluna = id citolab_excel
                    except (ValueError, TypeError):
                        pass
            if marcar_checked_labplus(
                "prestacao_contas.db", id_origem, id_citolab_excel=id_citolab
            ):
                self.df = carregar_dados_db()
                self.executar_busca()
            win.destroy()

        tree.bind("<Double-1>", _ao_double_click_modal)

        btn = tk.Button(
            frm,
            text="Fechar",
            font=FONT_NORMAL,
            bg=COLORS["accent"],
            fg="white",
            activebackground=COLORS["accent_hover"],
            relief=tk.FLAT,
            cursor="hand2",
            command=modal.destroy,
        )
        btn.pack(pady=(12, 0))

        modal.protocol("WM_DELETE_WINDOW", modal.destroy)
        modal.focus_force()

    def _verificar_proximidade(self):
        """
        Dispara o cálculo de proximidade em thread e exibe o gauge enquanto executa.
        Validação (exercício/competência e dados) é feita na thread principal; o cálculo pesado roda em background.
        """
        exercicio = self.card_competencia.get_exercicio()
        competencia = self.card_competencia.get_competencia()
        if not exercicio or not competencia:
            messagebox.showwarning(
                "Verificar",
                "É necessário informar Exercício e Competência (filtros da aba LabPlus/XML) para calcular a proximidade.",
                parent=self,
            )
            return

        df_labplus_filtrado = filtrar_por_competencia(self.df, exercicio, competencia)
        if df_labplus_filtrado.empty:
            messagebox.showwarning(
                "Verificar",
                "Nenhum registro na tabela LabPlus para o exercício e competência selecionados.",
                parent=self,
            )
            return

        result_queue = queue.Queue()
        total_linhas = len(df_labplus_filtrado)

        def _calculo():
            try:
                datas_unicas = set()
                for _, r in df_labplus_filtrado.iterrows():
                    d = _apenas_data(r.get("data_hora_atendimento"), "data_hora_atendimento")
                    if d:
                        datas_unicas.add(d)
                cache_excel_por_data = {
                    data: filtrar_excel_por_data_atendimento(self.df_citolab_excel, data)
                    for data in datas_unicas
                }
                id_valores = []
                id_citolab_por_labplus = []
                for i, (_, row) in enumerate(df_labplus_filtrado.iterrows()):
                    result_queue.put(("progress", i + 1, total_linhas))
                    id_labplus = int(row["id"])
                    data_labplus = _apenas_data(
                        row.get("data_hora_atendimento"), "data_hora_atendimento"
                    )
                    nome_beneficiario = row.get("nome_beneficiario", "") or ""
                    procedimento_descricao = row.get("procedimento_descricao", "") or ""
                    procedimento_valor = row.get("procedimento_valor", "") or ""
                    df_excel = (
                        cache_excel_por_data.get(data_labplus, pd.DataFrame())
                        if data_labplus
                        else pd.DataFrame()
                    )
                    if df_excel.empty:
                        id_valores.append((id_labplus, "0.0000"))
                        id_citolab_por_labplus.append(None)
                        continue
                    proximidade, id_citolab = calcular_melhor_proximidade(
                        nome_beneficiario,
                        procedimento_descricao,
                        procedimento_valor,
                        df_excel,
                    )
                    id_valores.append((id_labplus, f"{proximidade:.4f}"))
                    id_citolab_por_labplus.append(
                        id_citolab if proximidade > PERCENTUAL_SIMILARIDADE_MINIMA else None
                    )
                result_queue.put(("done", True, None, id_valores, id_citolab_por_labplus))
            except Exception as e:
                result_queue.put(("done", False, str(e), None, None))

        self._frm_gauge_verificar.pack(side=tk.LEFT, padx=(0, 8))
        self._gauge_verificar["maximum"] = max(total_linhas, 1)
        self._gauge_verificar["value"] = 0
        self.btn_verificar.config(state=tk.DISABLED)
        self.update_idletasks()
        threading.Thread(target=_calculo, daemon=True).start()

        def _check_queue():
            try:
                while True:
                    msg = result_queue.get_nowait()
                    if msg[0] == "progress":
                        _, current, total = msg
                        self._gauge_verificar["value"] = current
                        self.update_idletasks()
                    elif msg[0] == "done":
                        _, ok, err_msg, id_valores, id_citolab_por_labplus = msg
                        break
            except queue.Empty:
                self.after(80, _check_queue)
                return

            def _esconder_gauge_e_reativar():
                self._gauge_verificar["value"] = 0
                self._frm_gauge_verificar.pack_forget()
                self.btn_verificar.config(state=tk.NORMAL)
                self.update_idletasks()

            if not ok:
                messagebox.showerror(
                    "Verificar",
                    f"Erro ao calcular proximidade: {err_msg}",
                    parent=self,
                )
                _esconder_gauge_e_reativar()
                return
            if id_valores:
                if not atualizar_proximidade_labplus(
                    "prestacao_contas.db", id_valores, id_citolab_por_labplus
                ):
                    messagebox.showerror(
                        "Verificar",
                        "Erro ao gravar proximidade no banco de dados.",
                        parent=self,
                    )
                    _esconder_gauge_e_reativar()
                    return
            self.df = carregar_dados_db()
            self.executar_busca()
            messagebox.showinfo(
                "Proximidade calculada",
                "Proximidade calculada para todos os registros da tabela LabPlus.",
                parent=self,
            )
            _esconder_gauge_e_reativar()

        self.after(120, _check_queue)

    def _importar_xml(self):
        """Abre diálogo para escolher arquivo XML e importa para o banco."""
        caminho = filedialog.askopenfilename(
            title="Importar XML LabPlus",
            filetypes=[
                ("Arquivos XML", "*.xml"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not caminho:
            return
        inseridos, duplicados, erro = importar_xml_labplus(caminho)
        if erro:
            messagebox.showerror("Erro na importação", erro)
            return
        msg = f"Foram importadas {inseridos} guia(s) com sucesso."
        if duplicados:
            msg += f"\n{duplicados} guia(s) já existente(s) foram ignorada(s) (duplicadas)."
        messagebox.showinfo("Importação concluída", msg)
        self.df = carregar_dados_db()
        self.executar_busca()

    def _importar_excel_citolab(self):
        """Abre diálogo para escolher arquivo Excel CITOLAB e importa para a tabela citolab_excel."""
        caminho = filedialog.askopenfilename(
            title="Importar Excel CITOLAB",
            filetypes=[
                ("Arquivos Excel", "*.xls *.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Excel", "*.xlsx"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not caminho:
            return
        inseridos, duplicados, erro = importar_excel_citolab(caminho, substituir=True)
        if erro:
            messagebox.showerror("Erro na importação", erro)
            return
        msg = f"Foram importados {inseridos} registro(s) na tabela citolab_excel."
        if duplicados:
            msg += f"\n{duplicados} registro(s) já existente(s) foram ignorado(s) (duplicatas)."
        messagebox.showinfo("Importação concluída", msg)
        self.atualizar_tab_citolab_excel()


def main():
    inicializar_banco()
    app = AppPrestacaoContas()
    app.mainloop()


if __name__ == "__main__":
    main()
