import sqlite3
from sqlite3 import Error
import pandas as pd

from config import CAMPOS_TABELA


def create_connection(db_file):
    """Cria uma conexão com o banco de dados SQLite"""
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
    return conn


def create_table(conn, create_table_sql):
    """Cria uma tabela no banco de dados"""
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
        conn.commit()
    except Error as e:
        print(e)


# Colunas adicionais da tabela labplus (uma row por procedimento)
COLUNAS_NOVAS_LABPLUS = [
    "nome_beneficiario",
    "dados_solicitante_nome",
    "data_hora_atendimento",
    "procedimento_codigo",
    "procedimento_descricao",
    "procedimento_data",
    "procedimento_valor",
    "procedimento_valor_total",
    "proximidade",
    "proximidade_id",
    "checked",
]


def _colunas_existentes(conn):
    """Retorna o conjunto de nomes de colunas da tabela labplus."""
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(labplus)")
    return {row[1] for row in cur.fetchall()}


def _migrar_colunas_labplus(conn):
    """Adiciona colunas novas à tabela labplus se não existirem (migração)."""
    existentes = _colunas_existentes(conn)
    cur = conn.cursor()
    for col in COLUNAS_NOVAS_LABPLUS:
        if col not in existentes:
            cur.execute(f"ALTER TABLE labplus ADD COLUMN {col} TEXT")
            existentes.add(col)
    conn.commit()


def _remover_coluna_observacao(conn):
    """Remove a coluna observacao da tabela labplus se existir (SQLite 3.35+)."""
    if "observacao" not in _colunas_existentes(conn):
        return
    try:
        conn.cursor().execute("ALTER TABLE labplus DROP COLUMN observacao")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # SQLite antigo não suporta DROP COLUMN


def inicializar_banco():
    """Inicializa o banco de dados e cria a tabela se não existir"""
    database = "prestacao_contas.db"
    sql_create_table = """
    CREATE TABLE IF NOT EXISTS labplus (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero TEXT NOT NULL,
        data_emissao TEXT,
        registro_ans TEXT,
        paciente TEXT NOT NULL,
        numero_carteira TEXT,
        nome_plano TEXT,
        profissional TEXT NOT NULL,
        profissional_cpf TEXT,
        conselho_sigla TEXT,
        conselho_numero TEXT,
        conselho_uf TEXT,
        cbos TEXT,
        data_atendimento TEXT,
        valor_total TEXT,
        celular TEXT,
        identificador_beneficiario TEXT,
        indicacao_clinica TEXT,
        carater_atendimento TEXT,
        tipo_saida TEXT,
        tipo_atendimento TEXT,
        prestador_executante_nome TEXT,
        prestador_executante_codigo TEXT,
        prestador_executante_cnes TEXT,
        prestador_executante_endereco TEXT,
        executante_nome TEXT,
        executante_conselho_sigla TEXT,
        executante_conselho_numero TEXT,
        executante_conselho_uf TEXT,
        executante_cbos TEXT,
        procedimentos_descricao TEXT
    );
    """
    # Tabela para relatório Excel CITOLAB (ex.: CITOLAB - OUT-2025.XLS)
    sql_citolab_excel = """
    CREATE TABLE IF NOT EXISTS citolab_excel (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_empresa TEXT,
        data_atendimento TEXT,
        ano TEXT,
        nome_prestador TEXT,
        cod_familia TEXT,
        desc_nota TEXT,
        cod_prestador TEXT,
        servico TEXT,
        cod_unimed TEXT,
        qtde_pago_cobr TEXT,
        num_nota TEXT,
        soma TEXT,
        descricao TEXT,
        mes TEXT,
        nome_beneficiario TEXT
    );
    """
    conn = create_connection(database)
    if conn is not None:
        create_table(conn, sql_create_table)
        _migrar_colunas_labplus(conn)
        _remover_coluna_observacao(conn)
        create_table(conn, sql_citolab_excel)
        conn.close()
    else:
        print("Erro ao criar a conexão com o banco de dados")


def carregar_dados_db(database="prestacao_contas.db"):
    """Carrega dados do banco SQLite"""
    try:
        conn = create_connection(database)
        if conn is None:
            return pd.DataFrame(columns=CAMPOS_TABELA)

        query = f"SELECT {', '.join(CAMPOS_TABELA)} FROM labplus ORDER BY nome_beneficiario, numero, procedimento_codigo"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df.fillna("")
    except Exception as e:
        print(f"Erro ao carregar dados: {e}")
        return pd.DataFrame(columns=CAMPOS_TABELA)


# Colunas da tabela citolab_excel (mesma ordem do importar_excel_citolab.COLUNAS_BANCO)
CITOLAB_EXCEL_COLUNAS = [
    "data_atendimento",
    "nome_prestador",
    "servico",
    "num_nota",
    "soma",
    "descricao",
    "nome_beneficiario",
]


def carregar_citolab_excel_db(database="prestacao_contas.db"):
    """Carrega dados da tabela citolab_excel (inclui id para exibição na aba)."""
    try:
        conn = create_connection(database)
        if conn is None:
            return pd.DataFrame(columns=["id"] + CITOLAB_EXCEL_COLUNAS)
        query = f"SELECT id, {', '.join(CITOLAB_EXCEL_COLUNAS)} FROM citolab_excel ORDER BY nome_beneficiario"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df.fillna("")
    except Exception as e:
        print(f"Erro ao carregar citolab_excel: {e}")
        return pd.DataFrame(columns=["id"] + CITOLAB_EXCEL_COLUNAS)


def carregar_citolab_excel_db_com_id(database="prestacao_contas.db"):
    """Carrega dados da tabela citolab_excel incluindo a coluna id (para cálculo de proximidade)."""
    try:
        conn = create_connection(database)
        if conn is None:
            return pd.DataFrame(columns=["id"] + CITOLAB_EXCEL_COLUNAS)
        query = f"SELECT id, {', '.join(CITOLAB_EXCEL_COLUNAS)} FROM citolab_excel ORDER BY nome_beneficiario"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df.fillna("")
    except Exception as e:
        print(f"Erro ao carregar citolab_excel (com id): {e}")
        return pd.DataFrame(columns=["id"] + CITOLAB_EXCEL_COLUNAS)


def carregar_labplus_id_paciente(database="prestacao_contas.db"):
    """Retorna lista de (id, nome_beneficiario) da tabela labplus, ordenado por nome_beneficiario."""
    try:
        conn = create_connection(database)
        if conn is None:
            return []
        cur = conn.cursor()
        cur.execute(
            "SELECT id, COALESCE(nome_beneficiario, paciente) FROM labplus ORDER BY COALESCE(nome_beneficiario, paciente)"
        )
        rows = cur.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Erro ao carregar id/paciente labplus: {e}")
        return []


def atualizar_proximidade_labplus(database, id_valores, id_citolab_por_labplus=None):
    """Atualiza as colunas proximidade e, quando fornecido, proximidade_id na tabela labplus.
    id_valores: lista de (id_labplus, valor_proximidade_str).
    id_citolab_por_labplus: lista opcional com mesmo tamanho de id_valores; cada elemento é o id
    da tabela citolab_excel quando proximidade > 0.8, ou None para limpar."""
    try:
        conn = create_connection(database)
        if conn is None:
            return False
        cur = conn.cursor()
        for i, (id_, valor) in enumerate(id_valores):
            if id_citolab_por_labplus is not None and i < len(id_citolab_por_labplus):
                id_citolab = id_citolab_por_labplus[i]
                if id_citolab is not None:
                    cur.execute(
                        "UPDATE labplus SET proximidade = ?, proximidade_id = ? WHERE id = ?",
                        (str(valor), str(id_citolab), id_),
                    )
                else:
                    cur.execute(
                        "UPDATE labplus SET proximidade = ?, proximidade_id = NULL WHERE id = ?",
                        (str(valor), id_),
                    )
            else:
                cur.execute(
                    "UPDATE labplus SET proximidade = ? WHERE id = ?", (str(valor), id_)
                )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao atualizar proximidade: {e}")
        return False


def marcar_checked_labplus(
    database: str, id_labplus: int, id_citolab_excel: int | None = None
) -> bool:
    """Marca o registro labplus como checked = '1' (True).
    Se id_citolab_excel for informado, grava também em proximidade_id."""
    try:
        conn = create_connection(database)
        if conn is None:
            return False
        cur = conn.cursor()
        if id_citolab_excel is not None:
            cur.execute(
                "UPDATE labplus SET checked = ?, proximidade_id = ? WHERE id = ?",
                ("1", str(id_citolab_excel), id_labplus),
            )
        else:
            cur.execute(
                "UPDATE labplus SET checked = ? WHERE id = ?", ("1", id_labplus)
            )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao marcar checked: {e}")
        return False


def desmarcar_checked_labplus(database: str, id_labplus: int) -> bool:
    """Desmarca o registro labplus: checked = '' (False) e proximidade_id = NULL."""
    try:
        conn = create_connection(database)
        if conn is None:
            return False
        cur = conn.cursor()
        cur.execute(
            "UPDATE labplus SET checked = '', proximidade_id = NULL WHERE id = ?",
            (id_labplus,),
        )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao desmarcar checked: {e}")
        return False
