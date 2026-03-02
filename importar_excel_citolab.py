"""
Importação do arquivo Excel CITOLAB (ex.: CITOLAB - OUT-2025.XLS) para a tabela citolab_excel.
O relatório tem cabeçalho na linha 20 (índice 19).
Cada linha do Excel corresponde a uma linha no banco; não há agregação de dados.
"""

from __future__ import annotations

import pandas as pd
from pathlib import Path
from typing import Union

from init import create_connection

# Campos a importar: nome no Excel -> nome no banco (citolab_excel)
# DATA_ATENDIMENTO, NOME_PRESTADOR, SERVICO, NUM_NOTA, SOMA (TOTAL), DESCRICAO (NOME_PROCEDIMENTO), NOME_BENEFICIARIO (PACIENTE)
COLUNAS_EXCEL = [
    "DATA_ATENDIMENTO",
    "NOME_PRESTADOR",
    "SERVICO",
    "NUM_NOTA",
    "SOMA",
    "DESCRICAO",
    "NOME_BENEFICIARIO",
]

COLUNAS_BANCO = [
    "data_atendimento",
    "nome_prestador",
    "servico",
    "num_nota",
    "soma",
    "descricao",
    "nome_beneficiario",
]

def _valor_str(val) -> str:
    """Converte valor para string (trata NaN, NaT, etc.)."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    if hasattr(val, "strftime"):
        try:
            return val.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, OSError, TypeError):
            return ""
    return str(val).strip()


def ler_excel_citolab(caminho: Union[str, Path]) -> pd.DataFrame:
    """
    Lê o arquivo Excel CITOLAB. Cabeçalho na linha 20 (índice 19).
    Retorna DataFrame apenas com as colunas mapeadas para o banco (uma linha por linha do Excel).
    """
    df = pd.read_excel(caminho, sheet_name=0, header=19, engine="xlrd")
    # Selecionar colunas que existem no Excel
    disponiveis = [c for c in COLUNAS_EXCEL if c in df.columns]
    df = df[disponiveis].copy()
    # Renomear para nomes do banco
    df.columns = [COLUNAS_BANCO[COLUNAS_EXCEL.index(c)] for c in disponiveis]
    # Garantir que temos todas as colunas do banco (na ordem)
    for col in COLUNAS_BANCO:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_BANCO]
    return df


def inserir_citolab_excel(
    df: pd.DataFrame, database: str = "prestacao_contas.db"
) -> tuple:
    """Insere os registros na tabela citolab_excel, ignorando duplicatas.
    Duplicata = registro inteiro igual (todos os campos iguais).
    Cada linha do DataFrame vira uma linha no banco (sem agregação).
    Retorna (quantidade_inserida, quantidade_duplicados_ignorados)."""
    if df.empty:
        return 0, 0
    cols = COLUNAS_BANCO
    placeholders = ", ".join("?" for _ in cols)
    sql = f"INSERT INTO citolab_excel ({', '.join(cols)}) VALUES ({placeholders})"
    conn = create_connection(database)
    if conn is None:
        return 0, 0
    try:
        cur = conn.cursor()
        # Carregar todos os registros existentes (registro inteiro) para verificação de duplicata
        cur.execute(f"SELECT {', '.join(cols)} FROM citolab_excel")
        registros_existentes = {
            tuple(str(v or "").strip() for v in row)
            for row in cur.fetchall()
        }
        inseridos = 0
        duplicados = 0
        for _, row in df.iterrows():
            valores = [_valor_str(row.get(c, "")).strip() for c in cols]
            chave_registro = tuple(valores)
            if not any(chave_registro):
                continue  # ignora linha com todos os campos vazios
            if chave_registro in registros_existentes:
                duplicados += 1
                continue
            cur.execute(sql, valores)
            registros_existentes.add(chave_registro)
            inseridos += 1
        conn.commit()
        return inseridos, duplicados
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def importar_excel_citolab(
    caminho: str | Path,
    database: str = "prestacao_contas.db",
    substituir: bool = False,
) -> tuple:
    """
    Importa o Excel CITOLAB para a tabela citolab_excel. Uma linha do Excel = uma linha no banco.
    Registros duplicados (registro inteiro igual) são ignorados.
    Se substituir=True, remove todos os registros antes de inserir.
    Retorna (quantidade_inserida, quantidade_duplicados_ignorados, mensagem_erro).
    """
    try:
        df = ler_excel_citolab(caminho)
        if df.empty:
            return (
                0,
                0,
                "Nenhum dado encontrado no arquivo (verifique se o cabeçalho está na linha 20).",
            )
        conn = create_connection(database)
        if conn is None:
            return 0, 0, "Erro ao conectar no banco de dados."
        if substituir:
            cur = conn.cursor()
            cur.execute("DELETE FROM citolab_excel")
            conn.commit()
        conn.close()
        inseridos, duplicados = inserir_citolab_excel(df, database)
        return inseridos, duplicados, None
    except FileNotFoundError:
        return 0, 0, f"Arquivo não encontrado: {caminho}"
    except Exception as e:
        return 0, 0, str(e)
