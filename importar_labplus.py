"""
Importação de arquivos XML TISS LabPlus (guia SP-SADT) para a tabela labplus.
Usa nome expandido do namespace (Clark) para funcionar com qualquer prefixo no XML.
"""

import xml.etree.ElementTree as ET
from pathlib import Path

from config import CAMPOS_TABELA
from init import create_connection

# URI do padrão TISS (ANS) - no XML pode aparecer como ans: ou ansTISS:
URI_TISS = "http://www.ans.gov.br/padroes/tiss/schemas"


def _tag(local_name: str) -> str:
    """Retorna o nome expandido do elemento (Clark notation) para o namespace TISS."""
    return f"{{{URI_TISS}}}{local_name}"


def _local(tag: str) -> str:
    """Retorna o nome local da tag (sem namespace)."""
    if tag and "}" in tag:
        return tag.split("}", 1)[1]
    return tag or ""


def _text(el) -> str:
    """Retorna o texto do elemento ou string vazia."""
    if el is None:
        return ""
    return (el.text or "").strip()


def _find_el(el, local_name: str):
    """Encontra o primeiro subelemento com o nome local dado (qualquer namespace)."""
    if el is None:
        return None
    target = local_name
    for child in el:
        if _local(child.tag) == target:
            return child
    # Fallback: tentar com namespace TISS
    full = _tag(local_name)
    return el.find(full)


def _find_direct(el, local_name: str):
    """Encontra subelemento direto por nome local ou pelo tag completo TISS."""
    if el is None:
        return None
    r = _find_el(el, local_name)
    if r is not None:
        return r
    return el.find(_tag(local_name))


def _extrair_dados_guia(guia_el) -> dict:
    """Extrai os dados comuns da guia (beneficiário, solicitante, executante, totais)."""
    dados = {}

    ident = _find_direct(guia_el, "identificacaoGuiaSADTSP")
    if ident is not None:
        dados["numero"] = _text(_find_direct(ident, "numeroGuiaPrestador")) or ""

    ben = _find_direct(guia_el, "dadosBeneficiario")
    if ben is not None:
        dados["numero_carteira"] = _text(_find_direct(ben, "numeroCarteira")) or ""
        dados["nome_beneficiario"] = _text(_find_direct(ben, "nomeBeneficiario")) or ""

    sol = _find_direct(guia_el, "dadosSolicitante")
    if sol is not None:
        contr = _find_direct(sol, "contratado")
        if contr is not None:
            dados["dados_solicitante_nome"] = _text(_find_direct(contr, "nomeContratado")) or ""

    prest_exec = _find_direct(guia_el, "prestadorExecutante")
    if prest_exec is not None:
        dados["prestador_executante_nome"] = _text(_find_direct(prest_exec, "nomeContratado")) or ""

    dados["data_hora_atendimento"] = _text(_find_direct(guia_el, "dataHoraAtendimento")) or ""

    vtot_el = _find_direct(guia_el, "valorTotal")
    if vtot_el is not None:
        total_el = _find_direct(vtot_el, "totalGeral")
        if total_el is not None:
            dados["procedimento_valor_total"] = _text(total_el) or ""
        else:
            dados["procedimento_valor_total"] = ""
    else:
        dados["procedimento_valor_total"] = ""

    return dados


def extrair_guia(guia_el) -> list[dict]:
    """
    Extrai da guiaSP_SADT uma lista de linhas: uma por procedimento.
    Cada linha contém: numeroCarteira, nomeBeneficiario, dadosSolicitanteNome,
    prestadorExecutanteNome, dataHoraAtendimento, procedimentoCodigo, procedimentoDescricao,
    procedimentoData, procedimentoValor, procedimentoValorTotal e demais campos da tabela.
    """
    base = _extrair_dados_guia(guia_el)
    linhas = []

    proc_realizados = _find_direct(guia_el, "procedimentosRealizados")
    if proc_realizados is None:
        # Guia sem procedimentos: insere uma linha só com dados da guia
        row = {c: "" for c in CAMPOS_TABELA}
        for k, v in base.items():
            if k in CAMPOS_TABELA:
                row[k] = v
        row["procedimento_codigo"] = ""
        row["procedimento_descricao"] = ""
        row["procedimento_data"] = ""
        row["procedimento_valor"] = ""
        linhas.append(row)
        return linhas

    # No XML TISS, cada bloco <procedimentos> contém: <procedimento>, <data>, <horaInicio>,
    # <horaFim>, <quantidadeRealizada>, <valor>, <valorTotal> (irmãos, não dentro de procedimento).
    for bloco in proc_realizados:
        if _local(bloco.tag) != "procedimentos":
            continue
        # Mapear filhos do bloco por nome local para pegar valor e valorTotal corretos
        by_name = {}
        for filho in bloco:
            by_name[_local(filho.tag)] = filho
        proc_el = by_name.get("procedimento")
        if proc_el is None:
            continue
        row = {c: "" for c in CAMPOS_TABELA}
        for k, v in base.items():
            if k in CAMPOS_TABELA:
                row[k] = v
        row["procedimento_codigo"] = _text(_find_direct(proc_el, "codigo")) or ""
        row["procedimento_descricao"] = _text(_find_direct(proc_el, "descricao")) or ""
        # data, valor e valorTotal vêm dos irmãos no bloco <procedimentos>, não dentro de <procedimento>
        row["procedimento_data"] = _text(by_name.get("data")) or ""
        row["procedimento_valor"] = _text(by_name.get("valor")) or ""
        row["procedimento_valor_total"] = _text(by_name.get("valorTotal")) or ""
        if not row["procedimento_data"] and base.get("data_hora_atendimento"):
            row["procedimento_data"] = base["data_hora_atendimento"]
        # Se valorTotal do bloco estiver vazio, usar o total geral da guia
        if not row["procedimento_valor_total"] and base.get("procedimento_valor_total"):
            row["procedimento_valor_total"] = base["procedimento_valor_total"]
        linhas.append(row)

    return linhas


def ler_xml_labplus(caminho: str | Path) -> list[dict]:
    """
    Lê um arquivo XML TISS LabPlus e retorna lista de dicionários (uma entrada por procedimento).
    Cada item tem: numero_carteira, nome_beneficiario, dados_solicitante_nome,
    prestador_executante_nome, data_hora_atendimento, procedimento_codigo, procedimento_descricao,
    procedimento_data, procedimento_valor, procedimento_valor_total, etc.
    """
    tree = ET.parse(caminho)
    root = tree.getroot()
    guias = []
    tag_guia = "guiaSP_SADT"
    for el in root.iter():
        if _local(el.tag) == tag_guia:
            guias.append(el)
    registros = []
    for g in guias:
        registros.extend(extrair_guia(g))
    return registros


def _colunas_insercao_labplus(conn) -> list[str]:
    """Retorna as colunas da tabela labplus (exceto id) para INSERT."""
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(labplus)")
    return [row[1] for row in cur.fetchall() if row[1] != "id"]


def inserir_no_banco(
    registros: list[dict], database: str = "prestacao_contas.db"
) -> tuple[int, int]:
    """
    Insere os registros na tabela labplus (uma row por procedimento).
    Ignora duplicados quando o registro inteiro já existe (todos os campos iguais).
    Preenche colunas obrigatórias do schema antigo (ex.: paciente) a partir de nome_beneficiario.
    Retorna (quantidade_inserida, quantidade_duplicados_ignorados).
    """
    if not registros:
        return 0, 0
    conn = create_connection(database)
    if conn is None:
        return 0, 0
    try:
        cur = conn.cursor()
        cols = _colunas_insercao_labplus(conn)
        placeholders = ", ".join("?" for _ in cols)
        sql_insert = f"INSERT INTO labplus ({', '.join(cols)}) VALUES ({placeholders})"

        # Carregar todos os registros existentes (registro inteiro) para verificação de duplicata
        cur.execute(f"SELECT {', '.join(cols)} FROM labplus")
        registros_existentes = {
            tuple(str(v or "").strip() for v in row)
            for row in cur.fetchall()
        }
        inseridos = 0
        duplicados = 0
        for r in registros:
            # Monta a tupla do registro completo (mesma ordem de cols)
            nome_benef = str(r.get("nome_beneficiario", "") or "")
            valores = []
            for c in cols:
                if c == "paciente":
                    valores.append(nome_benef)
                else:
                    valores.append(str(r.get(c, "") or "").strip())
            chave_registro = tuple(valores)
            if chave_registro in registros_existentes:
                duplicados += 1
                continue
            cur.execute(sql_insert, valores)
            registros_existentes.add(chave_registro)
            inseridos += 1
        conn.commit()
        return inseridos, duplicados
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def importar_xml_labplus(
    caminho: str | Path, database: str = "prestacao_contas.db"
) -> tuple[int, int, str | None]:
    """
    Importa um arquivo XML LabPlus para o banco. Registros duplicados (registro inteiro igual) são ignorados.
    Retorna (quantidade_inserida, quantidade_duplicados_ignorados, mensagem_erro).
    Se houver erro, retorna (0, 0, mensagem).
    """
    try:
        registros = ler_xml_labplus(caminho)
        if not registros:
            return 0, 0, "Nenhum procedimento encontrado no arquivo (guias SP-SADT)."
        inseridos, duplicados = inserir_no_banco(registros, database)
        return inseridos, duplicados, None
    except ET.ParseError as e:
        return 0, 0, f"Arquivo XML inválido: {e}"
    except Exception as e:
        return 0, 0, str(e)
