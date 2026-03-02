"""
Constantes compartilhadas da aplicação Prestação de Contas Unimed - Citolab.
"""

# Fonte compatível com macOS (evita "Segoe UI" que gera TclError no Tk)
FONT_FAMILY = "Helvetica"
FONT_TITLE = (FONT_FAMILY, 18, "bold")
FONT_NORMAL = (FONT_FAMILY, 10)
FONT_INPUT = (FONT_FAMILY, 11)
FONT_SMALL = (FONT_FAMILY, 9)

# Cores e estilo (tema profissional)
COLORS = {
    "bg": "#f5f6fa",
    "card_bg": "#ffffff",
    "primary": "#2c3e50",
    "primary_light": "#34495e",
    "accent": "#3498db",
    "accent_hover": "#2980b9",
    "text": "#2c3e50",
    "text_secondary": "#7f8c8d",
    "border": "#dfe6e9",
    "success": "#27ae60",
    "header_bg": "#2c3e50",
}

# Campos da tabela labplus (uma row por procedimento; ordem para exibição e importação)
CAMPOS_TABELA = [
    "id",
    # "numero",
    # "numero_carteira",
    "nome_beneficiario",
    # "dados_solicitante_nome",
    # "prestador_executante_nome",
    "data_hora_atendimento",
    "procedimento_codigo",
    "procedimento_descricao",
    "procedimento_data",
    "procedimento_valor",
    # "procedimento_valor_total",
    "proximidade",
    "proximidade_id",
    "checked",
]

# Valores do combobox de filtro de busca
OPCOES_FILTRO_BUSCA = [
    "todos",
    "numero",
    "numero_carteira",
    "nome_beneficiario",
    "dados_solicitante_nome",
    "prestador_executante_nome",
    "procedimento_data",
    "procedimento_codigo",
    "procedimento_descricao",
]

# Valores do combobox de filtro de busca na aba Citolab Excel (id + colunas)
OPCOES_FILTRO_BUSCA_CITOLAB = [
    "todos",
    "id",
    "data_atendimento",
    "nome_prestador",
    "servico",
    "num_nota",
    "soma",
    "descricao",
    "nome_beneficiario",
]

# --- Tabela Citolab Excel (exibição na aba) ---
# Colunas exibidas na ordem (id + colunas do relatório)
COLUNAS_EXIBICAO_CITOLAB = (
    "id",
    "data_atendimento",
    "nome_beneficiario",
    # "nome_prestador",
    # "servico",
    # "num_nota",
    "descricao",
    "soma",
)
CABECALHOS_CITOLAB_EXCEL = {
    "id": "ID",
    "data_atendimento": "Data Atendimento",
    "nome_prestador": "Prestador",
    "servico": "Serviço",
    "num_nota": "Nº Nota",
    "soma": "Total",
    "descricao": "Procedimento",
    "nome_beneficiario": "Paciente",
}
LARGURAS_CITOLAB_EXCEL = {
    "id": 50,
    "data_atendimento": 70,
    "nome_prestador": 220,
    "servico": 100,
    "num_nota": 100,
    "soma": 90,
    "descricao": 280,
    "nome_beneficiario": 200,
}
COLUNAS_ALINHAMENTO_DIREITA_CITOLAB = ("id", "soma")

# Símbolo exibido na coluna checked da tabela LabPlus (quando marcado)
CHECK_SIMBOLO = "✓"

# Peso da data no cálculo de proximidade (0 a 1). Datas iguais entre LabPlus e Citolab Excel têm este peso; o restante é similaridade de texto.
PESO_DATA_PROXIMIDADE = 0.6

# Mapeamento de nomes de campos para cabeçalhos da tabela em português
CABECALHOS_TABELA = {
    "id": "ID",
    "numero": "Nº Guia",
    "numero_carteira": "Nº Carteira",
    "nome_beneficiario": "Beneficiário",
    "dados_solicitante_nome": "Solicitante",
    "prestador_executante_nome": "Prestador Executante",
    "data_hora_atendimento": "Data/Hora Atend.",  # legado; LabPlus usa procedimento_data
    "procedimento_codigo": "Cód. Procedimento",
    "procedimento_descricao": "Descrição Procedimento",
    "procedimento_data": "Data Procedimento",
    "procedimento_valor": "Valor Procedimento",
    "procedimento_valor_total": "Valor Total",
    "proximidade": "Proximidade",
    "proximidade_id": "Proximidade ID",
    "checked": "Checked",
}

# Larguras das colunas da tabela
LARGURAS_TABELA = {
    "id": 50,
    "numero": 150,
    "numero_carteira": 130,
    "nome_beneficiario": 200,
    "dados_solicitante_nome": 180,
    "prestador_executante_nome": 220,
    "data_hora_atendimento": 140,  # legado; LabPlus usa procedimento_data
    "procedimento_codigo": 120,
    "procedimento_descricao": 280,
    "procedimento_data": 110,
    "procedimento_valor": 110,
    "procedimento_valor_total": 100,
    "proximidade": 100,
    "proximidade_id": 100,
    "checked": 80,
}

# Colunas alinhadas à direita na tabela
COLUNAS_ALINHAMENTO_DIREITA = [
    "id",
    "procedimento_valor",
    "procedimento_valor_total",
]

# Nomes dos meses para o card de competência (exercício = ano, competência = mês)
MESES_NOMES = (
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
)

PERCENTUAL_SIMILARIDADE_MINIMA = 70
