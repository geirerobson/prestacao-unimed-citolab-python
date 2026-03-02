# Prestação de Contas Unimed — Citolab

Aplicação desktop para prestação de contas entre dados do **LabPlus (XML)** e do relatório **Citolab (Excel)**. Permite importar XML e Excel, filtrar por exercício/competência, comparar registros por similaridade e marcar itens conferidos.

## Requisitos

- **Python** 3.11 ou superior
- Dependências listadas em `requirements.txt`

## Instalação e execução local

```bash
# Clone o repositório (ou navegue até a pasta do projeto)
cd prestacao_unimed_citolab

# Crie um ambiente virtual (recomendado)
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # Linux/macOS

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
python main.py
```

A janela da aplicação será aberta. O banco SQLite (`prestacao_contas.db`) é criado automaticamente na primeira execução, na pasta do projeto.

## Funcionalidades principais

### Aba LabPlus (XML)

- **Importar XML (LabPlus):** menu **Arquivo → Importar XML (LabPlus)...** para carregar o XML de fichas.
- **Filtros:** Exercício (ano) — campo numérico com ano atual como padrão — e Competência (mês).
- **Busca:** filtro por termo em vários campos (nome, carteira, procedimento etc.).
- **Tabela:** exibe registros com data, procedimento, valor, proximidade e status (checked).
- **Duplo clique:** abre um modal com registros do Citolab Excel da mesma data; ao escolher um registro, marca o item LabPlus como conferido e grava o `proximidade_id`.
- **Clique direito:** menu **Marcar (checked)** ou **Desmarcar (checked e proximidade_id)**.

### Aba Citolab Excel

- **Importar Excel (CITOLAB):** menu **Arquivo → Importar Excel (CITOLAB)...** para carregar o relatório Excel.
- **Filtros:** Exercício, Competência e busca por termo.
- **Tabela:** exibe dados do Excel importado (data, paciente, procedimento, valor etc.).

### Verificar proximidade

- Botão **Verificar** (abaixo das abas): para o exercício e competência definidos na aba LabPlus, calcula a similaridade entre cada registro LabPlus e os registros do Citolab Excel da mesma data (nome + descrição). Atualiza os campos **proximidade** e **proximidade_id** no banco.
- Durante o processamento, é exibido um indicador de progresso; ao concluir, um alerta é mostrado e o indicador some após o usuário fechar o alerta.

## Estrutura do projeto

```
prestacao_unimed_citolab/
├── main.py                 # Aplicação principal (tkinter)
├── config.py               # Constantes, cores, colunas, fontes
├── init.py                 # Banco SQLite, carregar/gravar dados
├── card_busca.py           # Componente de busca
├── card_competencia.py     # Filtros Exercício (ano) e Competência (mês)
├── tabela_labplus.py       # Treeview dos dados LabPlus
├── tabela_citolab_excel.py # Treeview dos dados Citolab Excel
├── importar_labplus.py     # Importação do XML LabPlus
├── importar_excel_citolab.py # Importação do Excel Citolab
├── requirements.txt
├── prestacao_contas_unimed.spec  # PyInstaller (gerar .exe)
├── .github/workflows/      # GitHub Actions
│   ├── build-windows.yml   # Build EXE (push main/master ou manual)
│   └── release-windows.yml # Release com EXE ao publicar tag v*
└── README.md
```

## Gerar o executável

### No computador (Windows)

É necessário ter Python e as dependências instalados. Na pasta do projeto:

```bash
pip install pyinstaller
pyinstaller --clean --noconfirm prestacao_contas_unimed.spec
```

O arquivo **PrestacaoContasUnimed.exe** será gerado em `dist/`.

### No GitHub

- **Build em push ou manual:** o workflow **Build Windows EXE** roda em push para `main`/`master` ou ao clicar em **Run workflow** em Actions. O EXE fica em **Actions → [run] → Artifacts → PrestacaoContasUnimed-Windows**.
- **Release com tag:** crie uma tag no formato `v*` (ex.: `v1.0.0`) e envie para o GitHub. O workflow **Release Windows EXE** gera o EXE e anexa na Release. O arquivo **PrestacaoContasUnimed.exe** aparece em **Releases → [versão] → Assets**.

Detalhes em [docs/BUILD.md](docs/BUILD.md).

## Licença

Uso interno / conforme definido pelo projeto.
