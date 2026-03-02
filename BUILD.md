# Gerar executável (.exe) para Windows

O projeto usa **PyInstaller** para criar um `.exe` que roda no Windows sem precisar instalar Python.

## Pré-requisitos

- **Windows** (o .exe deve ser gerado em uma máquina Windows, ou use uma VM / GitHub Actions).
- **Python 3.10 ou 3.11** instalado no Windows (evite 3.12+ se der problema com PyInstaller).

## Passo a passo no Windows

### 1. Clonar ou copiar o projeto

Copie a pasta do projeto para o Windows (ou clone o repositório).

### 2. Criar ambiente virtual e instalar dependências

Abra o **Prompt de Comando** ou **PowerShell** na pasta do projeto:

```cmd
cd C:\caminho\para\fichas_citolab

python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
pip install pyinstaller
```

### 3. Gerar o executável

**Opção A – Usando o arquivo .spec (recomendado)**

```cmd
pyinstaller fichas_citolab.spec
```

O `.exe` será criado em: `dist\FichasCitolab.exe`

**Opção B – Comando direto (sem .spec)**

```cmd
pyinstaller --onefile --windowed --name FichasCitolab main.py
```

- `--onefile`: um único arquivo .exe  
- `--windowed`: não abre janela preta do CMD  
- `--name`: nome do executável  

Saída: `dist\FichasCitolab.exe`

### 4. Incluir a pasta `data` ao distribuir

O programa procura o CSV em `data\fichas.csv` **na mesma pasta do .exe**.

Ao distribuir, envie:

- `FichasCitolab.exe`
- Pasta `data` com o arquivo `fichas.csv` dentro

Exemplo de estrutura:

```
Pasta do usuário/
├── FichasCitolab.exe
└── data/
    └── fichas.csv
```

Se usar o `.spec` com `datas=[('data', 'data')]`, a pasta `data` é empacotada dentro do .exe. Nesse caso o CSV estará em um caminho interno; o código já trata o caminho quando roda “congelado”. Se quiser que o usuário edite o CSV fora do exe, **não** inclua `data` no `datas` e deixe a pasta `data` junto do .exe na distribuição.

## Resumo dos comandos (Windows)

```cmd
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
pyinstaller fichas_citolab.spec
```

Depois, use a pasta `dist` e a pasta `data` (com `fichas.csv`) para entregar o programa.

## Gerar .exe com GitHub Actions (automático)

O repositório inclui um workflow que gera o `.exe` no Windows, sem precisar de um PC Windows.

### Como usar

1. **Envie o projeto para o GitHub** (crie um repositório e faça push do código).
2. O build roda automaticamente em cada **push** nas branches `main` ou `master`.
3. Para rodar manualmente: no GitHub, vá em **Actions** → **Build Windows EXE** → **Run workflow**.
4. Ao terminar, abra a execução (run) e em **Artifacts** baixe **FichasCitolab-Windows** (contém o `FichasCitolab.exe`).

### Onde estão os workflows

- **Build (artifact)**: `.github/workflows/build-windows.yml`  
  Roda em push em `main`/`master` e manualmente. O `.exe` fica em **Actions** → artifact **FichasCitolab-Windows**.

- **Release (versão para download)**: `.github/workflows/release-windows.yml`  
  Roda quando você publica uma **tag** (ex.: `v1.0.0`). Cria uma Release no GitHub e anexa o `FichasCitolab.exe` para download.

  Para publicar uma versão:
  ```bash
  git tag v1.0.0
  git push origin v1.0.0
  ```
  Depois, em **Releases** no repositório, o .exe estará disponível para download.

### Distribuir para o usuário

Depois de baixar o artifact, envie para o usuário:

- `FichasCitolab.exe`
- Pasta `data` com o arquivo `fichas.csv` (copie do projeto e coloque na mesma pasta do .exe).

### Se o EXE não for gerado no GitHub

Confira o seguinte:

1. **A pasta `.github` está no repositório?**  
   A raiz do repo deve conter `.github/workflows/build-windows.yml`.  
   Verifique: no GitHub, na raiz do repositório, deve aparecer a pasta `.github`.

2. **O workflow está rodando?**  
   Vá em **Actions**. Deve aparecer o workflow **Build Windows EXE**.  
   - Se não aparecer: faça um **push** para a branch `main` ou `master`, ou clique em **Build Windows EXE** → **Run workflow**.  
   - Se o workflow existir mas estiver em amarelo/vermelho: abra a execução e veja em qual step falhou (log).

3. **Estrutura do repositório**  
   Na **raiz** do repositório (onde está a pasta `.github`) devem estar:
   - `main.py`
   - `fichas_citolab.spec`
   - `requirements.txt`  
   Se esses arquivos estiverem dentro de uma subpasta (ex.: `fichas_citolab/main.py`), o workflow vai falhar no step "Verificar arquivos do projeto". Nesse caso, ou mova tudo para a raiz, ou edite o workflow e use `working-directory: nome_da_subpasta` em cada step.

4. **Onde baixar o EXE quando der certo**  
   **Actions** → clique na última execução (verde) → role até **Artifacts** → **FichasCitolab-Windows** (link para download).

---

## Gerar .exe a partir do macOS/Linux (cross-compile)

PyInstaller **não** gera .exe para Windows a partir de macOS ou Linux. Para gerar o .exe você precisa:

1. Usar um **Windows** (físico, VM ou cloud), ou  
2. Usar **GitHub Actions** (recomendado): o workflow acima gera o .exe automaticamente; basta baixar o artifact.
