# Como construir o executável

Este documento descreve como gerar o **PrestacaoContasUnimed.exe** no seu computador e no GitHub Actions.

## Pré-requisitos (build local)

- Python 3.11+
- Dependências instaladas: `pip install -r requirements.txt`
- PyInstaller: `pip install pyinstaller`

## Build local (Windows)

Na raiz do projeto:

```bash
pyinstaller --clean --noconfirm prestacao_contas_unimed.spec
```

- O executável será gerado em **`dist/PrestacaoContasUnimed.exe`**.
- É um único arquivo (onefile), sem console (aplicação GUI).

Para apenas atualizar o build sem limpar cache:

```bash
pyinstaller prestacao_contas_unimed.spec
```

## Build no GitHub Actions

Os workflows do projeto estão em **`.github/workflows/`** e são executados automaticamente pelo GitHub Actions.

### 1. Build em push ou manual — `build-windows.yml`

**Quando roda:**

- Automaticamente em cada **push** nas branches **main** ou **master**.
- Manualmente: **Actions** → **Build Windows EXE** → **Run workflow** → **Run workflow**.

**Como obter o EXE:**

1. Abra **Actions** no GitHub.
2. Clique na execução desejada do workflow **Build Windows EXE**.
3. Role até a seção **Artifacts**.
4. Baixe **PrestacaoContasUnimed-Windows** (arquivo compactado contendo o `.exe`).

### 2. Release com tag — `release-windows.yml`

**Quando roda:**

- Ao publicar uma **tag** no formato **v*** (ex.: `v1.0.0`, `v1.2.3`).

**Como disparar:**

No terminal (com o repositório clonado e atualizado):

```bash
git tag v1.0.0
git push origin v1.0.0
```

Ou pela interface do GitHub: **Releases** → **Create a new release** → informe a tag (ex.: `v1.0.0`) e publique.

**Como obter o EXE:**

1. **Actions** → abra a execução do workflow **Release Windows EXE**; ou
2. **Releases** → clique na release criada (ex.: v1.0.0). O arquivo **PrestacaoContasUnimed.exe** estará listado nos assets da release.

## Resumo

| Objetivo                    | Método                          | Onde está o EXE                                      |
|----------------------------|----------------------------------|------------------------------------------------------|
| Testar / build rápido      | Push em main/master ou Run workflow | Actions → run → Artifacts → PrestacaoContasUnimed-Windows |
| Publicar versão (v1.0.0…)  | Criar tag `v*` e dar push        | Releases → [versão] → PrestacaoContasUnimed.exe      |
| Desenvolvimento local      | `pyinstaller prestacao_contas_unimed.spec` | `dist/PrestacaoContasUnimed.exe`                |
