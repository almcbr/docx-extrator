# DOCX Extractor

Script em Python para extrair capitulos e secoes de um arquivo `.docx`, como:

- capitulos
- prefacio
- introducao
- anexos
- glossario
- bibliografia

O programa analisa o documento e pode:

- mostrar um resumo das secoes detectadas no terminal
- gerar um arquivo JSON com o resultado
- salvar cada secao em um arquivo `.txt`

## Pre-requisitos

- Python 3.10 ou superior
- `pip` disponivel no ambiente

## Instalacao das dependencias

### 1. Clone o repositorio

```bash
git clone https://github.com/almcbr/docx-extrator.git
cd docx-extrator
```

### 2. Crie e ative um ambiente virtual

No Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

No Linux ou macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instale as bibliotecas necessarias

```bash
pip install -r requirements.txt
```

Ou, se preferir instalar manualmente:

```bash
pip install python-docx lxml
```

## Como rodar

O script principal do projeto e:

```bash
python docx-chapter-extractor.py caminho/para/arquivo.docx
```

Exemplo:

```bash
python docx-chapter-extractor.py "meu-arquivo.docx"
```

## Modos de uso

### 1. Mostrar o resumo no terminal

```bash
python docx-chapter-extractor.py livro.docx
```

Esse comando imprime algo como:

- quantidade de secoes detectadas
- tipo da secao
- score de deteccao
- confianca
- titulo encontrado

### 2. Gerar um arquivo JSON

```bash
python docx-chapter-extractor.py livro.docx --json saida.json
```

O arquivo JSON gerado contem, para cada secao:

- `kind`
- `title`
- `start_para`
- `end_para`
- `confidence`
- `detection_score`
- `text`

### 3. Salvar cada secao em arquivos `.txt`

```bash
python docx-chapter-extractor.py livro.docx --outdir capitulos
```

Nesse caso, o programa cria a pasta informada e salva arquivos com nomes no formato:

```text
001_capitulo_titulo_da_secao.txt
002_prefacio_apresentacao.txt
```

### 4. Rodar sem imprimir resumo no terminal

```bash
python docx-chapter-extractor.py livro.docx --json saida.json --quiet
```

## Estrutura esperada da saida

- `--json`: gera um unico arquivo JSON com todas as secoes extraidas
- `--outdir`: gera um arquivo `.txt` por secao
- sem opcoes: mostra apenas o resumo no terminal

## Dependencias usadas

- `python-docx`
- `lxml`

## Observacoes

- O arquivo de entrada precisa estar no formato `.docx`
- O script usa heuristicas para detectar titulos e secoes especiais
- O resultado pode variar de acordo com a formatacao do documento
