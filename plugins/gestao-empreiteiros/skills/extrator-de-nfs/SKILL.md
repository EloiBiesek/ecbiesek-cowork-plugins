---
name: extrator-de-nfs
description: >
  Extrai dados de Notas Fiscais de Serviço Eletrônicas (NFSe) dos empreiteiros
  de qualquer obra e consolida na planilha CONTROLE GERAL DE NOTAS FISCAIS DE
  EMPREITEIROS.xlsx. Use esta skill sempre que o trabalho envolver notas fiscais,
  NFSe, extração de dados de NF, controle fiscal de empreiteiros, ISS, INSS retido,
  ISS mensal, conferência de notas, completar dados pendentes, ou qualquer menção
  a "planilha de notas" ou "atualizar notas fiscais".
  Use também quando o usuário pedir para processar, conferir, completar, verificar
  ou atualizar dados de notas fiscais de qualquer obra — mesmo que o pedido pareça simples.
  IMPORTANTE: esta skill funciona para QUALQUER obra — não está restrita a uma
  obra específica. Ela detecta automaticamente pastas de obra com empreiteiros.
---

# Extração de NFSe — Genérica (qualquer obra)

Extrai dados dos PDFs de notas fiscais de cada empreiteiro, consolida na
planilha **CONTROLE GERAL DE NOTAS FISCAIS DE EMPREITEIROS.xlsx** na aba
**"NOTAS FISCAIS (NOVO)"**.

---

## Seleção obrigatória de obra

Antes de executar qualquer script, é **obrigatório** determinar em qual pasta
de obra trabalhar. A skill NUNCA processa mais de uma obra por vez.

### Passo 0 — Descobrir e selecionar a obra

Verifique o diretório de trabalho atual. Se for uma pasta de obra (contém
subpastas numeradas "01 NOME", "02 NOME", ...), use-a diretamente.

Se não for, procure pastas de obra no diretório atual e filhos:

```bash
PYTHONIOENCODING=utf-8 python3 scripts/atualizar_nfse.py --listar-obras
```

**SEMPRE pergunte ao usuário** qual obra processar, mesmo que só uma seja
encontrada. Apresente as opções e aguarde confirmação.

Uma vez selecionada, passe o caminho via `--obra-dir` em todos os comandos:

```bash
python3 scripts/atualizar_nfse.py --obra-dir "/caminho/para/prestadores"
```

Se o diretório de trabalho já É a pasta da obra, `--obra-dir` é opcional.

---

## Carregar skill nativa de PDF antes de qualquer trabalho

**Sempre** invoque a skill nativa de PDF do Claude antes de processar qualquer
arquivo ou dar início ao fluxo de extração:

```
Skill: anthropic-skills:pdf
```

---

## REGRA CRÍTICA — Sempre processar em batches

**NUNCA execute a extração de todos os empreiteiros de uma vez.** Obras grandes
(como Lagunas, com 12+ empreiteiros e centenas de PDFs) sobrecarregam o contexto
e causam falhas.

**Estratégia obrigatória para obras com mais de 3 empreiteiros:**

1. Primeiro, verifique o estado atual (`check_status.py`)
2. Processe **1 a 3 empreiteiros por vez** usando `--empreiteiros`
3. Após cada batch, verifique o resultado antes de continuar
4. Repita até processar todos

**Exemplo para uma obra com 12 empreiteiros:**
```bash
# Batch 1: empreiteiros 01-03
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 1 2 3
# Batch 2: empreiteiros 04-06
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 4 5 6 --incremental
# Batch 3: empreiteiros 07-09
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 7 8 9 --incremental
# Batch 4: empreiteiros 10-12
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 10 11 12 --incremental
```

Use `--incremental` a partir do segundo batch para não reprocessar o que já foi feito.
Use `--batch-size N` para limitar o número de PDFs por execução (default: 50).

Para obras pequenas (3 ou menos empreiteiros), pode executar tudo de uma vez.

---

## Fluxo recomendado — Sequencial

Execute os scripts um a um na ordem abaixo. Todos aceitam `--obra-dir`.

### Passo 1 — Verificar estado atual

```bash
PYTHONIOENCODING=utf-8 python3 scripts/check_status.py --obra-dir "/caminho"
```

### Passo 2 — Extrair dados dos PDFs (em batches)

Para obras grandes, **sempre** use `--empreiteiros` para processar em grupos:

```bash
# Processar empreiteiros específicos
PYTHONIOENCODING=utf-8 python3 scripts/extract_all_nfse.py --obra-dir "/caminho" --empreiteiros 1 2 3

# Com limite de PDFs por execução
python3 scripts/extract_all_nfse.py --obra-dir "/caminho" --empreiteiros 4 5 6 --batch-size 30

# Incremental (pula NFs já processadas)
python3 scripts/extract_all_nfse.py --obra-dir "/caminho" --incremental
```

### Passo 3 — OCR nos PDFs-imagem

```bash
PYTHONIOENCODING=utf-8 python3 scripts/ocr_nfse.py --obra-dir "/caminho" --batch-size 15
```

Usa **RapidOCR + PyMuPDF** (sem Tesseract). Detecta rotação 0°/180° automaticamente.
O pipeline unificado (`atualizar_nfse.py`) executa OCR em loop automático até
processar todos os PDFs-imagem pendentes.

> **Dica:** Se o OCR não resolver, abra o PDF com a skill `anthropic-skills:pdf`
> e extraia os dados manualmente — depois edite o JSON em `.nfse-state/`.

### Passo 4 — Atualizar planilha

```bash
PYTHONIOENCODING=utf-8 python3 scripts/populate_xlsx.py --obra-dir "/caminho"
```

Para adicionar linhas novas sem reescrever tudo:

```bash
python3 scripts/populate_xlsx.py --obra-dir "/caminho" --append
```

### Passo 5 — Validar resultados

```bash
PYTHONIOENCODING=utf-8 python3 scripts/validate_extraction.py --obra-dir "/caminho"
```

---

## Pipeline unificado (atualizar_nfse.py)

Executa todos os 5 passos de uma vez. **Para obras grandes, use `--empreiteiros`:**

```bash
# Obra pequena (3 ou menos empreiteiros) — pode rodar tudo
python3 scripts/atualizar_nfse.py --obra-dir "/caminho/para/prestadores"

# Obra grande — processar por grupos de empreiteiros
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 1 2 3
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --empreiteiros 4 5 6 --incremental

# Outras opções
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --force           # Reprocessar tudo
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --sem-ocr         # Pular OCR
python3 scripts/atualizar_nfse.py --obra-dir "/caminho" --batch-size 30   # Limitar PDFs
```

Se estiver dentro da pasta da obra, `--obra-dir` é opcional — o script detecta
automaticamente. Se houver múltiplas obras, o script pedirá para selecionar.

---

## Estrutura esperada de pastas

```
PRESTADORES DE SERVIÇO/                ← pasta da obra
  01 EMPREITEIRO A/
    NOTA FISCAL/2023/  2024/  2025/    ← PDFs de NFSe por ano
  02 EMPREITEIRO B/
    NOTA FISCAL/
      2024/  2025/
  11 VIGILÂNCIA/
    *.pdf                              ← NFs soltas na raiz (sem subpasta)
  ...
  .nfse-state/                         ← state da skill (criado automaticamente)
    nfse_extracted.json                ← dados extraídos
    obra.json                          ← configuração da obra (opcional)
  CONTROLE GERAL DE NOTAS FISCAIS DE EMPREITEIROS.xlsx  ← planilha de saída
```

---

## Formatos de PDF reconhecidos

**Formato 1 — NFSe padrão Porto Velho** (maioria):
Layout de tabela com cabeçalhos e valores em linhas separadas.

**Formato 2 — "Nota Portovelhense"** (NFs mais antigas):
"VALOR TOTAL DO SERVIÇO R$ xxx" inline, multi-coluna para ISS e INSS.

**Formato 3 — Vigilância** (ex: AQUILAIS):
"Data Fato Gerador" em vez de "Competência", INSS=0, ISS=5%, sem CNO.

---

## Módulo constants_nfse.py

Constantes compartilhadas centralizadas em `scripts/constants_nfse.py`.

Funções-chave:
- `init_obra(path)` — reinicializa todas as constantes para uma obra específica
- `descobrir_obras(dir)` — descobre pastas de obra a partir de um diretório
- `is_obra_dir(path)` — verifica se um diretório contém empreiteiros numerados

---

## Dependências

```bash
pip install pdfplumber openpyxl
# Para OCR (opcional):
pip install rapidocr-onnxruntime pymupdf
```
