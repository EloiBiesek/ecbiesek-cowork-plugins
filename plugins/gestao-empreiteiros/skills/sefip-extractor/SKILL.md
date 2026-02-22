---
name: sefip-extractor
description: >
  Extrai dados de relatórios SEFIP dos empreiteiros de qualquer obra e atualiza
  a aba "Alocação de colaboradores" da planilha de controle com a quantidade de
  trabalhadores CAT 01 por mês. Também registra log de execução na aba "RESUMO NOVO".
  Use esta skill sempre que o trabalho envolver SEFIPs, GFIP, FGTS, alocação de
  colaboradores, quantidade de trabalhadores, CAT 01, ou qualquer menção à aba
  de alocação ou ao resumo de execuções de qualquer obra.
  Use também quando o usuário pedir para "processar SEFIP", "atualizar alocação",
  "verificar cobertura de competências" ou variações desses termos.
  IMPORTANTE: esta skill funciona para QUALQUER obra — não está restrita a uma
  obra específica. Ela detecta automaticamente pastas de obra com empreiteiros.
---

# Extração de SEFIP — Genérica (qualquer obra)

Extrai a quantidade de trabalhadores **CAT 01** dos relatórios SEFIP de cada
empreiteiro, preenche a aba **"Alocação de colaboradores"** e registra um log
de execução na aba **"RESUMO NOVO"** da planilha-mestre.

---

## Seleção obrigatória de obra

Antes de executar qualquer script, é **obrigatório** determinar em qual pasta
de obra trabalhar. A skill NUNCA processa mais de uma obra por vez.

### Passo 0 — Descobrir e selecionar a obra

Verifique o diretório de trabalho atual. Se for uma pasta de obra (contém
subpastas numeradas "01 NOME", "02 NOME", ...), use-a diretamente.

Se não for, procure pastas de obra no diretório atual e filhos:

```bash
PYTHONIOENCODING=utf-8 python3 scripts/atualizar_sefip.py --listar-obras
```

**SEMPRE pergunte ao usuário** qual obra processar, mesmo que só uma seja
encontrada. Apresente as opções e aguarde confirmação.

Uma vez selecionada, passe o caminho via `--obra-dir` em todos os comandos:

```bash
python3 scripts/atualizar_sefip.py --obra-dir "/caminho/para/prestadores"
```

Se o diretório de trabalho já É a pasta da obra, `--obra-dir` é opcional.

### Confirmação obrigatória do CNO

Após selecionar a obra, **sempre confirme o CNO com o usuário antes de processar**.
O CNO (Cadastro Nacional de Obra) é essencial para filtrar corretamente os
trabalhadores alocados na obra específica. Se o CNO estiver errado, a extração
pode retornar o total de trabalhadores do empregador em vez dos alocados na obra.

1. Verifique se já existe um `obra.json` em `<obra>/.sefip-state/obra.json`
2. Se existir, leia o CNO configurado e **mostre ao usuário para confirmação**
3. Se não existir, **pergunte ao usuário qual é o CNO da obra**
4. Só prossiga com a extração após o usuário confirmar o CNO

Exemplo de interação:
> "O CNO configurado para esta obra é **90.015.22526/72**. Está correto?"

Se o CNO estiver errado, o usuário deve informar o correto e a configuração
será atualizada antes de prosseguir.

---

## Carregar skill nativa de PDF antes de qualquer trabalho

**Sempre** invoque a skill nativa de PDF do Claude antes de processar qualquer
arquivo ou dar início ao fluxo de extração:

```
Skill: anthropic-skills:pdf
```

---

## Fluxo recomendado — Sequencial

Execute os scripts um a um na ordem abaixo. Este é o fluxo padrão e mais
confiável. Todos os comandos aceitam `--obra-dir` para especificar a obra.

### Passo 1 — Verificar estado atual

```bash
PYTHONIOENCODING=utf-8 python3 scripts/check_status.py
```

Leia o relatório **e o VEREDICTO FINAL no rodapé**:

- **Se o veredicto for `TUDO ATUALIZADO`**: **PARE AQUI.** Não execute nenhum
  outro script. Informe ao usuário que tudo já está em dia.
- **Se o veredicto for `AÇÃO NECESSÁRIA`**: siga apenas os passos indicados
  no veredicto. Não execute scripts desnecessários.

> **IMPORTANTE:** Células vazias na planilha NÃO significam necessariamente que
> há PDFs para processar. Muitos meses não têm SEFIP (alguns empreiteiros nunca
> tiveram, e alguns meses têm apenas documentos não-SEFIP como BOLETO FGTS,
> FOLHA DE PAGAMENTO, CRÉDITO INSS ou DCTFWeb). O veredicto já leva isso em conta.

### Passo 2 — Extrair texto dos PDFs (somente se indicado pelo veredicto)

```bash
PYTHONIOENCODING=utf-8 python3 scripts/extract_text.py
```

Para processar apenas empreiteiros específicos:

```bash
python3 scripts/extract_text.py --empreiteiros 3 7 10
```

**Comportamento incremental:** meses já extraídos com sucesso são pulados
automaticamente. Para re-extrair tudo do zero:

```bash
python3 scripts/extract_text.py --force-reprocess
```

### Passo 3 — OCR nos PDFs escaneados

```bash
PYTHONIOENCODING=utf-8 python3 scripts/extract_ocr.py
```

Processa os arquivos que `extract_text.py` não conseguiu ler (PDFs escaneados,
rotacionados, ou sem camada de texto).

```bash
python3 scripts/extract_ocr.py --batch-size 10
python3 scripts/extract_ocr.py --empreiteiros 3 7
```

> **Dica:** Se o script Python de OCR falhar em algum arquivo, abra o PDF
> diretamente com a skill `anthropic-skills:pdf` e leia o valor CAT 01
> manualmente — depois adicione-o no state file de OCR.

### Passo 4 — Atualizar planilha

```bash
PYTHONIOENCODING=utf-8 python3 scripts/update_planilha.py
```

### Passo 5 — Atualizar aba RESUMO NOVO

```bash
PYTHONIOENCODING=utf-8 python3 scripts/write_resumo.py
```

### Passo 6 — Resolver divergências (se houver)

```bash
python3 scripts/resolve_divergences.py --list
python3 scripts/resolve_divergences.py --keep-all-planilha --apply
python3 scripts/resolve_divergences.py --accept-pdf 0 1 --apply
```

---

## Pipeline unificado (atualizar_sefip.py)

Para executar todos os passos de uma vez:

```bash
python3 scripts/atualizar_sefip.py --obra-dir "/caminho/para/prestadores"
python3 scripts/atualizar_sefip.py --obra-dir "/caminho" --force
python3 scripts/atualizar_sefip.py --obra-dir "/caminho" --empreiteiros 3 7
python3 scripts/atualizar_sefip.py --obra-dir "/caminho" --sem-ocr
```

Se estiver dentro da pasta da obra, `--obra-dir` é opcional — o script detecta
automaticamente. Se houver múltiplas obras disponíveis, o script pedirá para
selecionar uma interativamente.

---

## Configuração de uma obra nova

Para configurar uma obra nova (primeira vez):

```bash
python3 scripts/configurar_obra.py --obra-dir "/caminho/para/prestadores"
```

O wizard auto-detecta empreiteiros, subpastas e range de meses. Pede apenas
o CNO e o nome da obra. A configuração é salva em `<obra>/.sefip-state/obra.json`.

---

## Estrutura esperada de pastas de uma obra

```
PRESTADORES DE SERVIÇO/                ← pasta da obra
  01 EMPREITEIRO A/
    SEFIP/2023/  2024/  2025/          ← ou DOCUMENTOS MENSAIS/
    NOTA FISCAL/                        ← (opcional)
  02 EMPREITEIRO B/
    DOCUMENTOS MENSAIS/
      2024/07 2024/  08 2024/  ...
      2025/01 2025/  02 2025/  ...
  ...
  .sefip-state/                         ← state da skill (criado automaticamente)
    obra.json                           ← configuração da obra
    extractions_text.json
    extractions_ocr.json
    ...
  Controle de Alocação e ISS (...).xlsx ← planilha-mestre
```

Os scripts percorrem automaticamente `SEFIP/`, `DOCUMENTOS MENSAIS/`,
`DOCUMENTAÇÕES MENSAIS/` e `ENTREGA DE DOCUMENTOS MENSAIS/` por empreiteiro.

---

## Formatos de SEFIP reconhecidos

**Formato 1 — Extrato FGTS** (mais comum):
`Qtd. Trabalhadores: XX` ou `Qtd. Trabalhadores FGTS: XX`

**Formato 2 — SEFIP Clássico**:
Bloco `RESUMO DO FECHAMENTO` com inscrição CNO da obra, tabela com CAT 01.

**Formato 3 — Detalhe da Guia FGTS** (Gestão de Guias):
Valor da coluna "Qtd. Trabalhadores" do tomador CNO da obra,
**não** o "Qtd. Trabalhadores FGTS" do cabeçalho (que inclui todos os tomadores).

**Formato 4 — FGTS Digital (GFD)**:
Documento "GFD - Guia do FGTS Digital" com tabela contendo competência,
quantidade de trabalhadores e valores.

**PDFs com texto invertido (rotação 180°)**:
Detectados e corrigidos automaticamente pelo extrator.

**Documentos que NÃO são SEFIP** (filtrados automaticamente):
BOLETO FGTS, FOLHA DE PAGAMENTO, CRÉDITO INSS, COMPENSAÇÃO INSS, DCTFWeb,
GUIA DO FGTS (pagamento), HOLERITE, COMPROVANTE DE DECLARAÇÃO, etc.

---

## Arquivos de estado

Todos em `<obra>/.sefip-state/`:

| Arquivo | Conteúdo |
|---|---|
| `obra.json` | Configuração da obra (CNO, empreiteiros, meses, etc.) |
| `extractions_text.json` | Resultados de extração por texto |
| `extractions_ocr.json` | Resultados de extração por OCR |
| `needs_ocr.json` | PDFs identificados como escaneados |
| `divergences.json` | Conflitos planilha vs PDF |

---

## Módulo constants.py

Todas as constantes compartilhadas estão centralizadas em `scripts/constants.py`.
Para adicionar novos meses ou empreiteiros, usar `configurar_obra.py` ou editar
o `obra.json` da obra.

Funções-chave:
- `init_obra(path)` — reinicializa todas as constantes para uma obra específica
- `descobrir_obras(dir)` — descobre pastas de obra a partir de um diretório
- `is_obra_dir(path)` — verifica se um diretório contém empreiteiros numerados

---

## Dependências

```bash
pip install pdfplumber openpyxl pymupdf
# Para OCR (opcional):
pip install rapidocr-onnxruntime
```
