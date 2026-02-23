---
name: validador-estrutura
description: >
  Valida a estrutura de pastas de obras da ECBIESEK contra o template padrão
  (14 seções numeradas: 01 - DOCUMENTOS DA OBRA até 14 - OUTROS).
  Identifica pastas faltando, vazias, com nomeação incorreta e documentos
  obrigatórios ausentes. Gera relatório de conformidade no console ou Excel.
  Use esta skill quando o trabalho envolver validação de estrutura de obra,
  auditoria de pastas, verificação de conformidade, pasta padrão, ou qualquer
  menção a "verificar pastas", "auditar obra", "estrutura de obra",
  "pasta padrão", "conformidade de pasta", "relatório de conformidade".
  IMPORTANTE: esta skill funciona para QUALQUER obra — detecta automaticamente
  pastas de obra pela presença de subpastas numeradas no padrão "NN - DESCRIÇÃO".
---

# Validador de Estrutura de Obra — ECBIESEK

Verifica se as pastas de obra seguem a estrutura padrão ECBIESEK com 14 seções
numeradas. Identifica não-conformidades e gera relatórios.

---

## Seleção de obra ou Pasta Mãe

A skill pode operar em dois modos:

### Modo 1 — Obra individual
Valida uma pasta de obra específica:
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --obra-dir "/caminho/para/OBRA LAGUNAS"
```

### Modo 2 — Pasta Mãe (todas as obras)
Valida todas as obras encontradas dentro de um diretório:
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --pasta-mae "/caminho/para/PASTA MÃE ECBIESEK"
```

### Descobrir obras disponíveis
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --listar-obras
```

**SEMPRE pergunte ao usuário** se quer validar uma obra específica ou todas.

---

## Fluxo recomendado

### Passo 1 — Listar obras disponíveis

```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --listar-obras
```

Apresente as obras encontradas e pergunte ao usuário qual validar
(ou se deseja validar todas).

### Passo 2 — Executar validação

**Obra individual:**
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --obra-dir "/caminho/para/obra"
```

**Todas as obras:**
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --pasta-mae "/caminho/para/pasta-mae"
```

### Passo 3 — Exportar relatório (opcional)

Se o usuário quiser o relatório em Excel:
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --pasta-mae "/caminho" --xlsx "relatorio_conformidade.xlsx"
```

Se o usuário quiser o relatório em JSON:
```bash
PYTHONIOENCODING=utf-8 python3 scripts/validar_estrutura.py --pasta-mae "/caminho" --json "resultados.json"
```

---

## Template customizado

O template padrão está em `scripts/template_obra.json` com as 14 seções da
estrutura ECBIESEK. Para usar um template diferente:

```bash
python3 scripts/validar_estrutura.py --obra-dir "/caminho" --template "meu_template.json"
```

O template define:
- `pattern`: nome esperado da pasta (suporta `*` como wildcard)
- `required`: se a seção é obrigatória
- `expected_files`: padrões de arquivos que devem existir na seção
- `children`: subpastas esperadas dentro da seção

---

## Estrutura padrão ECBIESEK (14 seções)

| # | Seção | Obrigatória |
|---|-------|-------------|
| 01 | DOCUMENTOS DA OBRA | Sim |
| 02 | DOCUMENTOS DA EMPRESA (SPE) | Sim |
| 03 | DOCUMENTOS DO TERRENO | Sim |
| 04 | DOCUMENTOS DA SCP | Não |
| 05 | INCORPORAÇÃO | Sim |
| 06 | PROJETOS | Sim |
| 07 | PRESTADORES DE SERVIÇO | Sim |
| 08 | IMPOSTOS E ENCARGOS | Sim |
| 09 | CLIENTES | Sim |
| 10 | COMPRAS | Sim |
| 11 | CONTABILIDADE | Sim |
| 12 | ENGENHARIA | Sim |
| 13 | GESTÃO DE CANTEIRO | Sim |
| 14 | OUTROS | Não |

---

## O que o relatório verifica

1. **Pastas obrigatórias presentes** — Todas as 12 seções obrigatórias existem?
2. **Pastas vazias** — Seções que existem mas não contêm nenhum arquivo
3. **Documentos obrigatórios** — Arquivos esperados por seção (CNO, CNPJ, Memorial, etc.)
4. **Nomeação correta** — Pastas seguem o formato "NN - DESCRIÇÃO"?
5. **Pastas extras** — Pastas que não correspondem ao template (informativo)

---

## Dependências

O script funciona com Python puro (stdlib). Para exportar Excel:

```bash
pip install openpyxl
```
