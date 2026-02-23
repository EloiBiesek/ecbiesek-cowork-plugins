# gestao-obras

Ferramentas de gestao e auditoria de obras para ECBIESEK.

## Skills

| Skill | Descricao |
|-------|-----------|
| **validador-estrutura** | Valida a estrutura de pastas de obra contra o template padrao (14 secoes numeradas). Identifica pastas faltando, vazias e documentos ausentes. |

## Instalacao

### Claude Desktop / Cowork

1. Adicione o marketplace: `ecbiesek/ecbiesek-cowork-plugins`
2. Instale o plugin: `gestao-obras`

### Claude Code (CLI)

```bash
/plugin marketplace add ecbiesek/ecbiesek-cowork-plugins
/plugin install gestao-obras@ecbiesek-cowork-plugins
```

## Uso rapido

```bash
# Validar uma obra
python scripts/validar_estrutura.py --obra-dir "C:/PASTA MAE/OBRA LAGUNAS"

# Validar todas as obras
python scripts/validar_estrutura.py --pasta-mae "C:/PASTA MAE ECBIESEK"

# Exportar para Excel
python scripts/validar_estrutura.py --pasta-mae "C:/PASTA MAE" --xlsx relatorio.xlsx
```
