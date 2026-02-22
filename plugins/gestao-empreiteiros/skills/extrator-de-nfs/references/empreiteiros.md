# Referência de Empreiteiros — Exceções e Particularidades

Consulte esta referência quando estiver processando um empreiteiro específico e precisar saber de particularidades.

## Empreiteiro 01 — JVB

- **Tipo:** Construção civil
- **Formato das NFs:** Misto. NFs de 2023 usam formato "Nota Portovelhense" (layout diferente do padrão); NFs de 2024+ usam formato padrão NFSe Porto Velho
- **Nota Portovelhense:** `VALOR TOTAL DO SERVIÇO R$ xxx` aparece inline (não em linha separada); NF no campo "Número da Nota" (ex: 000000000000002/A → NF 2); Competência em campo dedicado
- **Multi-coluna pdfplumber:** ISS e INSS vêm em linhas multi-coluna (ISS = 4º valor na linha "Deduções/Base/Alíquota/ISSQN"; INSS = 3º valor na linha "PIS/COFINS/INSS/IR")
- **NF substituída:** NF 2 substitui NF 1 (ambas de 08/2023)
- **INSS:** Notas de 2023 têm INSS=0 com IR retido separadamente; notas de 2024+ têm INSS ~3.5% (Simples)
- **Pasta:** NOTA FISCAL/ com subpastas por ano (2023, 2024, 2025)

## Empreiteiro 11 — AQUILAIS ATIVIDADES DE VIGILÂNCIA

- **Tipo:** Vigilância/segurança (NÃO é construção civil)
- **Formato das NFs:** Layout com "Data Fato Gerador" (DD/MM/AAAA) em vez de "Competência"
- **INSS:** Sempre 0 — serviço executado pelo proprietário, sem folha de funcionários. Isto é NORMAL.
- **ISS:** 5% fixo (alíquota de vigilância)
- **CNO:** Ausente (não é obra de construção)
- **Pasta:** NFs soltas diretamente na raiz do empreiteiro (sem subpasta NOTA FISCAL/)
- **Nomes de arquivo:** Mistura de padrões: "NF 91...", "NFSE 91...", "NFse 94..."
- **Código de serviço:** 11.02 - Vigilância, segurança ou monitoramento

## Empreiteiro 12 — DELCINEY NOGUEIRA BRASIL

- Pasta NOTA FISCAL/ existe mas está **vazia** — nenhuma NF para processar

## Demais empreiteiros (02-10)

- **Tipo:** Construção civil
- **Formato:** NFSe padrão Porto Velho
- **Pasta:** NOTA FISCAL/ com ou sem subpastas por ano
- **INSS:** 11% (padrão cessão de mão de obra) ou ~3.5% (Simples Nacional)
- **ISS:** 1-5% conforme faixa do Simples Nacional
- **PDFs-imagem:** Meses recentes (set/2025+) frequentemente são PDFs-imagem sem texto extraível; requerem OCR ou leitura visual
