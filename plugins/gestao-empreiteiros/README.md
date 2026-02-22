# Gestão de Empreiteiros

Plugin para Claude Cowork que automatiza a extração e consolidação de documentos de empreiteiros em obras de construção civil.

## Skills incluídas

### sefip-extractor
Extrai a quantidade de trabalhadores **CAT 01** dos relatórios SEFIP/GFIP de cada empreiteiro e preenche a aba **"Alocação de colaboradores"** da planilha de controle.

**Formatos reconhecidos:**
- Extrato FGTS (Qtd. Trabalhadores)
- SEFIP Clássico (Resumo do Fechamento)
- Detalhe da Guia FGTS (Gestão de Guias)
- FGTS Digital (GFD)
- PDFs rotacionados 180° (correção automática)

### extrator-de-nfs
Extrai dados de Notas Fiscais de Serviço Eletrônicas (NFSe) dos empreiteiros e consolida na planilha de controle.

**Formatos reconhecidos:**
- NFSe padrão Porto Velho
- Nota Portovelhense (NFs antigas)
- Formato Vigilância (ex: AQUILAIS)

## Requisitos

```bash
pip install pdfplumber openpyxl pymupdf Pillow
# OCR (opcional):
pip install rapidocr-onnxruntime
```

## Estrutura esperada de pastas da obra

```
PRESTADORES DE SERVIÇO/
  01 EMPREITEIRO A/
    SEFIP/2023/ 2024/ 2025/
    NOTA FISCAL/2023/ 2024/ 2025/
  02 EMPREITEIRO B/
    DOCUMENTOS MENSAIS/
      2024/07 2024/ 08 2024/ ...
    NOTA FISCAL/
  ...
```

As skills detectam automaticamente pastas de obra com empreiteiros numerados (`01 NOME`, `02 NOME`, ...) e processam uma obra por vez.

## Uso

Após instalar o plugin, basta pedir ao Claude:
- "Processar SEFIP da obra X"
- "Atualizar notas fiscais dos empreiteiros"
- "Verificar cobertura de competências"
- "Completar dados pendentes de NFSe"
