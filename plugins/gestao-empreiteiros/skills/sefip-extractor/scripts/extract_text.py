#!/usr/bin/env python3
"""
Extrai quantidade de trabalhadores CAT 01 de SEFIPs em PDF texto — Lagoa Clube Resort.

DIFERENÇA ESTRUTURAL: Na Lagoa Clube, SEFIPs ficam dentro de cada empreiteiro:
  - XX EMPREITEIRO/SEFIP/YEAR/MM YEAR/  (empreiteiros 01, 02)
  - XX EMPREITEIRO/DOCUMENTOS MENSAIS/YEAR/MM YEAR/  (empreiteiros 03-10)

Uso:
    python3 scripts/extract_text.py                    # Processa todos
    python3 scripts/extract_text.py --empreiteiros 1 2 # Só os indicados

Saída: state/extractions_text.json
"""

import os, re, json, argparse
import pdfplumber

import constants


# ── Parsing de mês a partir do caminho ───────────────────────

def parse_month_from_path(folder_path, empr_path):
    """Determina (ano, mês) a partir da estrutura de pastas."""
    rel = os.path.relpath(folder_path, empr_path).replace("\\", "/")
    parts = [p for p in rel.split("/") if p not in (".",) and p not in constants.SEFIP_SUBFOLDERS]

    for part in parts:
        # "04 2025"
        m = re.match(r'^(\d{1,2})\s+(\d{4})$', part.strip())
        if m:
            return (int(m.group(2)), int(m.group(1)))
        # "01-2024" or "12-2024"
        m = re.match(r'^(\d{1,2})-(\d{4})$', part.strip())
        if m:
            return (int(m.group(2)), int(m.group(1)))

    # Padrão "YYYY/MM": 2024/07 → jul/2024
    for i, part in enumerate(parts):
        if re.match(r'^\d{4}$', part):
            year = int(part)
            if i + 1 < len(parts):
                m = re.match(r'^(\d{1,2})$', parts[i+1].strip())
                if m:
                    return (year, int(m.group(1)))
    return None


def parse_month_from_filename(filename):
    """Extrai (ano, mês) do nome do arquivo."""
    m = re.search(r'(\d{1,2})[-.]\s*(\d{4})', filename)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    m = re.search(r'(\d{1,2})\.(\d{4})', filename)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    return None


# ── Seleção de PDF ───────────────────────────────────────────

def _is_non_sefip(filename):
    """Identifica documentos que NÃO são SEFIP (boletos, folhas, DCTFWeb etc.)."""
    upper = filename.upper()
    non_sefip_patterns = [
        'BOLETO FGTS', 'BOLETO DE FGTS',
        'CRÉDITO INSS', 'CREDITO INSS',
        'COMPENSAÇÃO INSS', 'COMPENSACAO INSS',
        'DCTFWEB', 'DCTFWeb',
        'FOLHA DE PAGAMENTO', 'FOLHA PAGAMENTO',
        'FOLHA DE PONTO',
        'GUIA DO FGTS',  # GFD - Guia de pagamento, não é relatório SEFIP
        'HOLERITE',
        'COMPROVANTE DE DECLARAÇÃO', 'COMPROVANTE DE DECLARACAO',
        'COMPROVANTE DE PIX', 'PIX REALIZADO',
        'PROTOCOLO DE ENVIO',
        'PARCELAMENTO',  # Termo de parcelamento de débito FGTS
        'RELATÓRIO ANALÍTICO DA GPS', 'RELATORIO ANALITICO DA GPS',  # GPS previdenciário
    ]
    return any(pat in upper for pat in non_sefip_patterns)


def select_best_pdf(pdf_list):
    """Escolhe o melhor PDF da pasta baseado na prioridade documentada.

    Filtra documentos que não são SEFIP (boletos, CRÉDITO INSS, DCTFWeb etc.)
    antes de aplicar a prioridade.
    """
    # Filtrar documentos que não são SEFIP
    valid_pdfs = [p for p in pdf_list if not _is_non_sefip(os.path.basename(p))]
    if not valid_pdfs:
        return None  # Nenhum PDF relevante na pasta

    priority_keywords = [
        'relatorio re', 'relatório re',
        're.pdf',
        'sefip completa extrato fgts', 'sefip completa relatorio fgts',
        'relatorio fgts',
        'sefip completa', 'sefip comp',
        'sefip', 'sefipe',
        'fgts',
    ]
    # Filtrar PDFs com nome relevante
    sefip_pdfs = [p for p in valid_pdfs
                  if any(kw in os.path.basename(p).lower()
                         for kw in ['sefip', 'sefipe', 're.pdf', 'relatorio re',
                                    'relatório re', 'fgts', 'relatorio fgts'])]
    if not sefip_pdfs:
        sefip_pdfs = valid_pdfs

    for prio in priority_keywords:
        for p in sefip_pdfs:
            if prio in os.path.basename(p).lower():
                return p
    return sefip_pdfs[0] if sefip_pdfs else None


# ── Extração CAT 01 ─────────────────────────────────────────

def _detect_reversed_text(text):
    """Detecta se o texto está invertido (PDF rotacionado 180°).

    PDFs com rotação de 180° podem ter o texto extraído ao contrário pelo
    pdfplumber. Verificamos se palavras-chave conhecidas aparecem invertidas.
    """
    reversed_keywords = [
        'serodahlabarT',  # Trabalhadores
        'ehlateD',        # Detalhe
        'aiuG',           # Guia
        'rodagerpmE',     # Empregador
        'OTNEMAHCEF',     # FECHAMENTO
        'OMUSER',         # RESUMO
    ]
    normal_keywords = [
        'Trabalhadores', 'Detalhe', 'Guia', 'Empregador',
        'FECHAMENTO', 'RESUMO',
    ]
    reversed_score = sum(1 for kw in reversed_keywords if kw in text)
    normal_score = sum(1 for kw in normal_keywords if kw in text)
    return reversed_score > normal_score and reversed_score >= 2


def _reverse_text(text):
    """Desfaz rotação 180°: inverte caracteres de cada linha E a ordem das linhas.

    Em PDFs rotacionados 180°, pdfplumber extrai o texto de baixo para cima
    e cada linha é lida da direita para a esquerda. Precisamos reverter ambos.
    Após reverter, colapsa linhas curtas consecutivas (< 30 chars) em linhas
    maiores para reconstruir sentenças fragmentadas.
    """
    lines = text.split("\n")
    reversed_lines = [line[::-1] for line in lines]
    reversed_lines.reverse()

    # Colapsar linhas curtas consecutivas em linhas maiores
    collapsed = []
    current = ""
    for line in reversed_lines:
        stripped = line.strip()
        if not stripped:
            if current:
                collapsed.append(current)
                current = ""
            continue
        if len(stripped) < 30:
            current = (current + " " + stripped) if current else stripped
        else:
            if current:
                collapsed.append(current)
            collapsed.append(stripped)
            current = ""
    if current:
        collapsed.append(current)

    return "\n".join(collapsed)


def _extract_fgts_digital(pages, full):
    """Extrai trabalhadores de FGTS Digital (GFD).

    Formato: tabela com colunas Competência | Trabalhadores | FGTS Mensal | ...
    Padrão: "MM/YYYY  N  valor" onde N é a quantidade de trabalhadores.
    Retorna (value, method) ou (None, None) se não for FGTS Digital.
    """
    if 'Guia do FGTS Digital' not in full and 'GFD' not in full:
        return None, None

    # Padrão: competência seguida de quantidade de trabalhadores e valor monetário
    # Ex: "08/2024 7 1.024,80" ou "04/2025 3 429,76"
    m = re.search(r'(\d{2}/\d{4})\s+(\d+)\s+[\d.,]+\s', full)
    if m:
        return int(m.group(2)), "fgts_digital"

    # Padrão alternativo: "Quantidade\nCompetência Trabalhadores"
    m = re.search(r'Trabalhadores.*?\n.*?(\d{2}/\d{4})\s+(\d+)', full, re.DOTALL)
    if m:
        return int(m.group(2)), "fgts_digital"

    return None, None


def extract_cat01(pdf_path):
    """Extrai CAT 01 de um SEFIP em PDF texto.

    Reconhece 4 formatos:
    1. FGTS Extrato / Detalhe da Guia (Gestão de Guias)
    2. SEFIP Clássico (Resumo do Fechamento)
    3. FGTS Digital (GFD - novo formato)
    4. PDFs com texto invertido (rotação 180°)
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = [p.extract_text() or "" for p in pdf.pages]
            full = "\n".join(pages)

            if not full.strip():
                return None, "empty_text"

            # Texto muito curto pode indicar PDF escaneado com marca d'água
            if len(full.strip()) < 50:
                return None, "empty_text"

            # ── Detectar texto invertido (PDF rotacionado 180°) ──
            if _detect_reversed_text(full):
                pages = [_reverse_text(p) for p in pages]
                full = "\n".join(pages)

            # ── Formato FGTS Digital (GFD) ──
            val, method = _extract_fgts_digital(pages, full)
            if val is not None:
                return val, method

            # ── Formato FGTS Extrato / Detalhe da Guia ──
            if 'Detalhe da Guia' in full or 'Relatório da Guia' in full:
                # Tentar extrair do tomador CNO específico (evita pegar total do empregador)
                cno_clean = constants.CNO
                for text in pages:
                    text_clean = text.replace('.', '').replace('/', '').replace('-', '').replace(' ', '')
                    if cno_clean in text_clean:
                        # Procura "Qtd. Trabalhadores" na mesma página/bloco do tomador CNO
                        m_tom = re.search(r'Qtd\.?\s*Trabalhadores:?\s*(\d+)', text)
                        if m_tom:
                            return int(m_tom.group(1)), "fgts_detalhe_tomador"
                # Fallback: valor global (pode ser o total do empregador)
                m = re.search(r'Qtd\.?\s*Trabalhadores(?:\s+FGTS)?:\s*(\d+)', full)
                if m:
                    return int(m.group(1)), "fgts_extrato"

            # ── Formato SEFIP Clássico — busca bloco do tomador com nosso CNO ──
            for text in pages:
                clean = text.replace('.', '').replace('/', '').replace('-', '')
                if 'RESUMO DO FECHAMENTO' in text and constants.CNO in clean:
                    m = re.search(r'\b01\s+(\d+)\s+[\d.,]+', text)
                    if m:
                        return int(m.group(1)), "sefip_classico"
                    m = re.search(r'TOTAIS:\s*(\d+)', text)
                    if m:
                        return int(m.group(1)), "sefip_totais"

            # ── Busca mais ampla: qualquer página com nosso CNO ──
            for i, text in enumerate(pages):
                clean = text.replace('.', '').replace('/', '').replace('-', '')
                if constants.CNO in clean:
                    combined = "\n".join(pages[max(0,i-1):min(len(pages),i+2)])
                    m = re.search(r'\b01\s+(\d+)\s+[\d.,]+', combined)
                    if m:
                        return int(m.group(1)), "nearby_page"

            return None, "no_match"
    except Exception as e:
        return None, f"error: {e}"


# ── Varredura de pastas ──────────────────────────────────────

def find_month_pdfs(empr_num):
    """Descobre todos os PDFs SEFIP de um empreiteiro.

    Procura em DOIS locais:
    1. XX EMPREITEIRO/SEFIP/...
    2. XX EMPREITEIRO/DOCUMENTOS MENSAIS/...
    """
    folder = constants.EMPR_FOLDERS.get(empr_num)
    if not folder:
        return {}
    empr_path = os.path.join(constants.BASE_DIR, folder)
    if not os.path.isdir(empr_path):
        return {}

    month_pdfs = {}

    # Locais para procurar SEFIPs (configurável via obra.json)
    search_paths = []
    for sub_name in constants.SEFIP_SUBFOLDERS:
        sub_path = os.path.join(empr_path, sub_name)
        if os.path.isdir(sub_path):
            search_paths.append(sub_path)

    for search_path in search_paths:
        for root, dirs, files in os.walk(search_path):
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            if not pdfs:
                continue
            folder_month = parse_month_from_path(root, empr_path)
            for pdf in pdfs:
                pdf_path = os.path.join(root, pdf)
                month = folder_month or parse_month_from_filename(pdf)
                if month is None:
                    continue
                month_pdfs.setdefault(month, []).append(pdf_path)

    return month_pdfs


# ── Main ─────────────────────────────────────────────────────

def load_existing(path):
    """Carrega JSON existente para merge. Garante que retorna dict."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except (json.JSONDecodeError, IOError):
            pass
    return {}


def load_existing_ocr(path):
    """Carrega lista de needs_ocr existente. Garante que retorna list."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except (json.JSONDecodeError, IOError):
            pass
    return []


def merge_results(existing, new_batch):
    """Merge new_batch into existing, sobrescrevendo por empreiteiro processado."""
    merged = dict(existing)
    for empr_key, months in new_batch.items():
        merged[empr_key] = months  # substitui o empreiteiro inteiro (re-processado)
    return merged


def merge_ocr_lists(existing_list, new_list, processed_emprs):
    """Remove entradas dos empreiteiros reprocessados e adiciona novas."""
    # Manter entradas de empreiteiros que NÃO foram reprocessados
    kept = [e for e in existing_list if e.get("empr") not in processed_emprs]
    # Adicionar novas
    kept.extend(new_list)
    return kept


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--empreiteiros', nargs='+', type=int, default=constants.ALL_EMPR)
    parser.add_argument('--force', action='store_true',
                        help='Reprocessa tudo do zero, ignorando estado anterior')
    args = parser.parse_args()

    os.makedirs(constants.STATE_DIR, exist_ok=True)

    out_path = os.path.join(constants.STATE_DIR, "extractions_text.json")
    ocr_path = os.path.join(constants.STATE_DIR, "needs_ocr.json")

    if args.force:
        print("⚡ Modo --force: reprocessando tudo do zero\n")
        # Carregar estado existente — o merge_results vai substituir
        # apenas os empreiteiros reprocessados, preservando os demais
        existing_results = load_existing(out_path)
        existing_ocr = load_existing_ocr(ocr_path)
        # Limpar dados dos empreiteiros que serão reprocessados
        force_emprs = set(args.empreiteiros)
        for ek in list(existing_results.keys()):
            if int(ek) in force_emprs:
                del existing_results[ek]
        existing_ocr = [e for e in existing_ocr if e.get("empr") not in force_emprs]
    else:
        # Carregar estado anterior para merge
        existing_results = load_existing(out_path)
        existing_ocr = load_existing_ocr(ocr_path)

    batch_results = {}
    batch_ocr = []
    processed_emprs = set()

    for empr_num in sorted(args.empreiteiros):
        folder_name = constants.EMPR_FOLDERS.get(empr_num, f"Empr {empr_num}")
        processed_emprs.add(empr_num)
        print(f"\n--- Empreiteiro {empr_num:02d}: {folder_name} ---")
        month_pdfs = find_month_pdfs(empr_num)
        if not month_pdfs:
            print("  Nenhum PDF SEFIP encontrado")
            batch_results[str(empr_num)] = {}
            continue

        batch_results[str(empr_num)] = {}
        for month_key in sorted(month_pdfs.keys()):
            best_pdf = select_best_pdf(month_pdfs[month_key])
            if not best_pdf:
                continue
            value, method = extract_cat01(best_pdf)
            col = constants.MONTH_COL.get(month_key)
            month_str = f"{month_key[0]}-{month_key[1]:02d}"

            if value is not None:
                print(f"  {month_key[1]:02d}/{month_key[0]} -> CAT01={value} [{method}] (col {col})")
                batch_results[str(empr_num)][month_str] = {
                    "value": value, "method": method,
                    "col": col, "path": best_pdf
                }
            elif method == "empty_text":
                print(f"  {month_key[1]:02d}/{month_key[0]} -> PDF escaneado, precisa OCR")
                batch_ocr.append({"empr": empr_num, "month": month_str, "path": best_pdf})
            else:
                print(f"  {month_key[1]:02d}/{month_key[0]} -> FALHOU [{method}] {os.path.basename(best_pdf)}")
                batch_results[str(empr_num)][month_str] = {
                    "value": None, "method": method,
                    "col": col, "path": best_pdf
                }

    # Merge e salvar
    merged = merge_results(existing_results, batch_results)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2, ensure_ascii=False)
    print(f"\nResultados salvos em {out_path} ({len(merged)} empreiteiros)")

    merged_ocr = merge_ocr_lists(existing_ocr, batch_ocr, processed_emprs)
    if merged_ocr:
        with open(ocr_path, "w", encoding="utf-8") as f:
            json.dump(merged_ocr, f, indent=2, ensure_ascii=False)
        print(f"{len(merged_ocr)} PDFs precisam de OCR → {ocr_path}")
    elif os.path.exists(ocr_path):
        with open(ocr_path, "w", encoding="utf-8") as f:
            json.dump([], f)

    total_ok = sum(1 for e in merged.values() for d in e.values() if d.get("value") is not None)
    total_fail = sum(1 for e in merged.values() for d in e.values() if d.get("value") is None)
    total_ocr = len(merged_ocr)
    print(f"\nRESUMO TOTAL: {total_ok} extraídos, {total_fail} falharam, {total_ocr} precisam OCR")


if __name__ == "__main__":
    main()
