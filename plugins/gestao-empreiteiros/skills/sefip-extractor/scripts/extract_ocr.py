#!/usr/bin/env python3
"""
Extrai CAT 01 via OCR de SEFIPs escaneados — Lagoa Clube Resort.

Usa RapidOCR (pip install rapidocr-onnxruntime) — não requer Tesseract
nem qualquer dependência de sistema. PyMuPDF renderiza as páginas em
imagem e RapidOCR faz o reconhecimento de texto.

Uso:
    python3 scripts/extract_ocr.py --from-pending
    python3 scripts/extract_ocr.py --empreiteiros 1 2
    python3 scripts/extract_ocr.py --batch-size 10

Saída: state/extractions_ocr.json (merge incremental)
"""

import os, sys, re, json, argparse
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError

import fitz  # pymupdf — renderiza PDF em imagem
from rapidocr_onnxruntime import RapidOCR

# Garantir que constants.py é encontrado
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import constants

PDF_TIMEOUT = 120  # segundos por PDF (RapidOCR é mais lento que Tesseract na 1ª chamada)
OCR_DPI = 200      # resolução para renderização — 200 dpi dá boa qualidade sem ser muito pesado

# Instância global do OCR (lazy-load dos modelos na primeira chamada)
_ocr_engine = None

def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        _ocr_engine = RapidOCR()
    return _ocr_engine


def _ocr_page(doc, page_num, dpi=OCR_DPI):
    """Renderiza uma página do PDF e retorna o texto via RapidOCR."""
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = doc[page_num].get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    result, _ = _get_ocr()(img_bytes)
    if not result:
        return ""
    return " ".join(line[1] for line in result)


def _detect_rotation(doc):
    """Detecta se o PDF está rotacionado 180° comparando scores de keywords."""
    zoom = OCR_DPI / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix0 = doc[0].get_pixmap(matrix=mat)
    img_bytes_0 = pix0.tobytes("png")

    # Renderizar rotacionado 180°
    mat180 = fitz.Matrix(zoom, zoom) * fitz.Matrix(-1, 0, 0, -1, pix0.width, pix0.height)
    pix180 = doc[0].get_pixmap(matrix=mat180)
    img_bytes_180 = pix180.tobytes("png")

    ocr = _get_ocr()
    result0, _ = ocr(img_bytes_0)
    result180, _ = ocr(img_bytes_180)

    text0 = " ".join(line[1] for line in result0) if result0 else ""
    text180 = " ".join(line[1] for line in result180) if result180 else ""

    keywords = ['Empregador', 'RESUMO', 'FECHAMENTO', 'Trabalhador',
                'TOMADOR', 'CAT', 'Detalhe', 'Guia', 'FGTS', 'Digital']
    score0 = sum(1 for kw in keywords if kw in text0)
    score180 = sum(1 for kw in keywords if kw in text180)
    return 180 if score180 > score0 else 0


def extract_cat01_ocr(pdf_path, max_pages=6):
    """Extrai CAT 01 via OCR (RapidOCR)."""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return None, f"convert_error: {e}"

    if len(doc) == 0:
        doc.close()
        return None, "no_pages"

    try:
        rotation = _detect_rotation(doc)
    except Exception:
        rotation = 0

    # Se rotacionado, precisamos renderizar com matriz invertida
    zoom = OCR_DPI / 72.0
    ocr = _get_ocr()
    page_texts = []

    for i in range(min(len(doc), max_pages)):
        if rotation == 180:
            mat = fitz.Matrix(zoom, zoom)
            pix_normal = doc[i].get_pixmap(matrix=mat)
            mat180 = fitz.Matrix(zoom, zoom) * fitz.Matrix(-1, 0, 0, -1, pix_normal.width, pix_normal.height)
            pix = doc[i].get_pixmap(matrix=mat180)
        else:
            mat = fitz.Matrix(zoom, zoom)
            pix = doc[i].get_pixmap(matrix=mat)

        img_bytes = pix.tobytes("png")
        result, _ = ocr(img_bytes)
        text = " ".join(line[1] for line in result) if result else ""
        page_texts.append(text)

    doc.close()
    full_text = "\n".join(page_texts)

    # ── Formato FGTS Detalhe da Guia / Extrato ──
    if 'Detalhe' in full_text or 'Guia' in full_text:
        m = re.search(r'Qtd\.?\s*Trabalhadores\s*(?:FGTS)?:?\s*(\d+)', full_text)
        if m:
            val = int(m.group(1))
            if val > 0:
                return val, "ocr_fgts"
        # Fallback: em tabelas OCR o número pode aparecer separado do rótulo.
        # Padrão "N  Origem: Gestao de Guias" onde N é a qtd de trabalhadores.
        m2 = re.search(r'\b(\d{1,3})\s+Origem:\s*Gest[aã]o\s+de\s+Guias', full_text)
        if m2:
            val = int(m2.group(1))
            if val > 0:
                return val, "ocr_fgts"

    # ── Formato GFD — Guia do FGTS Digital ──
    if 'GFD' in full_text or 'FGTS Digital' in full_text or 'Guia do FGTS' in full_text:
        # Tabela: Competência | Qtd Trabalhadores | valores...
        m = re.search(r'(\d{2}/\d{4})\s+(\d+)\s+[\d.,]+\s', full_text)
        if m:
            val = int(m.group(2))
            if val > 0:
                return val, "ocr_fgts_digital"

    # ── Formato SEFIP Clássico ──
    for text in page_texts:
        clean = re.sub(r'[.\-/\s]', '', text)
        if constants.CNO in clean:
            m = re.search(r'\b0[1l]\s+(\d+)\s+[\d.,]+', text)
            if m:
                return int(m.group(1)), "ocr_sefip"
            m = re.search(r'TOTAIS:?\s*(\d+)', text)
            if m:
                return int(m.group(1)), "ocr_totais"

    return None, "ocr_no_match"


def find_all_sefip_pdfs(empr_num):
    """Lista todos os PDFs SEFIP de um empreiteiro, com mês inferido."""
    from extract_text import parse_month_from_path, parse_month_from_filename, select_best_pdf, find_month_pdfs
    month_pdfs = find_month_pdfs(empr_num)
    jobs = []
    for month_key in sorted(month_pdfs.keys()):
        best = select_best_pdf(month_pdfs[month_key])
        if best:
            month_str = f"{month_key[0]}-{month_key[1]:02d}"
            jobs.append((empr_num, month_str, best))
    return jobs


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--empreiteiros', nargs='+', type=int, default=None)
    parser.add_argument('--from-pending', action='store_true')
    parser.add_argument('--batch-size', type=int, default=0)
    args = parser.parse_args()

    os.makedirs(constants.STATE_DIR, exist_ok=True)

    jobs = []
    if args.from_pending:
        pending_path = os.path.join(constants.STATE_DIR, "needs_ocr.json")
        if os.path.exists(pending_path):
            pending = json.load(open(pending_path, encoding="utf-8"))
            jobs = [(j["empr"], j["month"], j["path"]) for j in pending]
            print(f"Carregados {len(jobs)} PDFs pendentes de OCR")
        else:
            print(f"Arquivo {pending_path} não encontrado. Rode extract_text.py primeiro.")
            return
    elif args.empreiteiros:
        for empr in args.empreiteiros:
            jobs.extend(find_all_sefip_pdfs(empr))
    else:
        for empr in constants.ALL_EMPR:
            jobs.extend(find_all_sefip_pdfs(empr))

    # Carregar resultados anteriores de OCR
    out_path = os.path.join(constants.STATE_DIR, "extractions_ocr.json")
    if os.path.exists(out_path):
        results = json.load(open(out_path, encoding="utf-8"))
    else:
        results = {}

    # Carregar resultados de extração por texto para evitar re-processar via OCR
    text_results = {}
    text_path = os.path.join(constants.STATE_DIR, "extractions_text.json")
    if os.path.exists(text_path):
        text_results = json.load(open(text_path, encoding="utf-8"))

    # Pular jobs já processados (por OCR ou por texto com valor válido)
    filtered = []
    for empr_num, month_str, pdf_path in jobs:
        empr_key = str(empr_num)
        # Já tem resultado OCR com valor válido para este mês?
        ocr_entry = results.get(empr_key, {}).get(month_str)
        if ocr_entry is not None and ocr_entry.get("value") is not None:
            continue
        # Já tem resultado de texto válido para este mês?
        text_val = text_results.get(empr_key, {}).get(month_str, {}).get("value")
        if text_val is not None:
            continue
        filtered.append((empr_num, month_str, pdf_path))
    skipped = len(jobs) - len(filtered)
    jobs = filtered

    if args.batch_size > 0:
        jobs = jobs[:args.batch_size]

    if skipped:
        print(f"Pulados {skipped} PDFs já processados")
    print(f"\nProcessando {len(jobs)} PDFs via OCR...\n")

    for idx, (empr_num, month_str, pdf_path) in enumerate(jobs):
        y, m = month_str.split("-")
        col = constants.MONTH_COL.get((int(y), int(m)))
        print(f"[{idx+1}/{len(jobs)}] Empr {empr_num:02d} {m}/{y} (col {col}): ", end="", flush=True)

        if not os.path.exists(pdf_path):
            print("ARQUIVO NÃO ENCONTRADO")
            continue

        # Timeout por PDF para não travar em arquivos problemáticos (cross-platform)
        try:
            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(extract_cat01_ocr, pdf_path)
                value, method = future.result(timeout=PDF_TIMEOUT)
        except FuturesTimeoutError:
            value, method = None, "ocr_timeout"
            print(f"TIMEOUT ({PDF_TIMEOUT}s) ", end="")
        except Exception as e:
            value, method = None, f"ocr_error: {e}"

        if value is not None:
            print(f"CAT01={value} [{method}]")
        else:
            print(f"FALHOU [{method}]")

        empr_key = str(empr_num)
        if empr_key not in results:
            results[empr_key] = {}
        results[empr_key][month_str] = {
            "value": value, "method": method,
            "col": col, "path": pdf_path
        }

        # Salvar após cada PDF (proteção contra timeout/interrupção)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"\nResultados salvos em {out_path}")

    total_ok = sum(1 for e in results.values() for d in e.values()
                   if d.get("value") is not None and d.get("method") != "ocr_zero_suspicious")
    total_zero = sum(1 for e in results.values() for d in e.values()
                     if d.get("method") == "ocr_zero_suspicious")
    total_fail = sum(1 for e in results.values() for d in e.values()
                     if d.get("value") is None)
    print(f"\nRESUMO OCR: {total_ok} extraídos, {total_zero} zeros suspeitos, {total_fail} falharam")


if __name__ == "__main__":
    main()
