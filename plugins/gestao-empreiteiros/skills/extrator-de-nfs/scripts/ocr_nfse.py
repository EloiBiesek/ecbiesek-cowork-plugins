#!/usr/bin/env python3
"""
OCR batch processor for NFSe PDFs — genérico (qualquer obra).

Usa RapidOCR + PyMuPDF (não requer Tesseract nem Poppler).
Renderiza páginas do PDF em imagem via fitz, detecta rotação 0°/180°,
e extrai dados estruturados usando regex para formato NFSe Porto Velho/RO.

Usage:
    python3 ocr_nfse.py [--obra-dir PATH] [--json PATH] [--batch-size 15] [--start 0]

Requires: rapidocr-onnxruntime, pymupdf
    pip install rapidocr-onnxruntime pymupdf
"""
import json, re, os, sys, argparse

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)
import constants_nfse

try:
    import fitz  # pymupdf
    from rapidocr_onnxruntime import RapidOCR
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

OCR_DPI = 200
_ocr_engine = None


def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        _ocr_engine = RapidOCR()
    return _ocr_engine


def _ocr_page(doc, page_num, dpi=OCR_DPI, rotation=0):
    """Renderiza uma página do PDF e retorna o texto via RapidOCR."""
    zoom = dpi / 72.0
    if rotation == 180:
        pix0 = doc[page_num].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        mat180 = fitz.Matrix(zoom, zoom) * fitz.Matrix(-1, 0, 0, -1, pix0.width, pix0.height)
        pix = doc[page_num].get_pixmap(matrix=mat180)
    else:
        pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img_bytes = pix.tobytes("png")
    result, _ = _get_ocr()(img_bytes)
    if not result:
        return ""
    return " ".join(line[1] for line in result)


def _detect_rotation(doc):
    """Detecta se o PDF está rotacionado 180° comparando scores de keywords NFSe."""
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

    keywords = ['NOTA', 'FISCAL', 'SERVI', 'PRESTADOR', 'TOMADOR', 'CNPJ',
                'VALOR', 'ISS', 'INSS', 'Compet']
    score0 = sum(1 for kw in keywords if kw in text0)
    score180 = sum(1 for kw in keywords if kw in text180)
    return 180 if score180 > score0 else 0


def ocr_pdf(pdf_path, max_pages=4):
    """Renderiza PDF via PyMuPDF, detecta rotação, e extrai texto via RapidOCR."""
    doc = fitz.open(pdf_path)
    n_pages = min(len(doc), max_pages)
    if n_pages == 0:
        doc.close()
        return ""

    rotation = _detect_rotation(doc)
    all_text = []
    for i in range(n_pages):
        text = _ocr_page(doc, i, rotation=rotation)
        all_text.append(text)
    doc.close()
    return "\n".join(all_text)


def parse_br_number(s):
    """Parse Brazilian number format: 320.000,00 -> 320000.00"""
    s = s.strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None


def extract_from_ocr(text, filename):
    """Extract NFSe fields from OCR text using regex patterns."""
    result = {
        "nf": None, "razao_social": None, "cnpj_prestador": None,
        "cno": None, "competencia": None, "valor_total": None,
        "inss": None, "iss": None
    }
    lines = text.split('\n')

    # NF
    for i, line in enumerate(lines):
        if any(kw in line for kw in ['Recolhimento', 'Recolhiment']):
            m = re.search(r'(\d+)\s*$', line)
            if m:
                nf = int(m.group(1))
                if 0 < nf < 100000:
                    result["nf"] = nf
                    break
        if re.search(r'N[°o?]?\s*da\s*Nota', line, re.I):
            for j in range(i + 1, min(i + 3, len(lines))):
                m = re.search(r'(\d+)\s*$', lines[j])
                if m:
                    nf = int(m.group(1))
                    if 0 < nf < 100000:
                        result["nf"] = nf
                        break
            if result["nf"]:
                break

    if not result["nf"]:
        m = re.search(r'(?:NFSE?|NF)\s*(\d+)', filename, re.I)
        if m:
            result["nf"] = int(m.group(1))

    # Razao Social
    m = re.search(r'(?:Raz[ãa]o\s*Social|Razao\s*Social)[:\s]*([^\n]+)', text, re.I)
    if m:
        val = re.sub(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}.*', '', m.group(1)).strip()
        if len(val) > 3:
            result["razao_social"] = val

    # CNPJ
    section = ""
    m_p = re.search(r'PRESTADOR(.*?)TOMADOR', text, re.I | re.S)
    if m_p:
        section = m_p.group(1)
    cnpjs = re.findall(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', section or text)
    if cnpjs:
        result["cnpj_prestador"] = cnpjs[0]

    # CNO — dinâmico via constants_nfse.CNO
    m = re.search(r'(?:CNO|C\.N\.O)[:\s(]*([0-9][0-9.\-/]+)', text, re.I)
    if not m:
        cno_digits = re.sub(r'[^0-9]', '', constants_nfse.CNO)
        if len(cno_digits) >= 6:
            # Buscar por prefixo flexível (OCR pode ter espaços)
            prefix = cno_digits[:6]
            pat = r'[\.\s]*'.join(prefix)
            m = re.search(r'(' + pat + r'[0-9.\-/\s]*)', text)
    if m:
        result["cno"] = re.sub(r'\s+', '', m.group(1)).strip().rstrip(')')

    # Competencia
    for pat in [
        r'Compet[êe]ncia[:\s]*(\d{1,2})[/\-](\d{4})',
        r'Per[ií]odo[:\s]*(\d{1,2})[/\-](\d{4})',
        r'(?:MES|COMP|MÊS|COMPETENCIA)\s*(\d{1,2})\s*[-/]\s*(\d{4})',
        r'referente.*?(\d{1,2})\s*[-/]\s*(\d{4})'
    ]:
        m = re.search(pat, text, re.I)
        if m:
            mo, yr = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 2020 <= yr <= 2030:
                result["competencia"] = f"{mo:02d}/{yr}"
                break

    if not result["competencia"]:
        m = re.search(r'(?:COMP\s*)?(\d{1,2})\s*[-]\s*(\d{4})', filename, re.I)
        if m:
            mo, yr = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 2020 <= yr <= 2030:
                result["competencia"] = f"{mo:02d}/{yr}"

    # Valor Total + ISS
    for i, line in enumerate(lines):
        if re.search(r'VALOR\s*SERVI', line, re.I) and re.search(r'ISS', line, re.I):
            if i + 1 < len(lines):
                nums = re.findall(r'[\d]+(?:\.[\d]+)*(?:,[\d]+)?', lines[i + 1])
                if nums:
                    v = parse_br_number(nums[0])
                    if v and v > 0:
                        result["valor_total"] = round(v, 2)
                    if len(nums) >= 2:
                        iss = parse_br_number(nums[-1])
                        if iss and iss > 0:
                            result["iss"] = round(iss, 2)
            break

    # INSS
    for i, line in enumerate(lines):
        if re.search(r'INSS', line, re.I) and re.search(r'IR', line, re.I):
            if i + 1 < len(lines):
                nums = re.findall(r'[\d]+(?:\.[\d]+)*(?:,[\d]+)?', lines[i + 1])
                if nums:
                    v = parse_br_number(nums[0])
                    if v is not None and v >= 0:
                        result["inss"] = round(v, 2)
            break

    return result


def main():
    if not HAS_OCR:
        print("ERRO: Dependências de OCR não instaladas.")
        print("  pip install rapidocr-onnxruntime pymupdf")
        sys.exit(1)

    parser = argparse.ArgumentParser(description="OCR batch processor — NFSe")
    parser.add_argument("--obra-dir",
                        help="Caminho para a pasta da obra")
    parser.add_argument("--json",
                        help="JSON de extração (default: <obra>/.nfse-state/nfse_extracted.json)")
    parser.add_argument("--batch-size", type=int, default=15)
    parser.add_argument("--start", type=int, default=0)
    args = parser.parse_args()

    # Inicializar obra
    obra_dir = args.obra_dir
    if not obra_dir:
        if constants_nfse.is_obra_dir(os.getcwd()):
            obra_dir = os.getcwd()
        else:
            obra_dir = constants_nfse.BASE_DIR
    constants_nfse.init_obra(obra_dir)

    json_path = args.json or os.path.join(constants_nfse.STATE_DIR, "nfse_extracted.json")

    if not os.path.exists(json_path):
        print(f"ERRO: JSON não encontrado: {json_path}")
        print("  Rode extract_all_nfse.py primeiro.")
        sys.exit(1)

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    records = data["records"]

    work_indices = [
        i for i, r in enumerate(records)
        if not (r.get("nf") and r.get("valor_total") and r.get("competencia"))
    ]

    batch = work_indices[args.start:args.start + args.batch_size]
    print(f"Obra: {constants_nfse.OBRA_NOME}")
    print(f"Processando {len(batch)} registros "
          f"(índices {args.start} a {args.start + args.batch_size - 1} "
          f"dos {len(work_indices)} pendentes)")

    if not batch:
        print("Nenhum registro pendente neste intervalo.")
        return

    updated = 0
    for count, idx in enumerate(batch):
        rec = records[idx]
        pdf_path = rec.get("arquivo_path", "")
        if not pdf_path or not os.path.exists(pdf_path):
            print(f"  [{count+1}/{len(batch)}] ARQUIVO NÃO ENCONTRADO: {pdf_path}")
            continue

        fname = rec.get("arquivo", os.path.basename(pdf_path))
        print(f"  [{count+1}/{len(batch)}] {fname[:60]}...", end=" ", flush=True)

        try:
            text = ocr_pdf(pdf_path)
            if not text.strip():
                print("SEM TEXTO")
                continue

            extracted = extract_from_ocr(text, fname)
            changes = []
            for field in ["nf", "razao_social", "cnpj_prestador", "cno",
                          "competencia", "valor_total", "inss", "iss"]:
                if rec.get(field) is None and extracted.get(field) is not None:
                    rec[field] = extracted[field]
                    changes.append(field)

            if changes:
                updated += 1
                missing = [f for f in ["nf", "valor_total", "competencia"] if not rec.get(f)]
                obs = []
                if missing:
                    obs.append(f"Campos faltantes: {', '.join(missing)}")
                obs.append("OCR (RapidOCR)")
                rec["observacao"] = ". ".join(obs)
                print(f"OK ({', '.join(changes)})")
            else:
                print("sem novos dados")
        except Exception as e:
            print(f"ERRO: {str(e)[:50]}")

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    complete = sum(
        1 for r in records
        if r.get("nf") and r.get("valor_total") and r.get("competencia")
    )
    print(f"\nAtualizados: {updated} | Completos: {complete}/{len(records)} | JSON salvo em {json_path}")


if __name__ == "__main__":
    main()
