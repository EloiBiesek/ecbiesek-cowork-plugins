#!/usr/bin/env python3
"""
Agente de extração SEFIP por empreiteiro — Lagoa Clube Resort.

Comportamento padrão: INCREMENTAL — pula meses já extraídos com sucesso.
Use --force-reprocess para re-extrair tudo do zero.

Uso:
    python3 agent_sefip_extractor.py --empreiteiro 1
    python3 agent_sefip_extractor.py --empreiteiro 3 --force-reprocess
"""

import os, sys, json, argparse
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)

import constants
from extract_text import find_month_pdfs, select_best_pdf, extract_cat01

PDF_TIMEOUT = 120


def load_known_state(empr_key):
    """Retorna dict {month_str: entry} com extrações bem-sucedidas já conhecidas."""
    known = {}
    for fname in ("extractions_text.json", "extractions_ocr.json"):
        path = os.path.join(constants.STATE_DIR, fname)
        if not os.path.exists(path):
            continue
        try:
            data = json.load(open(path, encoding="utf-8"))
            for month_str, entry in data.get(empr_key, {}).items():
                if entry.get("value") is not None:
                    known[month_str] = entry
        except Exception:
            pass
    return known


def run_ocr(pdf_path):
    try:
        from extract_ocr import extract_cat01_ocr
        return extract_cat01_ocr(pdf_path)
    except Exception as e:
        return None, f"ocr_error:{e}"


def process_empreiteiro(empr_num, force_reprocess=False):
    name     = constants.EMPR_FOLDERS.get(empr_num, f"Empr {empr_num}")
    empr_key = str(empr_num)

    print(f"\n{'='*60}")
    print(f"AGENTE {empr_num:02d}: {name}")
    mode = "COMPLETO (--force-reprocess)" if force_reprocess else "INCREMENTAL"
    print(f"Modo: {mode}")
    print(f"{'='*60}")

    # ── Carregar estado existente ─────────────────────────────
    known = {} if force_reprocess else load_known_state(empr_key)
    if known:
        print(f"  {len(known)} meses já extraídos no estado — serão pulados")

    results_text = {}
    needs_ocr    = []
    skipped      = {}   # meses pulados por já terem valor

    # ── 1. Extração de texto ──────────────────────────────────
    month_pdfs = find_month_pdfs(empr_num)
    if not month_pdfs:
        print("  Nenhum PDF SEFIP encontrado")
    else:
        print(f"  {len(month_pdfs)} meses encontrados nos PDFs")

    for month_key in sorted(month_pdfs.keys()):
        best_pdf  = select_best_pdf(month_pdfs[month_key])
        if not best_pdf:
            continue

        col       = constants.MONTH_COL.get(month_key)
        month_str = f"{month_key[0]}-{month_key[1]:02d}"

        # ── Skip incremental ──────────────────────────────────
        if month_str in known:
            entry = known[month_str]
            print(f"  ↷ {month_key[1]:02d}/{month_key[0]} → PULADO (já ok: CAT01={entry['value']} [{entry['method']}])")
            skipped[month_str] = entry
            results_text[month_str] = entry   # reaproveita para saída
            continue

        value, method = extract_cat01(best_pdf)

        if value is not None:
            print(f"  ✓ {month_key[1]:02d}/{month_key[0]} → CAT01={value} [{method}] col {col}")
            results_text[month_str] = {
                "value": value, "method": method,
                "col": col, "path": best_pdf
            }
        elif method == "empty_text":
            print(f"  ⊙ {month_key[1]:02d}/{month_key[0]} → escaneado, precisa OCR")
            needs_ocr.append({
                "empr": empr_num, "month": month_str,
                "path": best_pdf, "col": col
            })
        else:
            print(f"  ✗ {month_key[1]:02d}/{month_key[0]} → FALHOU [{method}]  {os.path.basename(best_pdf)}")
            results_text[month_str] = {
                "value": None, "method": method,
                "col": col, "path": best_pdf
            }

    # ── 2. OCR ────────────────────────────────────────────────
    ocr_results = {}

    # Verificar quais dos pendentes de OCR já estão no estado
    needs_ocr_filtered = []
    for item in needs_ocr:
        month_str = item["month"]
        if month_str in known:
            entry = known[month_str]
            print(f"  ↷ {month_str} → OCR PULADO (já ok: CAT01={entry['value']})")
            skipped[month_str] = entry
            ocr_results[month_str] = entry
        else:
            needs_ocr_filtered.append(item)
    needs_ocr = needs_ocr_filtered

    if needs_ocr:
        print(f"\n  Iniciando OCR em {len(needs_ocr)} PDFs...")
        for item in needs_ocr:
            pdf_path  = item["path"]
            month_str = item["month"]
            col       = item["col"]
            fname     = os.path.basename(pdf_path)

            if not os.path.exists(pdf_path):
                print(f"  ✗ {month_str} → arquivo não encontrado: {fname}")
                ocr_results[month_str] = {
                    "value": None, "method": "file_not_found",
                    "col": col, "path": pdf_path
                }
                continue

            try:
                with ThreadPoolExecutor(max_workers=1) as executor:
                    future = executor.submit(run_ocr, pdf_path)
                    value, method = future.result(timeout=PDF_TIMEOUT)
            except FuturesTimeoutError:
                value, method = None, f"ocr_timeout_{PDF_TIMEOUT}s"
            except Exception as e:
                value, method = None, f"ocr_exception:{e}"

            mark = "✓" if value is not None else "✗"
            msg  = f"CAT01={value} [{method}]" if value is not None else f"FALHOU [{method}]"
            print(f"  {mark} {month_str} → {msg}  {fname}")

            ocr_results[month_str] = {
                "value": value, "method": method,
                "col": col, "path": pdf_path
            }

    # ── 3. Estatísticas ───────────────────────────────────────
    all_results = {**results_text, **ocr_results}
    novos_text  = sum(1 for m, v in results_text.items()
                      if m not in skipped and v.get("value") is not None)
    novos_ocr   = sum(1 for m, v in ocr_results.items()
                      if m not in skipped and v.get("value") is not None)
    n_pulados   = len(skipped)
    n_erros     = sum(1 for v in all_results.values() if v.get("value") is None)
    ok_months   = sorted(m for m, v in all_results.items() if v.get("value") is not None)
    comp_inicial = ok_months[0]  if ok_months else None
    comp_final   = ok_months[-1] if ok_months else None

    print(f"\n  RESUMO {empr_num:02d} │ novos texto={novos_text} │ novos OCR={novos_ocr} "
          f"│ pulados={n_pulados} │ erros={n_erros}")
    if comp_final:
        print(f"  Cobertura: {comp_inicial} → {comp_final} ({len(ok_months)} meses ok)")

    # ── 4. Salvar arquivo exclusivo ───────────────────────────
    os.makedirs(constants.STATE_DIR, exist_ok=True)
    out_path = os.path.join(constants.STATE_DIR, f"agent_results_empr_{empr_num:02d}.json")

    output = {
        "empr":          empr_num,
        "name":          name,
        "text_results":  results_text,
        "ocr_results":   ocr_results,
        # Estatísticas para o log do RESUMO
        "stats": {
            "novos_texto":    novos_text,
            "novos_ocr":      novos_ocr,
            "pulados":        n_pulados,
            "erros":          n_erros,
            "comp_inicial":   comp_inicial,
            "comp_final":     comp_final,
            "total_ok":       len(ok_months),
        }
    }

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"  Salvo → {out_path}")
    return output


def main():
    parser = argparse.ArgumentParser(description="Agente SEFIP incremental por empreiteiro")
    parser.add_argument("--empreiteiro",     type=int, required=True)
    parser.add_argument("--force-reprocess", action="store_true",
                        help="Re-extrai tudo, ignorando estado existente")
    args = parser.parse_args()

    if args.empreiteiro not in range(1, 13):
        print(f"Empreiteiro inválido: {args.empreiteiro}. Use 1-12.")
        sys.exit(1)

    process_empreiteiro(args.empreiteiro, force_reprocess=args.force_reprocess)
    print("\n[AGENTE CONCLUÍDO]")


if __name__ == "__main__":
    main()
