#!/usr/bin/env python3
"""
Resolve divergências entre planilha e PDFs — Lagoa Clube Resort.

Uso:
    python3 scripts/resolve_divergences.py --list
    python3 scripts/resolve_divergences.py --accept-pdf 1 3 5
    python3 scripts/resolve_divergences.py --keep-planilha 2 4
    python3 scripts/resolve_divergences.py --apply
"""

import os, json, argparse, sys
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import constants


def load_divergences():
    if not os.path.exists(os.path.join(constants.STATE_DIR, "divergences.json")):
        print("Nenhuma divergência registrada. Rode update_planilha.py primeiro.")
        sys.exit(0)
    return json.load(open(os.path.join(constants.STATE_DIR, "divergences.json")))


def save_divergences(divs):
    with open(os.path.join(constants.STATE_DIR, "divergences.json"), "w") as f:
        json.dump(divs, f, indent=2, ensure_ascii=False)


def list_divergences(divs):
    pending = [d for d in divs if not d.get("resolved")]
    resolved = [d for d in divs if d.get("resolved")]

    if pending:
        print(f"\n{'='*70}")
        print(f"DIVERGÊNCIAS PENDENTES ({len(pending)})")
        print(f"{'='*70}")
        for i, d in enumerate(divs):
            if d.get("resolved"):
                continue
            print(f"\n  [{i}] Empreiteiro {d['empr']:02d} — {d.get('name', '')}")
            print(f"      Mês: {d['month']}  |  Célula: {d['col']}{d['row']}")
            print(f"      Planilha: {d['planilha']}  |  PDF: {d['pdf']}")
            print(f"      Método: {d.get('method', '?')}")
            if d.get('path'):
                print(f"      Arquivo: ...{d['path'][-80:]}")

    if resolved:
        print(f"\n  Já resolvidas: {len(resolved)}")
        for d in resolved:
            print(f"    Empr {d['empr']:02d} {d['month']}: {d['resolution']}")

    if not pending:
        print("\n✓ Todas as divergências foram resolvidas!")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--list', action='store_true')
    parser.add_argument('--accept-pdf', nargs='+', type=int, default=[])
    parser.add_argument('--keep-planilha', nargs='+', type=int, default=[])
    parser.add_argument('--accept-all-pdf', action='store_true')
    parser.add_argument('--keep-all-planilha', action='store_true')
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()

    divs = load_divergences()

    if args.list or (not any([args.accept_pdf, args.keep_planilha,
                              args.accept_all_pdf, args.keep_all_planilha, args.apply])):
        list_divergences(divs)
        return

    changed = False

    if args.accept_all_pdf:
        for d in divs:
            if not d.get("resolved"):
                d["resolved"] = True
                d["resolution"] = "accept_pdf"
                changed = True

    if args.keep_all_planilha:
        for d in divs:
            if not d.get("resolved"):
                d["resolved"] = True
                d["resolution"] = "keep_planilha"
                changed = True

    for idx in args.accept_pdf:
        if 0 <= idx < len(divs) and not divs[idx].get("resolved"):
            divs[idx]["resolved"] = True
            divs[idx]["resolution"] = "accept_pdf"
            changed = True
            print(f"  [{idx}] Empr {divs[idx]['empr']:02d} {divs[idx]['month']}: aceito PDF ({divs[idx]['pdf']})")

    for idx in args.keep_planilha:
        if 0 <= idx < len(divs) and not divs[idx].get("resolved"):
            divs[idx]["resolved"] = True
            divs[idx]["resolution"] = "keep_planilha"
            changed = True
            print(f"  [{idx}] Empr {divs[idx]['empr']:02d} {divs[idx]['month']}: mantida planilha ({divs[idx]['planilha']})")

    if changed:
        save_divergences(divs)
        print(f"\nDivergências atualizadas em {os.path.join(constants.STATE_DIR, "divergences.json")}")

    if args.apply:
        to_apply = [d for d in divs if d.get("resolved") and d["resolution"] == "accept_pdf"]
        if not to_apply:
            print("\nNenhuma divergência com 'accept_pdf' para aplicar.")
            return

        wb = openpyxl.load_workbook(constants.XLSX_PATH)
        ws = wb['Alocação de colaboradores']
        applied = 0

        for d in to_apply:
            row = d["row"]
            col_idx = openpyxl.utils.column_index_from_string(d["col"])
            ws.cell(row=row, column=col_idx, value=d["pdf"])
            applied += 1
            print(f"  Empr {d['empr']:02d} {d['month']} {d['col']}{row}: {d['planilha']} → {d['pdf']}")

        wb.save(constants.XLSX_PATH)
        print(f"\n✓ {applied} células atualizadas.")


if __name__ == "__main__":
    main()
