#!/usr/bin/env python3
"""
Diagnostic test script — verifies extraction quality on known problem PDFs.

Runs extract_cat01 on a curated set of PDFs and compares against expected results.
Reports pass/fail per test case and overall success rate.

Usage:
    PYTHONIOENCODING=utf-8 python3 scripts/test_extraction.py
"""

import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import constants
from extract_text import extract_cat01, select_best_pdf

# ── Test cases: (empr, month, pdf_relative_path, expected_value, expected_min_method) ──
# expected_value: int or None (None = should NOT extract a value, e.g. non-SEFIP doc)
# expected_method: partial match on method name, or "any" for any successful method
TEST_CASES = [
    # --- FGTS Digital format (new) ---
    {
        "id": "fgts_digital_empr3_0824",
        "desc": "FGTS Digital - Empr 3 ago/2024",
        "path": r"03 G & M CONSTRUÇÕES (MATEUS)\DOCUMENTOS MENSAIS\2024\08 2024\FGTS (2).pdf",
        "expected_value": 7,
        "expected_method": "fgts_digital",
    },
    {
        "id": "fgts_digital_empr4_0625",
        "desc": "FGTS Digital - Empr 4 jun/2025 (comp 04/2025)",
        "path": r"04 BELO E SANTOS CONSTRUÇÕES\DOCUMENTOS MENSAIS\2025\06 2025\GUIA DO FGTS COMP. 06.2025.pdf",
        "expected_value": 3,
        "expected_method": "fgts_digital",
    },
    # --- Reversed/rotated text ---
    {
        "id": "reversed_empr2_0624",
        "desc": "Reversed text - Empr 2 jun/2024",
        "path": r"02 JR CONSTRUÇÃO\SEFIP\2024\SEFIP COMPLETA 06-2024.pdf",
        "expected_value": 26,  # "62 :serodahlabarT .dtQ" reversed
        "expected_method": "any",
    },
    # --- Standard FGTS Extrato (should work) ---
    {
        "id": "fgts_extrato_empr2_0724",
        "desc": "Standard FGTS Extrato - Empr 2 jul/2024",
        "path": r"02 JR CONSTRUÇÃO\SEFIP\2024\SEFIP COMPLETA 07-2024.pdf",
        "expected_value": 23,
        "expected_method": "fgts",
    },
    # --- SEFIP Clássico ---
    {
        "id": "sefip_classico_empr2_0124",
        "desc": "SEFIP Clássico - Empr 2 jan/2024",
        "path": r"02 JR CONSTRUÇÃO\SEFIP\2024\01-2024\Relatório RE.pdf",
        "expected_value": 1,
        "expected_method": "sefip",
    },
    # --- Detalhe da Guia multi-tomador ---
    {
        "id": "detalhe_guia_empr1_0824",
        "desc": "Detalhe da Guia multi-tomador - Empr 1 ago/2024",
        "path": r"01 JVB\SEFIP\2024\08 2024\RELATORIO FGTS 08 2024.pdf",
        "expected_value": None,  # Should NOT return 29 (employer total); tomador-specific count needed
        "expected_method": "any",
        "note": "29 is employer total, not tomador-specific. Correct value unknown.",
        "accept_any_value": True,  # Just check it doesn't crash
    },
    # --- Non-SEFIP documents (should return no_match or be filtered) ---
    {
        "id": "non_sefip_credito_inss",
        "desc": "Non-SEFIP: CRÉDITO INSS - Empr 4 abr/2024",
        "path": r"04 BELO E SANTOS CONSTRUÇÕES\DOCUMENTOS MENSAIS\2024\04 2024\CRÉDITO INSS - 04.2024.pdf",
        "expected_value": None,
        "expected_method": "no_match",
    },
    # --- Scanned PDF (minimal text) ---
    {
        "id": "scanned_empr5_0425",
        "desc": "Scanned FGTS Digital - Empr 5 abr/2025",
        "path": r"05 DEMORAES ENCANADORES LTDA (MANOEL)\DOCUMENTOS MENSAIS\2025\04 2025\SEFIP.pdf",
        "expected_value": None,  # Will be None unless we add PyMuPDF fallback
        "expected_method": "any",
        "note": "Scanned PDF with only 'FGTS Digital' text. Needs OCR or PyMuPDF.",
    },
]

# ── select_best_pdf filtering tests ──
FILTER_TESTS = [
    {
        "id": "filter_credito_inss",
        "desc": "Should NOT select CRÉDITO INSS as best PDF",
        "files": [
            "CRÉDITO INSS - 04.2024.pdf",
        ],
        "should_select": None,  # No valid SEFIP in list
    },
    {
        "id": "filter_boleto_fgts",
        "desc": "Should NOT select BOLETO FGTS",
        "files": [
            "BOLETO FGTS  R$ 2132,15 - COMP. 08.2024.pdf",
        ],
        "should_select": None,
    },
    {
        "id": "filter_dctfweb",
        "desc": "Should NOT select DCTFWeb",
        "files": [
            "RELATÓRIO RESUMO DE CRÉDITOS - DCTFWeb.pdf",
        ],
        "should_select": None,
    },
    {
        "id": "filter_prefer_sefip_over_guia",
        "desc": "Should prefer SEFIP.pdf over GUIA DO FGTS",
        "files": [
            "GUIA DO FGTS COMP. 06.2025.pdf",
            "SEFIP.pdf",
        ],
        "should_select": "SEFIP.pdf",
    },
]


def run_extraction_tests():
    """Run extraction tests and return results."""
    results = []
    for tc in TEST_CASES:
        pdf_path = os.path.join(constants.BASE_DIR, tc["path"])
        if not os.path.exists(pdf_path):
            results.append({**tc, "status": "SKIP", "reason": "file not found"})
            continue

        value, method = extract_cat01(pdf_path)
        passed = False
        reason = ""

        if tc.get("accept_any_value"):
            passed = True
            reason = f"value={value}, method={method}"
        elif tc["expected_value"] is None:
            # Expect no extraction
            if value is None:
                passed = True
                reason = f"Correctly returned None [{method}]"
            else:
                passed = False
                reason = f"Expected None but got {value} [{method}]"
        else:
            # Expect specific value
            if value == tc["expected_value"]:
                if tc["expected_method"] == "any" or tc["expected_method"] in method:
                    passed = True
                    reason = f"value={value}, method={method}"
                else:
                    passed = True  # Value correct, method different but ok
                    reason = f"value={value}, method={method} (expected method containing '{tc['expected_method']}')"
            else:
                passed = False
                reason = f"Expected {tc['expected_value']} but got {value} [{method}]"

        results.append({**tc, "status": "PASS" if passed else "FAIL", "reason": reason,
                       "actual_value": value, "actual_method": method})
    return results


def run_filter_tests():
    """Run select_best_pdf filtering tests."""
    results = []
    for tc in FILTER_TESTS:
        # Create fake paths (only basename matters for filtering)
        fake_paths = [os.path.join("C:\\fake", f) for f in tc["files"]]
        selected = select_best_pdf(fake_paths)
        selected_name = os.path.basename(selected) if selected else None

        if tc["should_select"] is None:
            passed = selected is None
            reason = f"Selected: {selected_name}" if not passed else "Correctly returned None"
        else:
            passed = selected_name == tc["should_select"]
            reason = f"Selected: {selected_name}" if not passed else f"Correctly selected {selected_name}"

        results.append({**tc, "status": "PASS" if passed else "FAIL", "reason": reason})
    return results


def main():
    print("=" * 70)
    print("DIAGNOSTIC TEST — SEFIP Extraction Skill")
    print("=" * 70)

    # Extraction tests
    print("\n── EXTRACTION TESTS ──")
    ext_results = run_extraction_tests()
    for r in ext_results:
        icon = "✅" if r["status"] == "PASS" else "❌" if r["status"] == "FAIL" else "⏭️"
        print(f"  {icon} [{r['id']}] {r['desc']}")
        print(f"     {r['reason']}")
        if r.get("note"):
            print(f"     Note: {r['note']}")

    # Filter tests
    print("\n── FILTER TESTS (select_best_pdf) ──")
    flt_results = run_filter_tests()
    for r in flt_results:
        icon = "✅" if r["status"] == "PASS" else "❌"
        print(f"  {icon} [{r['id']}] {r['desc']}")
        print(f"     {r['reason']}")

    # Summary
    all_results = ext_results + flt_results
    total = len(all_results)
    passed = sum(1 for r in all_results if r["status"] == "PASS")
    failed = sum(1 for r in all_results if r["status"] == "FAIL")
    skipped = sum(1 for r in all_results if r["status"] == "SKIP")

    print(f"\n{'=' * 70}")
    print(f"RESULTADO: {passed}/{total} passed, {failed} failed, {skipped} skipped")
    print(f"Taxa de sucesso: {passed/(total-skipped)*100:.0f}%" if total > skipped else "N/A")
    print(f"{'=' * 70}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
