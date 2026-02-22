#!/usr/bin/env python3
"""
Validate NFSe extraction results — genérico (qualquer obra).

Usage:
    python3 validate_extraction.py [--obra-dir PATH] [--json PATH]

Defaults (após init_obra):
    --json  <obra>/.nfse-state/nfse_extracted.json
"""
import json
import os
import sys
import argparse
from collections import defaultdict

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)
import constants_nfse


def main():
    parser = argparse.ArgumentParser(description='Validate NFSe extraction')
    parser.add_argument('--obra-dir',
                        help='Caminho para a pasta da obra')
    parser.add_argument('--json',
                        help='JSON file to validate (default: <obra>/.nfse-state/nfse_extracted.json)')
    args = parser.parse_args()

    # Inicializar obra
    obra_dir = args.obra_dir
    if not obra_dir:
        if constants_nfse.is_obra_dir(os.getcwd()):
            obra_dir = os.getcwd()
        else:
            obra_dir = constants_nfse.BASE_DIR
    constants_nfse.init_obra(obra_dir)

    if not args.json:
        args.json = os.path.join(constants_nfse.STATE_DIR, "nfse_extracted.json")

    with open(args.json, 'r', encoding='utf-8') as f:
        data = json.load(f)

    records = data["records"]
    stats = data["stats"]
    issues = []

    print("=" * 70)
    print(f"RELATÓRIO DE VALIDAÇÃO — NFSe {constants_nfse.OBRA_NOME}")
    print("=" * 70)
    print(f"\nTotal de registros: {len(records)}")
    print(f"Extração OK: {stats['extracted_ok']}")
    print(f"Com campos faltantes: {stats['missing_fields']}")

    # Field completeness
    fields = ['nf', 'razao_social', 'cnpj_prestador', 'cno',
              'competencia', 'valor_total', 'inss', 'iss']
    print(f"\n{'Campo':<20s} {'Preenchido':>12s} {'Taxa':>8s}")
    print("-" * 42)
    for field in fields:
        count = sum(1 for r in records if r.get(field))
        pct = 100 * count / len(records) if records else 0
        # CNO is not applicable for surveillance services, so a lower rate is expected
        is_cno_with_vigilancia = (field == 'cno' and any(
            r.get('tipo_servico') == 'vigilancia' for r in records
        ))
        threshold = 70 if is_cno_with_vigilancia else 90
        flag = " ⚠" if pct < threshold else ""
        if is_cno_with_vigilancia and pct < 90:
            construcao_recs = [r for r in records if r.get('tipo_servico', 'construcao') == 'construcao']
            cno_construcao = sum(1 for r in construcao_recs if r.get('cno'))
            pct_c = 100 * cno_construcao / len(construcao_recs) if construcao_recs else 0
            flag = f" (construção: {pct_c:.0f}%, vigilância: N/A)"
        print(f"{field:<20s} {count:>5d}/{len(records):<5d} {pct:>6.1f}%{flag}")

    # Image PDFs
    no_text = [r for r in records if 'sem texto' in (r.get('observacao') or '').lower()]
    print(f"\nPDFs imagem (sem texto): {len(no_text)}")
    if no_text:
        by_contractor = defaultdict(int)
        for r in no_text:
            by_contractor[r['empreiteiro_num']] += 1
        print("  Por empreiteiro:")
        for num in sorted(by_contractor.keys()):
            print(f"    {num}: {by_contractor[num]} PDFs")

    # Service type breakdown
    construcao = [r for r in records if r.get('tipo_servico', 'construcao') == 'construcao']
    vigilancia = [r for r in records if r.get('tipo_servico') == 'vigilancia']
    print(f"\nTipo de serviço:")
    print(f"  Construção civil: {len(construcao)} registros")
    print(f"  Vigilância/segurança: {len(vigilancia)} registros")

    # INSS validation — rules differ by service type
    # Construction: INSS should be ~11% (cessão de mão de obra) or ~3.5% (Simples)
    # Surveillance: INSS = 0 is normal when owner-operated (no employee payroll)
    print(f"\n{'='*70}")
    print("CONFERÊNCIA INSS")
    print(f"{'='*70}")
    inss_ok = inss_simples = inss_bad = inss_zero_construcao = inss_zero_vigilancia = 0
    for r in records:
        if r['valor_total'] and r['valor_total'] > 0:
            tipo = r.get('tipo_servico', 'construcao')
            if r['inss'] is None or r['inss'] == 0:
                if tipo == 'vigilancia':
                    inss_zero_vigilancia += 1  # Expected for owner-operated surveillance
                else:
                    inss_zero_construcao += 1
            else:
                pct = r['inss'] / r['valor_total'] * 100
                if abs(pct - 11) < 1:
                    inss_ok += 1
                elif 3 <= pct <= 4:
                    inss_simples += 1
                else:
                    inss_bad += 1
                    issues.append(
                        f"INSS anormal: {r['empreiteiro'][:25]} NF {r['nf']} — "
                        f"Valor {r['valor_total']:,.2f}, INSS {r['inss']:,.2f} ({pct:.1f}%)"
                    )
    print(f"  INSS ≈ 11%: {inss_ok} registros")
    print(f"  INSS ≈ 3.5% (Simples): {inss_simples} registros")
    print(f"  INSS anormal: {inss_bad} registros")
    if inss_zero_construcao > 0:
        print(f"  INSS zero (construção): {inss_zero_construcao} registros ⚠")
    if inss_zero_vigilancia > 0:
        print(f"  INSS zero (vigilância — esperado): {inss_zero_vigilancia} registros ✓")

    # ISS validation — surveillance uses flat 5%, construction uses Simples 2-5%
    print(f"\n{'='*70}")
    print("CONFERÊNCIA ISS")
    print(f"{'='*70}")
    iss_ok = iss_bad = iss_zero = 0
    for r in records:
        if r['valor_total'] and r['valor_total'] > 0:
            tipo = r.get('tipo_servico', 'construcao')
            if r['iss'] is None or r['iss'] == 0:
                iss_zero += 1
            else:
                pct = r['iss'] / r['valor_total'] * 100
                # Construction: Simples Nacional range (varies by revenue bracket)
                # can be as low as 1.18% in the initial bracket
                if tipo == 'construcao' and 1.0 <= pct <= 6:
                    iss_ok += 1
                # Surveillance: typically flat 5%
                elif tipo == 'vigilancia' and 4.5 <= pct <= 5.5:
                    iss_ok += 1
                else:
                    iss_bad += 1
                    issues.append(
                        f"ISS anormal ({tipo}): {r['empreiteiro'][:25]} NF {r['nf']} — "
                        f"Valor {r['valor_total']:,.2f}, ISS {r['iss']:,.2f} ({pct:.1f}%)"
                    )
    print(f"  ISS dentro da faixa: {iss_ok} registros")
    print(f"  ISS fora da faixa: {iss_bad} registros")
    print(f"  ISS zero/ausente: {iss_zero} registros")

    # Competência gaps
    print(f"\n{'='*70}")
    print("SEQUÊNCIA DE COMPETÊNCIAS")
    print(f"{'='*70}")
    by_contractor = defaultdict(list)
    for r in records:
        if r['competencia']:
            by_contractor[r['empreiteiro_num']].append(r['competencia'])

    for num in sorted(by_contractor.keys()):
        comps = sorted(set(by_contractor[num]), key=lambda c: (
            int(c.split('/')[1]) * 100 + int(c.split('/')[0])
        ))
        if len(comps) >= 2:
            first = comps[0]
            last = comps[-1]
            print(f"  {num}: {first} → {last} ({len(comps)} competências)")

    # PDFs with text but missing NF
    text_no_nf = [r for r in records
                  if not r['nf'] and 'sem texto' not in (r.get('observacao') or '').lower()]
    if text_no_nf:
        print(f"\n{'='*70}")
        print(f"PDFs COM TEXTO mas SEM Nº NF ({len(text_no_nf)})")
        print(f"{'='*70}")
        for r in text_no_nf:
            print(f"  {r['empreiteiro'][:30]} | {r['arquivo'][:50]}")
        issues.append(f"{len(text_no_nf)} PDFs com texto mas sem NF identificado")

    # Totals
    total_valor = sum(r['valor_total'] or 0 for r in records)
    total_inss = sum(r['inss'] or 0 for r in records)
    total_iss = sum(r['iss'] or 0 for r in records)
    print(f"\n{'='*70}")
    print("TOTAIS")
    print(f"{'='*70}")
    print(f"  Valor Total:  R$ {total_valor:>15,.2f}")
    print(f"  INSS Total:   R$ {total_inss:>15,.2f}")
    print(f"  ISS Total:    R$ {total_iss:>15,.2f}")
    if total_valor > 0:
        print(f"  INSS/Valor:   {100 * total_inss / total_valor:.1f}%")
        print(f"  ISS/Valor:    {100 * total_iss / total_valor:.1f}%")

    if issues:
        print(f"\n{'='*70}")
        print(f"PROBLEMAS ENCONTRADOS ({len(issues)})")
        print(f"{'='*70}")
        for issue in issues[:20]:
            print(f"  • {issue}")
        if len(issues) > 20:
            print(f"  ... e mais {len(issues) - 20} problemas")
    else:
        print(f"\n✓ Nenhum problema encontrado")

    print(f"\n{'='*70}")


if __name__ == "__main__":
    main()
