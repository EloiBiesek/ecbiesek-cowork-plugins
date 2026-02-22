#!/usr/bin/env python3
"""
Mescla resultados dos agentes paralelos nos arquivos de estado,
atualiza a planilha (Alocação de colaboradores) e registra log no RESUMO NOVO.

Uso:
    python3 merge_agent_results.py [--dry-run]
"""

import os, sys, json, argparse

MERGE_DIR   = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, MERGE_DIR)

import constants
from update_planilha import main as update_planilha_main
from write_resumo import write_resumo


def load_json(path, default=None):
    if os.path.exists(path):
        try:
            return json.load(open(path, encoding="utf-8"))
        except Exception:
            pass
    return default if default is not None else {}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    print("=" * 60)
    print("MERGE — Resultados dos agentes paralelos")
    print("=" * 60)

    # ── 1. Coletar arquivos de agentes ───────────────────────────────────────
    agent_files = sorted([
        f for f in os.listdir(constants.STATE_DIR)
        if f.startswith("agent_results_empr_") and f.endswith(".json")
    ])

    if not agent_files:
        print("Nenhum arquivo de agente encontrado em", constants.STATE_DIR)
        sys.exit(1)

    print(f"\n{len(agent_files)} arquivos de agente encontrados:")
    for f in agent_files:
        print(f"  {f}")

    # ── 2. Carregar estado existente ─────────────────────────────────────────
    text_path = os.path.join(constants.STATE_DIR, "extractions_text.json")
    ocr_path  = os.path.join(constants.STATE_DIR, "extractions_ocr.json")
    text_data = load_json(text_path)
    ocr_data  = load_json(ocr_path)

    # ── 3. Mesclar resultados ────────────────────────────────────────────────
    run_totals = {
        "novos_texto":               0,
        "novos_ocr":                 0,
        "pulados":                   0,
        "erros":                     0,
        "empreiteiros_processados":  len(agent_files),
        "comp_mais_recente":         None,
        "divergencias_detectadas":   0,
        "divergencias_resolvidas":   0,
        "observacoes":               [],
    }

    for fname in agent_files:
        fpath = os.path.join(constants.STATE_DIR, fname)
        agent = json.load(open(fpath, encoding="utf-8"))
        empr  = str(agent["empr"])
        name  = agent.get("name", f"Empr {empr}")

        text_r = agent.get("text_results", {})
        ocr_r  = agent.get("ocr_results",  {})
        stats  = agent.get("stats", {})

        # Substituir entrada do empreiteiro nos state files
        text_data[empr] = text_r
        if ocr_r:
            ocr_data[empr] = ocr_r

        # Acumular totais
        run_totals["novos_texto"] += stats.get("novos_texto", 0)
        run_totals["novos_ocr"]   += stats.get("novos_ocr",   0)
        run_totals["pulados"]     += stats.get("pulados",      0)
        run_totals["erros"]       += stats.get("erros",        0)

        comp_fim = stats.get("comp_final")
        if comp_fim:
            if (run_totals["comp_mais_recente"] is None
                    or comp_fim > run_totals["comp_mais_recente"]):
                run_totals["comp_mais_recente"] = comp_fim

        ok_t = sum(1 for v in text_r.values() if v.get("value") is not None)
        ok_o = sum(1 for v in ocr_r.values()  if v.get("value") is not None)
        print(f"\n  Empr {empr:>2} {name[:38]:<38} | texto={ok_t:2d} | OCR={ok_o:2d} "
              f"| novos={stats.get('novos_texto',0)+stats.get('novos_ocr',0)} "
              f"| pulados={stats.get('pulados',0)}")

    total_novos = run_totals["novos_texto"] + run_totals["novos_ocr"]
    print(f"\n  TOTAIS -> novos={total_novos} | pulados={run_totals['pulados']} "
          f"| erros={run_totals['erros']}")

    if args.dry_run:
        print("\n[DRY RUN] Nada foi salvo.")
        return

    # ── 4. Salvar state files ─────────────────────────────────────────────────
    with open(text_path, "w", encoding="utf-8") as f:
        json.dump(text_data, f, indent=2, ensure_ascii=False)
    n_text = sum(len(v) for v in text_data.values())
    print(f"\n  extractions_text.json salvo ({n_text} registros)")

    with open(ocr_path, "w", encoding="utf-8") as f:
        json.dump(ocr_data, f, indent=2, ensure_ascii=False)
    n_ocr = sum(len(v) for v in ocr_data.values())
    print(f"  extractions_ocr.json salvo ({n_ocr} registros)")

    # ── 5. Atualizar aba "Alocação de colaboradores" (import direto) ────────
    print("\nAtualizando planilha (Alocação de colaboradores)...")
    # Simula sys.argv para o argparse do update_planilha
    old_argv = sys.argv
    sys.argv = ["update_planilha.py"]
    try:
        update_planilha_main()
    except SystemExit:
        pass
    finally:
        sys.argv = old_argv

    # Contar divergências diretamente do state file
    div_path = os.path.join(constants.STATE_DIR, "divergences.json")
    if os.path.exists(div_path):
        div_data = load_json(div_path, [])
        if isinstance(div_data, list):
            run_totals["divergencias_detectadas"] = sum(
                1 for d in div_data if not d.get("resolved"))
            run_totals["divergencias_resolvidas"] = sum(
                1 for d in div_data if d.get("resolved"))

    # ── 6. Escrever RESUMO NOVO (import direto) ─────────────────────────────
    print("\nAtualizando aba RESUMO NOVO...")
    if run_totals["observacoes"]:
        run_totals["observacoes"] = " | ".join(run_totals["observacoes"])
    else:
        run_totals["observacoes"] = (
            f"{len(agent_files)} empreiteiros processados. "
            f"Novos: {total_novos}. Pulados: {run_totals['pulados']}."
        )

    write_resumo(run_totals)

    # ── 7. Limpar temporários ────────────────────────────────────────────────
    print("Limpando arquivos temporários dos agentes...")
    for fname in agent_files:
        fpath = os.path.join(constants.STATE_DIR, fname)
        try:
            os.remove(fpath)
            print(f"  Removido: {fname}")
        except Exception as e:
            print(f"  Não removido ({fname}): {e}")

    print("\n[MERGE CONCLUIDO]")


if __name__ == "__main__":
    main()
