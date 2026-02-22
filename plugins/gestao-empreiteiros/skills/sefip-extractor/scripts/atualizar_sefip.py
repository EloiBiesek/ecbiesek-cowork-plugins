#!/usr/bin/env python3
"""
ATUALIZAR SEFIP
=====================================
Script unificado que executa todo o pipeline de extração SEFIP:
  1. Verifica estado atual
  2. Extrai texto dos PDFs
  3. OCR nos PDFs escaneados (se rapidocr disponível)
  4. Atualiza planilha
  5. Atualiza aba RESUMO NOVO
  6. Mostra resultado final

Uso:
    python atualizar_sefip.py --obra-dir "/caminho/para/prestadores"
    python atualizar_sefip.py                    # Atualização incremental (detecta obra automaticamente)
    python atualizar_sefip.py --force            # Reprocessa tudo do zero
    python atualizar_sefip.py --empreiteiros 3 7 # Só empreiteiros específicos
    python atualizar_sefip.py --sem-ocr          # Pula etapa de OCR
    python atualizar_sefip.py --listar-obras     # Lista obras disponíveis e sai
"""

import os, sys, json, time, argparse, subprocess
from datetime import datetime

# Garantir que imports dos scripts funcionem
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)

import constants

# ── Helpers ──────────────────────────────────────────────────────────────────

def tempo_formatado(segundos):
    if segundos < 60:
        return f"{segundos:.0f}s"
    return f"{int(segundos // 60)}min {int(segundos % 60)}s"


def verificar_dependencias():
    """Verifica se as dependências Python estão instaladas."""
    faltando = []
    for mod, pip_name in [("pdfplumber", "pdfplumber"),
                          ("openpyxl", "openpyxl"),
                          ("fitz", "pymupdf"),
                          ("rapidocr_onnxruntime", "rapidocr-onnxruntime")]:
        try:
            __import__(mod)
        except ImportError:
            faltando.append(pip_name)

    if faltando:
        print(f"\n  DEPENDÊNCIAS FALTANDO: {', '.join(faltando)}")
        print(f"  Execute: pip install {' '.join(faltando)}")
        print()
        return False
    return True


def verificar_ocr():
    """Verifica se RapidOCR está disponível para OCR."""
    try:
        import rapidocr_onnxruntime
        return True
    except ImportError:
        return False


def limpar_state(force, empreiteiros):
    """Limpa state files quando --force é usado."""
    os.makedirs(constants.STATE_DIR, exist_ok=True)
    files_to_reset = {
        "extractions_text.json": {},
        "extractions_ocr.json": {},
        "needs_ocr.json": [],
        "changes_log.json": [],
        "divergences.json": [],
    }
    for fname, empty_val in files_to_reset.items():
        fpath = os.path.join(constants.STATE_DIR, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(empty_val, f)


def listar_obras(search_dir):
    """Lista obras encontradas a partir de um diretório."""
    obras = constants.descobrir_obras(search_dir)
    if not obras:
        print(f"\n  Nenhuma obra encontrada em: {search_dir}")
        print("  (Obras são pastas com subpastas numeradas: '01 NOME', '02 NOME', ...)")
        return []

    print(f"\n  Obras encontradas ({len(obras)}):\n")
    for i, obra in enumerate(obras):
        print(f"    [{i+1}] {obra['name']}")
        print(f"        {obra['path']}")
        print(f"        {obra['empreiteiros']} empreiteiros detectados")
        print()
    return obras


def selecionar_obra(search_dir):
    """Detecta e permite seleção interativa de uma obra.

    Retorna o caminho da obra selecionada ou None se nenhuma encontrada.
    """
    obras = constants.descobrir_obras(search_dir)

    if not obras:
        print(f"\n  ERRO: Nenhuma obra encontrada em '{search_dir}'")
        print("  Pastas de obra devem conter subpastas numeradas (ex: '01 NOME', '02 NOME')")
        print("  Use --obra-dir para especificar o caminho diretamente.")
        return None

    if len(obras) == 1:
        obra = obras[0]
        print(f"\n  Obra detectada: {obra['name']}")
        print(f"  Caminho: {obra['path']}")
        print(f"  Empreiteiros: {obra['empreiteiros']}")
        resp = input("\n  Processar esta obra? [S/n]: ").strip().lower()
        if resp and resp != "s":
            print("  Cancelado.")
            return None
        return obra["path"]

    # Múltiplas obras: forçar seleção
    print(f"\n  {len(obras)} obra(s) encontrada(s). Selecione UMA:\n")
    for i, obra in enumerate(obras):
        print(f"    [{i+1}] {obra['name']} ({obra['empreiteiros']} empreiteiros)")
        print(f"        {obra['path']}")
    print()

    while True:
        resp = input(f"  Escolha [1-{len(obras)}]: ").strip()
        try:
            idx = int(resp) - 1
            if 0 <= idx < len(obras):
                return obras[idx]["path"]
        except ValueError:
            pass
        print(f"  Opção inválida. Digite um número de 1 a {len(obras)}.")


# ── Pipeline steps ───────────────────────────────────────────────────────────

def step_check_status():
    """Verifica se há trabalho pendente. Retorna True se precisa processar."""
    from check_status import main as check_main
    check_main()
    return True  # Sempre continua — o script decide baseado no --force


def step_extract_text(empreiteiros=None, force=False):
    """Executa extração de texto."""
    argv_backup = sys.argv
    sys.argv = ["extract_text.py"]
    if empreiteiros:
        sys.argv += ["--empreiteiros"] + [str(e) for e in empreiteiros]
    if force:
        sys.argv.append("--force")

    try:
        from extract_text import main as text_main
        text_main()
    finally:
        sys.argv = argv_backup


def step_extract_ocr(empreiteiros=None):
    """Executa OCR nos PDFs escaneados."""
    argv_backup = sys.argv
    sys.argv = ["extract_ocr.py", "--from-pending"]

    try:
        from extract_ocr import main as ocr_main
        ocr_main()
    finally:
        sys.argv = argv_backup


def step_update_planilha():
    """Atualiza a planilha com os dados extraídos."""
    argv_backup = sys.argv
    sys.argv = ["update_planilha.py"]

    try:
        from update_planilha import main as planilha_main
        planilha_main()
    finally:
        sys.argv = argv_backup


def step_write_resumo(run_stats=None):
    """Atualiza a aba RESUMO NOVO."""
    from write_resumo import write_resumo
    write_resumo(run_stats)


def _load_json_safe(path, expected_type=dict):
    """Carrega JSON com validação de tipo."""
    if not os.path.exists(path):
        return expected_type()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, expected_type):
            return data
        return expected_type()
    except Exception:
        return expected_type()


def contar_resultados():
    """Conta resultados dos state files para o log."""
    stats = {
        "novos_texto": 0, "novos_ocr": 0, "pulados": 0,
        "erros": 0, "divergencias_detectadas": 0,
        "divergencias_resolvidas": 0, "comp_mais_recente": None,
        "empreiteiros_processados": 0, "observacoes": ""
    }

    text_data = _load_json_safe(os.path.join(constants.STATE_DIR, "extractions_text.json"), dict)
    ocr_data = _load_json_safe(os.path.join(constants.STATE_DIR, "extractions_ocr.json"), dict)
    divs = _load_json_safe(os.path.join(constants.STATE_DIR, "divergences.json"), list)

    # Contar extrações válidas
    all_ok_months = []
    emprs_with_data = set()
    for empr, months in text_data.items():
        if not isinstance(months, dict):
            continue
        for m, v in months.items():
            if not isinstance(v, dict):
                continue
            if v.get("value") is not None:
                stats["novos_texto"] += 1
                all_ok_months.append(m)
                emprs_with_data.add(empr)

    for empr, months in ocr_data.items():
        if not isinstance(months, dict):
            continue
        for m, v in months.items():
            if not isinstance(v, dict):
                continue
            if v.get("value") is not None:
                stats["novos_ocr"] += 1
                if m not in all_ok_months:
                    all_ok_months.append(m)
                emprs_with_data.add(empr)

    stats["empreiteiros_processados"] = len(emprs_with_data)
    stats["divergencias_detectadas"] = len([d for d in divs if isinstance(d, dict) and not d.get("resolved")])
    stats["divergencias_resolvidas"] = len([d for d in divs if isinstance(d, dict) and d.get("resolved")])

    if all_ok_months:
        stats["comp_mais_recente"] = max(all_ok_months)

    return stats


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Atualiza SEFIP",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python atualizar_sefip.py --obra-dir "C:/caminho/prestadores"  Especifica a obra
  python atualizar_sefip.py                     Atualização incremental (detecta obra)
  python atualizar_sefip.py --force             Reprocessa tudo do zero
  python atualizar_sefip.py --empreiteiros 3 7  Só empreiteiros 03 e 07
  python atualizar_sefip.py --sem-ocr           Pula OCR (mais rápido)
  python atualizar_sefip.py --listar-obras      Lista obras disponíveis
"""
    )
    parser.add_argument('--obra-dir', type=str, default=None,
                        help='Caminho para a pasta de prestadores da obra')
    parser.add_argument('--listar-obras', action='store_true',
                        help='Lista obras disponíveis e sai')
    parser.add_argument('--force', action='store_true',
                        help='Reprocessa tudo do zero')
    parser.add_argument('--empreiteiros', nargs='+', type=int, default=None,
                        help='Empreiteiros específicos (ex: 3 7 10)')
    parser.add_argument('--sem-ocr', action='store_true',
                        help='Pula etapa de OCR')
    args = parser.parse_args()

    # ── Determinar pasta da obra ──
    if args.listar_obras:
        listar_obras(os.getcwd())
        sys.exit(0)

    if args.obra_dir:
        obra_dir = os.path.normpath(args.obra_dir)
        if not os.path.isdir(obra_dir):
            print(f"\n  ERRO: Pasta não encontrada: {obra_dir}")
            sys.exit(1)
        if not constants.is_obra_dir(obra_dir):
            print(f"\n  AVISO: '{obra_dir}' não parece ser uma pasta de obra")
            print("  (Esperado: subpastas numeradas '01 NOME', '02 NOME', ...)")
            resp = input("  Continuar mesmo assim? [s/N]: ").strip().lower()
            if resp != "s":
                sys.exit(0)
    else:
        # Auto-detectar: verificar se estamos dentro de uma obra
        cwd = os.getcwd()
        if constants.is_obra_dir(cwd):
            obra_dir = cwd
        else:
            # Tentar detectar obras no CWD
            obra_dir = selecionar_obra(cwd)
            if obra_dir is None:
                sys.exit(1)

    # ── Inicializar constants para a obra selecionada ──
    constants.init_obra(obra_dir)

    inicio = time.time()

    config_path = os.path.join(constants.STATE_DIR, "obra.json")
    if not os.path.exists(config_path):
        config_path_skill = os.path.join(constants.SKILL_DIR, "obra.json")
        if not os.path.exists(config_path_skill):
            print()
            print("  AVISO: obra.json nao encontrado — usando configuracao padrao.")
            print("  Para configurar: python scripts/configurar_obra.py")
        else:
            # Verificar se obra.json corresponde às pastas reais
            pastas_config = set(constants.EMPR_FOLDERS.values())
            pastas_reais = set()
            import re as _re
            for entry in os.listdir(obra_dir):
                if os.path.isdir(os.path.join(obra_dir, entry)) and _re.match(r'^\d{1,2}\s', entry):
                    pastas_reais.add(entry)
            if pastas_config and pastas_reais and not pastas_config.intersection(pastas_reais):
                print()
                print("  ERRO: obra.json nao corresponde a esta obra!")
                print(f"  Config: {list(pastas_config)[:3]}...")
                print(f"  Pastas: {list(pastas_reais)[:3]}...")
                print("  Parece que o obra.json foi copiado de outra obra.")
                print("  Delete obra.json e rode novamente, ou:")
                print("  python scripts/configurar_obra.py")
                sys.exit(1)

    print()
    print("=" * 70)
    print(f"  ATUALIZACAO SEFIP — {constants.OBRA_NOME}")
    print(f"  Pasta: {obra_dir}")
    print("  " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    modo = "COMPLETO (--force)" if args.force else "INCREMENTAL"
    print(f"  Modo: {modo}")
    print("=" * 70)

    # ── Verificar dependências ──
    if not verificar_dependencias():
        sys.exit(1)

    tem_ocr = verificar_ocr()
    if not tem_ocr:
        print("\n  AVISO: rapidocr-onnxruntime não encontrado — OCR será desabilitado.")
        print("  PDFs escaneados não serão processados.")
        print("  Para instalar: pip install rapidocr-onnxruntime")
        args.sem_ocr = True

    # ── Limpar state se --force ──
    if args.force:
        print("\n  Limpando estado anterior...")
        limpar_state(args.force, args.empreiteiros)

    # ── Passo 1: Verificar estado ──
    print("\n")
    print("-" * 70)
    print("  PASSO 1/5 — Verificando estado atual")
    print("-" * 70)
    step_check_status()

    # ── Passo 2: Extrair texto ──
    print("\n")
    print("-" * 70)
    print("  PASSO 2/5 — Extraindo texto dos PDFs")
    print("-" * 70)
    t2 = time.time()
    step_extract_text(empreiteiros=args.empreiteiros, force=args.force)
    print(f"\n  Tempo: {tempo_formatado(time.time() - t2)}")

    # ── Passo 3: OCR ──
    if not args.sem_ocr:
        print("\n")
        print("-" * 70)
        print("  PASSO 3/5 — OCR nos PDFs escaneados")
        print("-" * 70)
        t3 = time.time()
        step_extract_ocr(empreiteiros=args.empreiteiros)
        print(f"\n  Tempo: {tempo_formatado(time.time() - t3)}")
    else:
        print("\n")
        print("-" * 70)
        print("  PASSO 3/5 — OCR: PULADO (--sem-ocr ou rapidocr indisponível)")
        print("-" * 70)

    # ── Passo 4: Atualizar planilha ──
    print("\n")
    print("-" * 70)
    print("  PASSO 4/5 — Atualizando planilha")
    print("-" * 70)
    t4 = time.time()
    step_update_planilha()
    print(f"\n  Tempo: {tempo_formatado(time.time() - t4)}")

    # ── Passo 5: RESUMO NOVO ──
    print("\n")
    print("-" * 70)
    print("  PASSO 5/5 — Atualizando aba RESUMO NOVO")
    print("-" * 70)
    run_stats = contar_resultados()
    run_stats["observacoes"] = f"Pipeline unificado {'(--force)' if args.force else '(incremental)'}"
    step_write_resumo(run_stats)

    # ── Resultado final ──
    total = time.time() - inicio
    print("\n")
    print("=" * 70)
    print(f"  CONCLUÍDO em {tempo_formatado(total)}")
    print("=" * 70)

    # Resumo final rápido
    stats = run_stats
    print(f"  Texto: {stats['novos_texto']} meses extraídos")
    print(f"  OCR:   {stats['novos_ocr']} meses extraídos")
    if stats['divergencias_detectadas']:
        print(f"  DIVERGÊNCIAS: {stats['divergencias_detectadas']} pendente(s)")
        print(f"  Execute: python scripts/resolve_divergences.py --list")
    else:
        print(f"  Divergências: 0")
    print()


if __name__ == "__main__":
    main()
