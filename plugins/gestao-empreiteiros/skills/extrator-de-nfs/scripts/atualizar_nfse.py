#!/usr/bin/env python3
"""
ATUALIZAR NFSe
=====================================
Script unificado que executa todo o pipeline de extração NFSe:
  1. Verifica estado atual
  2. Extrai dados dos PDFs de nota fiscal
  3. OCR nos PDFs escaneados (se rapidocr disponível)
  4. Popula planilha de controle de NFs
  5. Valida resultados
  6. Mostra resultado final

Uso:
    python atualizar_nfse.py --obra-dir "/caminho/para/prestadores"
    python atualizar_nfse.py                    # Detecta obra automaticamente
    python atualizar_nfse.py --force            # Reprocessa tudo do zero
    python atualizar_nfse.py --sem-ocr          # Pula etapa de OCR
    python atualizar_nfse.py --listar-obras     # Lista obras disponíveis e sai
"""

import os, sys, json, time, argparse
from datetime import datetime

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)

import constants_nfse


# ── Helpers ──────────────────────────────────────────────────────────────────

def tempo_formatado(segundos):
    if segundos < 60:
        return f"{segundos:.0f}s"
    return f"{int(segundos // 60)}min {int(segundos % 60)}s"


def verificar_dependencias():
    """Verifica se as dependências Python estão instaladas."""
    faltando = []
    for mod, pip_name in [("pdfplumber", "pdfplumber"),
                          ("openpyxl", "openpyxl")]:
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
    """Verifica se RapidOCR + PyMuPDF estão disponíveis para OCR."""
    try:
        import fitz
        import rapidocr_onnxruntime
        return True
    except ImportError:
        return False


def listar_obras(search_dir):
    """Lista obras encontradas a partir de um diretório."""
    obras = constants_nfse.descobrir_obras(search_dir)
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
    """Detecta e permite seleção interativa de uma obra."""
    obras = constants_nfse.descobrir_obras(search_dir)

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
    """Verifica estado atual da extração."""
    argv_backup = sys.argv
    sys.argv = ["check_status.py", "--obra-dir", constants_nfse.BASE_DIR]
    try:
        from check_status import main as check_main
        check_main()
    except SystemExit:
        pass
    finally:
        sys.argv = argv_backup


def step_extract(incremental=False, force=False):
    """Executa extração de texto dos PDFs de NF."""
    argv_backup = sys.argv
    sys.argv = ["extract_all_nfse.py", "--obra-dir", constants_nfse.BASE_DIR]
    if incremental:
        sys.argv.append("--incremental")
    try:
        from extract_all_nfse import main as extract_main
        extract_main()
    except SystemExit:
        pass
    finally:
        sys.argv = argv_backup


def step_ocr(batch_size=15):
    """Executa OCR nos PDFs-imagem."""
    argv_backup = sys.argv
    sys.argv = ["ocr_nfse.py", "--obra-dir", constants_nfse.BASE_DIR,
                "--batch-size", str(batch_size)]
    try:
        from ocr_nfse import main as ocr_main
        ocr_main()
    except SystemExit:
        pass
    finally:
        sys.argv = argv_backup


def step_populate_xlsx(append=False):
    """Popula planilha de controle de NFs."""
    argv_backup = sys.argv
    sys.argv = ["populate_xlsx.py", "--obra-dir", constants_nfse.BASE_DIR]
    if append:
        sys.argv.append("--append")
    try:
        from populate_xlsx import main as populate_main
        populate_main()
    except SystemExit:
        pass
    finally:
        sys.argv = argv_backup


def step_validate():
    """Valida resultados da extração."""
    argv_backup = sys.argv
    sys.argv = ["validate_extraction.py", "--obra-dir", constants_nfse.BASE_DIR]
    try:
        from validate_extraction import main as validate_main
        validate_main()
    except SystemExit:
        pass
    finally:
        sys.argv = argv_backup


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Atualiza NFSe — pipeline unificado",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python atualizar_nfse.py --obra-dir "C:/caminho/prestadores"  Especifica a obra
  python atualizar_nfse.py                     Detecta obra automaticamente
  python atualizar_nfse.py --force             Reprocessa tudo do zero
  python atualizar_nfse.py --sem-ocr           Pula OCR (mais rápido)
  python atualizar_nfse.py --listar-obras      Lista obras disponíveis
"""
    )
    parser.add_argument('--obra-dir', type=str, default=None,
                        help='Caminho para a pasta de prestadores da obra')
    parser.add_argument('--listar-obras', action='store_true',
                        help='Lista obras disponíveis e sai')
    parser.add_argument('--force', action='store_true',
                        help='Reprocessa tudo do zero')
    parser.add_argument('--sem-ocr', action='store_true',
                        help='Pula etapa de OCR')
    parser.add_argument('--incremental', action='store_true',
                        help='Pula PDFs já registrados na planilha')
    args = parser.parse_args()

    # ── Listar obras ──
    if args.listar_obras:
        listar_obras(os.getcwd())
        sys.exit(0)

    # ── Determinar pasta da obra ──
    if args.obra_dir:
        obra_dir = os.path.normpath(args.obra_dir)
        if not os.path.isdir(obra_dir):
            print(f"\n  ERRO: Pasta não encontrada: {obra_dir}")
            sys.exit(1)
        if not constants_nfse.is_obra_dir(obra_dir):
            print(f"\n  AVISO: '{obra_dir}' não parece ser uma pasta de obra.")
            print("  (Não foram encontradas subpastas numeradas de empreiteiros)")
            resp = input("  Continuar mesmo assim? [s/N]: ").strip().lower()
            if resp != "s":
                sys.exit(1)
    elif constants_nfse.is_obra_dir(os.getcwd()):
        obra_dir = os.getcwd()
    else:
        obra_dir = selecionar_obra(os.getcwd())
        if not obra_dir:
            sys.exit(1)

    # ── Inicializar ──
    constants_nfse.init_obra(obra_dir)

    print()
    print("=" * 70)
    print(f"  ATUALIZAR NFSe — {constants_nfse.OBRA_NOME}")
    print("=" * 70)
    print(f"  Obra:      {constants_nfse.BASE_DIR}")
    print(f"  Planilha:  {constants_nfse.XLSX_PATH}")
    print(f"  State:     {constants_nfse.STATE_DIR}")
    print(f"  Início:    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    # ── Verificar dependências ──
    if not verificar_dependencias():
        sys.exit(1)

    has_ocr = verificar_ocr()
    if not has_ocr:
        print("  OCR indisponível (instale: pip install rapidocr-onnxruntime pymupdf)")
        if not args.sem_ocr:
            print("  OCR será pulado automaticamente.")
            args.sem_ocr = True

    t_start = time.time()

    # ── Passo 1: Verificar estado ──
    print(f"\n{'─'*70}")
    print("  PASSO 1/5: Verificar estado atual")
    print(f"{'─'*70}")
    step_check_status()

    # ── Passo 2: Extrair PDFs ──
    print(f"\n{'─'*70}")
    print("  PASSO 2/5: Extrair dados dos PDFs de NF")
    print(f"{'─'*70}")
    t2 = time.time()
    step_extract(incremental=args.incremental, force=args.force)
    print(f"  Tempo: {tempo_formatado(time.time() - t2)}")

    # ── Passo 3: OCR ──
    if not args.sem_ocr and has_ocr:
        print(f"\n{'─'*70}")
        print("  PASSO 3/5: OCR nos PDFs-imagem")
        print(f"{'─'*70}")
        t3 = time.time()
        step_ocr()
        print(f"  Tempo: {tempo_formatado(time.time() - t3)}")
    else:
        print(f"\n{'─'*70}")
        print("  PASSO 3/5: OCR — PULADO")
        print(f"{'─'*70}")

    # ── Passo 4: Planilha ──
    print(f"\n{'─'*70}")
    print("  PASSO 4/5: Atualizar planilha")
    print(f"{'─'*70}")
    step_populate_xlsx()

    # ── Passo 5: Validação ──
    print(f"\n{'─'*70}")
    print("  PASSO 5/5: Validar resultados")
    print(f"{'─'*70}")
    step_validate()

    # ── Resumo final ──
    t_total = time.time() - t_start
    print(f"\n{'='*70}")
    print(f"  PIPELINE CONCLUÍDO — {tempo_formatado(t_total)}")
    print(f"  Planilha: {constants_nfse.XLSX_PATH}")
    print(f"  State:    {constants_nfse.STATE_DIR}")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
