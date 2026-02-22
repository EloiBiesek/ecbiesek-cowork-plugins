"""
Configurar obra — wizard de setup para SEFIP Extractor.

Auto-detecta empreiteiros, subpastas SEFIP e range de meses.
Pede apenas CNO e nome da obra.
Gera obra.json em .sefip-state/ dentro da pasta da obra.

Uso:
  python configurar_obra.py --obra-dir "/caminho/para/prestadores"
  python configurar_obra.py --auto --cno X --nome Y --obra-dir /path
"""

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import constants

# Subpastas conhecidas que contêm documentos SEFIP
_KNOWN_SEFIP_SUBS = [
    "SEFIP", "DOCUMENTOS MENSAIS", "DOCUMENTAÇÕES MENSAIS",
    "ENTREGA DE DOCUMENTOS MENSAIS", "DOMENTAÇÃO MENSAL",
]


def detect_empreiteiros(base_dir):
    """Detecta pastas de empreiteiros numeradas (ex: '01 JVB', '02 JR CONSTRUÇÃO')."""
    empreiteiros = []
    try:
        entries = sorted(os.listdir(base_dir))
    except OSError:
        return empreiteiros

    for entry in entries:
        full = os.path.join(base_dir, entry)
        if not os.path.isdir(full):
            continue
        # Ignorar pastas especiais
        if entry.startswith(".") or entry.startswith("_"):
            continue
        if entry.upper() in ("PLANILHA DE INFORMAÇÕES DAS NOTAS FISCAIS",):
            continue
        # Detectar padrão "NN NOME" ou "NN_NOME"
        m = re.match(r'^(\d{1,2})\s*(.*)', entry)
        if m:
            num = int(m.group(1))
            nome = m.group(2).strip() or entry
            empreiteiros.append({
                "num": num,
                "pasta": entry,
                "nome_curto": nome[:30],
                "nome_completo": entry,
            })
    return empreiteiros


def detect_sefip_subfolders(base_dir, empreiteiros):
    """Detecta quais subpastas SEFIP existem nos empreiteiros."""
    found = set()
    for e in empreiteiros:
        empr_path = os.path.join(base_dir, e["pasta"])
        if not os.path.isdir(empr_path):
            continue
        for sub in _KNOWN_SEFIP_SUBS:
            if os.path.isdir(os.path.join(empr_path, sub)):
                found.add(sub)
    return sorted(found) if found else ["SEFIP", "DOCUMENTOS MENSAIS"]


def detect_month_range(base_dir, empreiteiros, sefip_subs):
    """Detecta o range de meses a partir das subpastas de ano/mês."""
    months_found = set()
    _month_pattern = re.compile(r'(\d{2})\s+(\d{4})')

    for e in empreiteiros:
        empr_path = os.path.join(base_dir, e["pasta"])
        for sub in sefip_subs:
            sub_path = os.path.join(empr_path, sub)
            if not os.path.isdir(sub_path):
                continue
            for year_entry in os.listdir(sub_path):
                year_path = os.path.join(sub_path, year_entry)
                if not os.path.isdir(year_path):
                    continue
                for month_entry in os.listdir(year_path):
                    m = _month_pattern.search(month_entry)
                    if m:
                        mm, yy = int(m.group(1)), int(m.group(2))
                        if 1 <= mm <= 12 and 2020 <= yy <= 2030:
                            months_found.add((yy, mm))
            for month_entry in os.listdir(sub_path):
                m = _month_pattern.search(month_entry)
                if m:
                    mm, yy = int(m.group(1)), int(m.group(2))
                    if 1 <= mm <= 12 and 2020 <= yy <= 2030:
                        months_found.add((yy, mm))

    from datetime import date
    today = date.today()
    now = (today.year, today.month)

    if not months_found:
        return (today.year - 2, today.month), now

    s = min(months_found)
    e = max(max(months_found), now)
    return s, e


def detect_planilha(base_dir):
    """Detecta planilha existente na pasta da obra."""
    try:
        for f in os.listdir(base_dir):
            if f.lower().endswith(".xlsx") and "controle" in f.lower() and "empreiteiro" in f.lower():
                return f
    except OSError:
        pass
    return "Controle de Alocação e ISS mensal de empreiteiros.xlsx"


def main():
    parser = argparse.ArgumentParser(description="Configurar SEFIP Extractor para uma obra")
    parser.add_argument("--obra-dir", type=str, default=None,
                        help="Caminho para a pasta de prestadores da obra")
    parser.add_argument("--auto", action="store_true", help="Modo não-interativo")
    parser.add_argument("--cno", type=str, help="CNO da obra (12 dígitos)")
    parser.add_argument("--nome", type=str, help="Nome da obra")
    args = parser.parse_args()

    # Determinar pasta da obra
    if args.obra_dir:
        base_dir = os.path.normpath(args.obra_dir)
    else:
        # Fallback: se estamos dentro de uma obra, usar CWD; senão usar o padrão relativo
        cwd = os.getcwd()
        if constants.is_obra_dir(cwd):
            base_dir = cwd
        else:
            base_dir = constants.BASE_DIR

    if not os.path.isdir(base_dir):
        print(f"ERRO: Pasta não encontrada: {base_dir}")
        sys.exit(1)

    # Config será salva em .sefip-state/ da obra
    state_dir = os.path.join(base_dir, ".sefip-state")
    os.makedirs(state_dir, exist_ok=True)
    config_path = os.path.join(state_dir, "obra.json")

    print()
    print("=" * 60)
    print("  CONFIGURAR SEFIP EXTRACTOR")
    print("=" * 60)
    print()
    print(f"  Pasta da obra: {base_dir}")
    print()

    # 1. Detectar empreiteiros
    empreiteiros = detect_empreiteiros(base_dir)
    if not empreiteiros:
        print("ERRO: Nenhuma pasta de empreiteiro detectada (formato '01 NOME').")
        print(f"  Verifique a pasta: {base_dir}")
        sys.exit(1)

    print(f"Detectei {len(empreiteiros)} empreiteiro(s):")
    for e in empreiteiros:
        print(f"  {e['num']:02d} {e['nome_curto']}")
    print()

    if not args.auto:
        resp = input("Correto? [S/n]: ").strip().lower()
        if resp and resp != "s":
            print("Edite as pastas de empreiteiros e tente novamente.")
            sys.exit(0)

    # 2. CNO
    cno = args.cno
    if not cno:
        if args.auto:
            print("ERRO: --cno obrigatório no modo --auto")
            sys.exit(1)
        cno = input("CNO da obra (12 digitos): ").strip().replace(".", "").replace("/", "")
    cno = cno.replace(".", "").replace("/", "").replace("-", "")
    if len(cno) != 12 or not cno.isdigit():
        print(f"AVISO: CNO '{cno}' nao tem 12 digitos. Continuando mesmo assim.")

    # 3. Nome da obra
    nome = args.nome
    if not nome:
        if args.auto:
            nome = os.path.basename(base_dir)
            if not nome or nome == ".":
                nome = "Obra sem nome"
        else:
            nome = input("Nome da obra: ").strip()
    if not nome:
        nome = "Obra sem nome"

    # 4. Detectar subpastas SEFIP
    sefip_subs = detect_sefip_subfolders(base_dir, empreiteiros)
    print(f"\nSubpastas SEFIP detectadas: {', '.join(sefip_subs)}")

    # 5. Detectar range de meses
    mes_inicio, mes_fim = detect_month_range(base_dir, empreiteiros, sefip_subs)
    print(f"Meses detectados: {mes_inicio[1]:02d}/{mes_inicio[0]} a {mes_fim[1]:02d}/{mes_fim[0]}")

    if not args.auto:
        resp = input(f"Mes inicio [{mes_inicio[1]:02d}/{mes_inicio[0]}]: ").strip()
        if resp:
            parts = resp.replace("-", "/").split("/")
            if len(parts) == 2:
                mes_inicio = (int(parts[1]), int(parts[0]))
        resp = input(f"Mes fim [{mes_fim[1]:02d}/{mes_fim[0]}]: ").strip()
        if resp:
            parts = resp.replace("-", "/").split("/")
            if len(parts) == 2:
                mes_fim = (int(parts[1]), int(parts[0]))

    # 6. Detectar planilha
    planilha = detect_planilha(base_dir)
    print(f"\nPlanilha: {planilha}")
    if not args.auto:
        resp = input(f"Nome da planilha [{planilha}]: ").strip()
        if resp:
            planilha = resp

    # 7. Montar config
    config = {
        "nome_obra": nome,
        "cno": cno,
        "mes_inicio": list(mes_inicio),
        "mes_fim": list(mes_fim),
        "planilha": planilha,
        "aba_alocacao": "Alocação de colaboradores",
        "aba_resumo": "RESUMO NOVO",
        "sefip_subfolders": sefip_subs,
        "linha_inicio_empr": 5,
        "coluna_inicio_meses": "C",
        "empreiteiros": empreiteiros,
    }

    # 8. Salvar em .sefip-state/ da obra
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    print()
    print(f"Configuracao salva em: {config_path}")
    print(f"  Obra: {nome}")
    print(f"  CNO: {cno}")
    print(f"  Empreiteiros: {len(empreiteiros)}")
    print(f"  Meses: {mes_inicio[1]:02d}/{mes_inicio[0]} a {mes_fim[1]:02d}/{mes_fim[0]}")

    # 9. Criar planilha se não existir
    planilha_path = os.path.join(base_dir, planilha)
    if not os.path.exists(planilha_path):
        print(f"\nPlanilha nao encontrada. Criando: {planilha}")
        try:
            from criar_planilha import criar_planilha
            criar_planilha(config, planilha_path)
            print("Planilha criada com sucesso!")
        except Exception as exc:
            print(f"AVISO: Nao foi possivel criar planilha: {exc}")
            print("  Crie manualmente ou rode novamente apos instalar openpyxl.")

    print()
    print(f"Pronto! Agora rode: python scripts/atualizar_sefip.py --obra-dir \"{base_dir}\"")
    print()


if __name__ == "__main__":
    main()
