"""
Constantes compartilhadas — SEFIP Extractor.

Carrega configuração de obra.json (se existir) ou usa fallback Lagoa Clube Resort.
Exporta: CNO, OBRA_NOME, MONTHS_LIST, MONTH_COL, MONTH_COL_IDX, COL_LETTERS,
         TOTAL_MONTHS, EMPR_FOLDERS, EMPR_NAMES, EMPR_NAMES_FULL, EMPR_ROW,
         ALL_EMPR, PLANILHA_FILENAME, ABA_ALOCACAO, ABA_RESUMO, SEFIP_SUBFOLDERS,
         EMPR_ROW_START, SKILL_DIR, BASE_DIR, STATE_DIR, XLSX_PATH

Uso agnóstico:
    import constants
    constants.init_obra("/caminho/para/pasta/prestadores")
    # Todos os módulos que importam constants verão os valores atualizados
"""

import json
import os
import re

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, ".."))
_CONFIG_PATH = os.path.join(SKILL_DIR, "obra.json")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _generate_months(start, end):
    """Gera lista de (ano, mes) de start a end inclusive."""
    months = []
    y, m = start
    ey, em = end
    while (y, m) <= (ey, em):
        months.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def _col_letter(index):
    """Converte índice 0-based em letra(s) de coluna Excel. 0=A, 25=Z, 26=AA."""
    result = ""
    n = index
    while True:
        result = chr(65 + n % 26) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def _load_config_from(path):
    """Carrega obra.json de um caminho específico, retorna dict ou None."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return None


def _recompute_derived():
    """Recalcula variáveis derivadas (TOTAL_MONTHS, COL_LETTERS, MONTH_COL, etc.)."""
    global TOTAL_MONTHS, ALL_EMPR, COL_LETTERS, MONTH_COL, MONTH_COL_IDX
    TOTAL_MONTHS = len(MONTHS_LIST)
    ALL_EMPR = sorted(EMPR_FOLDERS.keys())
    _COL_START = 2  # C = index 2 (A=0, B=1, C=2)
    COL_LETTERS = [_col_letter(_COL_START + i) for i in range(TOTAL_MONTHS)]
    MONTH_COL = {}
    for i, (y, m) in enumerate(MONTHS_LIST):
        MONTH_COL[(y, m)] = COL_LETTERS[i]
    MONTH_COL_IDX = {}
    for i, (y, m) in enumerate(MONTHS_LIST):
        MONTH_COL_IDX[f"{y}-{m:02d}"] = i + 3  # col C = idx 3


def _apply_config(cfg):
    """Aplica configuração de obra.json nas variáveis globais."""
    global CNO, OBRA_NOME, PLANILHA_FILENAME, ABA_ALOCACAO, ABA_RESUMO
    global SEFIP_SUBFOLDERS, EMPR_ROW_START, MONTHS_LIST
    global EMPR_FOLDERS, EMPR_NAMES, EMPR_NAMES_FULL, EMPR_ROW
    CNO = cfg["cno"]
    OBRA_NOME = cfg["nome_obra"]
    PLANILHA_FILENAME = cfg.get("planilha", "Controle de Alocação e ISS mensal de empreiteiros.xlsx")
    ABA_ALOCACAO = cfg.get("aba_alocacao", "Alocação de colaboradores")
    ABA_RESUMO = cfg.get("aba_resumo", "RESUMO NOVO")
    SEFIP_SUBFOLDERS = cfg.get("sefip_subfolders", [
        "SEFIP", "DOCUMENTOS MENSAIS", "DOCUMENTAÇÕES MENSAIS",
        "ENTREGA DE DOCUMENTOS MENSAIS", "DOMENTAÇÃO MENSAL",
    ])
    EMPR_ROW_START = cfg.get("linha_inicio_empr", 5)
    _start = tuple(cfg["mes_inicio"])
    _end = tuple(cfg["mes_fim"])
    MONTHS_LIST = _generate_months(_start, _end)
    _emprs = cfg["empreiteiros"]
    EMPR_FOLDERS = {e["num"]: e["pasta"] for e in _emprs}
    EMPR_NAMES = {e["num"]: e["nome_curto"] for e in _emprs}
    EMPR_NAMES_FULL = {str(e["num"]): e["nome_completo"] for e in _emprs}
    EMPR_ROW = {e["num"]: EMPR_ROW_START + i for i, e in enumerate(_emprs)}


def _set_defaults_lagoa():
    """Define constantes padrão da Lagoa Clube Resort (retrocompatível)."""
    global CNO, OBRA_NOME, PLANILHA_FILENAME, ABA_ALOCACAO, ABA_RESUMO
    global SEFIP_SUBFOLDERS, EMPR_ROW_START, MONTHS_LIST
    global EMPR_FOLDERS, EMPR_NAMES, EMPR_NAMES_FULL, EMPR_ROW
    CNO = "900152252672"
    OBRA_NOME = "Lagoa Clube Resort"
    PLANILHA_FILENAME = "Controle de Alocação e ISS mensal de empreiteiros.xlsx"
    ABA_ALOCACAO = "Alocação de colaboradores"
    ABA_RESUMO = "RESUMO NOVO"
    SEFIP_SUBFOLDERS = [
        "SEFIP", "DOCUMENTOS MENSAIS", "DOCUMENTAÇÕES MENSAIS",
        "ENTREGA DE DOCUMENTOS MENSAIS", "DOMENTAÇÃO MENSAL",
    ]
    EMPR_ROW_START = 5
    MONTHS_LIST = _generate_months((2023, 8), (2026, 2))
    EMPR_FOLDERS = {
        1: "01 JVB", 2: "02 JR CONSTRUÇÃO",
        3: "03 G & M CONSTRUÇÕES (MATEUS)", 4: "04 BELO E SANTOS CONSTRUÇÕES",
        5: "05 DEMORAES ENCANADORES LTDA (MANOEL)", 6: "06 A P DE LIMA (ALMIR)",
        7: "07 KLEYSON RIBEIRO LTDA", 8: "08 L C DA S DOS SANTOS (ALEX)",
        9: "09 S S M PRESTACAO DE SERVICOS LTDA (SILVIO)",
        10: "10 C F N PRESTACAO DE SERVICOS LTDA (CLEUTON)",
        11: "11 AQUILAIS ATIVIDADES DE VIGILANCIA", 12: "12 DELCINEY NOGUEIRA BRASIL",
    }
    EMPR_NAMES = {
        1: "JVB", 2: "JR CONSTRUÇÃO", 3: "G & M CONSTRUÇÕES",
        4: "BELO E SANTOS", 5: "DEMORAES ENCANADORES",
        6: "A P DE LIMA (ALMIR)", 7: "KLEYSON RIBEIRO",
        8: "L C DA S DOS SANTOS (ALEX)", 9: "S S M PRESTACAO (SILVIO)",
        10: "C F N PRESTACAO (CLEUTON)", 11: "AQUILAIS VIGILANCIA",
        12: "DELCINEY NOGUEIRA",
    }
    EMPR_NAMES_FULL = {
        "1": "01 JVB", "2": "02 JR CONSTRUÇÃO",
        "3": "03 G & M CONSTRUÇÕES (MATEUS)", "4": "04 BELO E SANTOS CONSTRUÇÕES",
        "5": "05 DEMORAES ENCANADORES LTDA", "6": "06 A P DE LIMA (ALMIR)",
        "7": "07 KLEYSON RIBEIRO LTDA", "8": "08 L C DA S DOS SANTOS (ALEX)",
        "9": "09 S S M PRESTACAO DE SERVICOS LTDA",
        "10": "10 C F N PRESTACAO DE SERVICOS LTDA",
        "11": "11 AQUILAIS ATIVIDADES DE VIGILANCIA",
        "12": "12 DELCINEY NOGUEIRA BRASIL",
    }
    EMPR_ROW = {i: i + 4 for i in range(1, 13)}


# ---------------------------------------------------------------------------
# Carga inicial (retrocompatível)
# ---------------------------------------------------------------------------

_cfg = _load_config_from(_CONFIG_PATH)
if _cfg:
    _apply_config(_cfg)
else:
    _set_defaults_lagoa()

_recompute_derived()


# ---------------------------------------------------------------------------
# Caminhos centralizados
# ---------------------------------------------------------------------------

# BASE_DIR: pasta da obra (onde ficam as pastas de empreiteiros)
# Default: 3 níveis acima de scripts/ (scripts → skill → .skills → obra)
BASE_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))

# STATE_DIR: onde ficam os arquivos de estado (.json)
# Default: state/ dentro da skill (retrocompatível)
STATE_DIR = os.path.join(SKILL_DIR, "state")

# XLSX_PATH: caminho completo para a planilha-mestre
XLSX_PATH = os.path.normpath(os.path.join(BASE_DIR, PLANILHA_FILENAME))


# ---------------------------------------------------------------------------
# Detecção de obras
# ---------------------------------------------------------------------------

# Subpastas que indicam que uma pasta numerada é de empreiteiro (não seção de obra)
_EMPR_MARKERS = {
    "SEFIP", "NOTA FISCAL", "DOCUMENTOS MENSAIS", "DOCUMENTAÇÕES MENSAIS",
    "ENTREGA DE DOCUMENTOS MENSAIS", "DOMENTAÇÃO MENSAL", "CONTRACHEQUE",
    "FGTS", "INSS",
}


def _is_empreiteiro_dir(path):
    """Verifica se uma pasta numerada é realmente de empreiteiro.

    Uma pasta de empreiteiro contém subpastas como SEFIP, NOTA FISCAL, etc.
    Já seções de obra ('01 - DOCUMENTOS DA OBRA') não contêm essas subpastas.
    """
    try:
        subs = {s.upper() for s in os.listdir(path) if os.path.isdir(os.path.join(path, s))}
    except OSError:
        return False
    return bool(subs & _EMPR_MARKERS)


def is_obra_dir(path):
    """Verifica se um diretório contém pastas de empreiteiros numeradas.

    Diferente de diretórios de obra (como 'OBRA IMPERIAL') que têm pastas
    numeradas de seções ('01 - DOCUMENTOS', '02 - EMPRESA'), aqui verificamos
    que pelo menos 2 subpastas numeradas contêm marcadores de empreiteiro.
    """
    count = 0
    try:
        for entry in os.listdir(path):
            full = os.path.join(path, entry)
            if os.path.isdir(full) and re.match(r'^\d{1,2}\s', entry):
                if _is_empreiteiro_dir(full):
                    count += 1
    except OSError:
        return False
    return count >= 2


def descobrir_obras(search_dir):
    """Descobre pastas de prestadores de obras a partir de um diretório.

    Procura em até 3 níveis de profundidade por diretórios que contenham
    pastas de empreiteiros (ex: '01 JVB/SEFIP/', '02 JR CONSTRUÇÃO/NOTA FISCAL/').

    O nome exibido inclui contexto do diretório-pai para distinguir obras
    com subpastas nomeadas igualmente (ex: '07 - PRESTADORES DE SERVIÇO').

    Retorna lista de dicts: [{"path": str, "name": str, "empreiteiros": int}]
    """
    obras = []
    seen = set()
    search_dir = os.path.normpath(search_dir)

    def _count_empr(path):
        count = 0
        try:
            for entry in os.listdir(path):
                full = os.path.join(path, entry)
                if os.path.isdir(full) and re.match(r'^\d{1,2}\s', entry):
                    if _is_empreiteiro_dir(full):
                        count += 1
        except OSError:
            pass
        return count

    def _friendly_name(path):
        """Gera nome amigável incluindo pai se necessário."""
        base = os.path.basename(path)
        parent = os.path.basename(os.path.dirname(path))
        # Se o pai é genérico (OneDrive, search_dir), usar só o nome
        if os.path.normpath(os.path.dirname(path)) == search_dir:
            return base
        return f"{parent} > {base}"

    def _add_if_obra(path):
        norm = os.path.normpath(path)
        if norm in seen:
            return
        n = _count_empr(path)
        if n >= 2:
            seen.add(norm)
            obras.append({
                "path": norm,
                "name": _friendly_name(norm),
                "empreiteiros": n,
            })

    def _scan_children(parent):
        try:
            for entry in sorted(os.listdir(parent)):
                child = os.path.join(parent, entry)
                if os.path.isdir(child) and not entry.startswith("."):
                    _add_if_obra(child)
        except OSError:
            pass

    # Nível 0: o próprio diretório
    _add_if_obra(search_dir)

    # Nível 1: filhos diretos
    _scan_children(search_dir)

    # Nível 2: netos
    try:
        for entry in sorted(os.listdir(search_dir)):
            child = os.path.join(search_dir, entry)
            if os.path.isdir(child) and not entry.startswith("."):
                _scan_children(child)
    except OSError:
        pass

    # Nível 3: bisnetos (ex: PASTA MÃE / OBRA IMPERIAL / 07 - PRESTADORES)
    try:
        for entry in sorted(os.listdir(search_dir)):
            child = os.path.join(search_dir, entry)
            if not os.path.isdir(child) or entry.startswith("."):
                continue
            for sub in sorted(os.listdir(child)):
                grandchild = os.path.join(child, sub)
                if not os.path.isdir(grandchild) or sub.startswith("."):
                    continue
                _scan_children(grandchild)
    except OSError:
        pass

    return obras


# ---------------------------------------------------------------------------
# init_obra — reinicializa para uma obra específica
# ---------------------------------------------------------------------------

def _migrate_state_if_needed(new_state_dir):
    """Migra state files do local antigo (<skill>/state/) para o novo (.sefip-state/).

    Copia arquivos do diretório antigo se o novo estiver vazio.
    Executa apenas uma vez — após a migração, o novo diretório terá os arquivos.
    """
    old_state_dir = os.path.join(SKILL_DIR, "state")
    if not os.path.isdir(old_state_dir):
        return

    state_files = [
        "extractions_text.json", "extractions_ocr.json",
        "needs_ocr.json", "divergences.json", "changes_log.json",
    ]

    # Verificar se já há state files no novo local
    has_new = any(os.path.exists(os.path.join(new_state_dir, f)) for f in state_files)
    if has_new:
        return

    # Verificar se há algo para migrar
    has_old = any(os.path.exists(os.path.join(old_state_dir, f)) for f in state_files)
    if not has_old:
        return

    import shutil
    migrated = 0
    for f in state_files:
        src = os.path.join(old_state_dir, f)
        dst = os.path.join(new_state_dir, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            migrated += 1

    if migrated:
        print(f"  Migrados {migrated} state files de {old_state_dir} → {new_state_dir}")


def init_obra(obra_dir):
    """Reinicializa todas as constantes para uma pasta de obra específica.

    Deve ser chamado ANTES de importar funções dos outros scripts.
    Atualiza BASE_DIR, STATE_DIR, XLSX_PATH, e carrega obra.json se existir.

    Args:
        obra_dir: Caminho para a pasta que contém as pastas de empreiteiros
    """
    global BASE_DIR, STATE_DIR, XLSX_PATH

    BASE_DIR = os.path.normpath(obra_dir)

    # State fica em .sefip-state/ dentro da pasta da obra
    STATE_DIR = os.path.join(BASE_DIR, ".sefip-state")
    os.makedirs(STATE_DIR, exist_ok=True)

    # Migrar state files do local antigo se necessário
    _migrate_state_if_needed(STATE_DIR)

    # Tentar carregar obra.json: primeiro de .sefip-state/, depois da skill
    config_path_obra = os.path.join(STATE_DIR, "obra.json")
    config_path_skill = _CONFIG_PATH

    cfg = _load_config_from(config_path_obra) or _load_config_from(config_path_skill)
    if cfg:
        _apply_config(cfg)
    else:
        _set_defaults_lagoa()

    _recompute_derived()
    XLSX_PATH = os.path.normpath(os.path.join(BASE_DIR, PLANILHA_FILENAME))
