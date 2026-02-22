"""
Constantes compartilhadas — NFSe Extractor.

Carrega configuração de obra.json (se existir) ou usa fallback Lagoa Clube Resort.
Exporta: CNO, OBRA_NOME, PLANILHA_FILENAME, ABA_NFSE, NF_SUBFOLDERS,
         EMPR_FOLDERS, EMPR_NAMES, ALL_EMPR,
         SKILL_DIR, BASE_DIR, STATE_DIR, XLSX_PATH

Uso agnóstico:
    import constants_nfse
    constants_nfse.init_obra("/caminho/para/pasta/prestadores")
    # Todos os módulos que importam constants_nfse verão os valores atualizados
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

def _load_config_from(path):
    """Carrega obra.json de um caminho específico, retorna dict ou None."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return None


def _apply_config(cfg):
    """Aplica configuração de obra.json nas variáveis globais."""
    global CNO, OBRA_NOME, PLANILHA_FILENAME, ABA_NFSE
    global NF_SUBFOLDERS, EMPR_FOLDERS, EMPR_NAMES, ALL_EMPR
    CNO = cfg["cno"]
    OBRA_NOME = cfg["nome_obra"]
    PLANILHA_FILENAME = cfg.get("planilha_nfse",
                                "CONTROLE GERAL DE NOTAS FISCAIS DE EMPREITEIROS.xlsx")
    ABA_NFSE = cfg.get("aba_nfse", "NOTAS FISCAIS (NOVO)")
    NF_SUBFOLDERS = cfg.get("nf_subfolders", ["NOTA FISCAL"])
    _emprs = cfg["empreiteiros"]
    EMPR_FOLDERS = {e["num"]: e["pasta"] for e in _emprs}
    EMPR_NAMES = {e["num"]: e["nome_curto"] for e in _emprs}
    ALL_EMPR = sorted(EMPR_FOLDERS.keys())


def _set_defaults_lagoa():
    """Define constantes padrão da Lagoa Clube Resort (retrocompatível)."""
    global CNO, OBRA_NOME, PLANILHA_FILENAME, ABA_NFSE
    global NF_SUBFOLDERS, EMPR_FOLDERS, EMPR_NAMES, ALL_EMPR
    CNO = "900152252672"
    OBRA_NOME = "Lagoa Clube Resort"
    PLANILHA_FILENAME = "CONTROLE GERAL DE NOTAS FISCAIS DE EMPREITEIROS.xlsx"
    ABA_NFSE = "NOTAS FISCAIS (NOVO)"
    NF_SUBFOLDERS = ["NOTA FISCAL"]
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
    ALL_EMPR = sorted(EMPR_FOLDERS.keys())


# ---------------------------------------------------------------------------
# Carga inicial (retrocompatível)
# ---------------------------------------------------------------------------

_cfg = _load_config_from(_CONFIG_PATH)
if _cfg:
    _apply_config(_cfg)
else:
    _set_defaults_lagoa()


# ---------------------------------------------------------------------------
# Caminhos centralizados
# ---------------------------------------------------------------------------

# BASE_DIR: pasta da obra (onde ficam as pastas de empreiteiros)
BASE_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))

# STATE_DIR: onde ficam os arquivos de estado (.json)
STATE_DIR = os.path.join(SKILL_DIR, "state")

# XLSX_PATH: caminho completo para a planilha de NFs
XLSX_PATH = os.path.normpath(os.path.join(BASE_DIR, PLANILHA_FILENAME))


# ---------------------------------------------------------------------------
# Detecção de obras (copiado de SEFIP constants.py — funções genéricas)
# ---------------------------------------------------------------------------

_EMPR_MARKERS = {
    "SEFIP", "NOTA FISCAL", "DOCUMENTOS MENSAIS", "DOCUMENTAÇÕES MENSAIS",
    "ENTREGA DE DOCUMENTOS MENSAIS", "DOMENTAÇÃO MENSAL", "CONTRACHEQUE",
    "FGTS", "INSS",
}


def _is_empreiteiro_dir(path):
    """Verifica se uma pasta numerada é realmente de empreiteiro."""
    try:
        subs = {s.upper() for s in os.listdir(path) if os.path.isdir(os.path.join(path, s))}
    except OSError:
        return False
    return bool(subs & _EMPR_MARKERS)


def is_obra_dir(path):
    """Verifica se um diretório contém pastas de empreiteiros numeradas."""
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
    pastas de empreiteiros (ex: '01 JVB/NOTA FISCAL/').

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
        base = os.path.basename(path)
        parent = os.path.basename(os.path.dirname(path))
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
            obras.append({"path": norm, "name": _friendly_name(norm), "empreiteiros": n})

    def _scan_children(parent):
        try:
            for entry in sorted(os.listdir(parent)):
                child = os.path.join(parent, entry)
                if os.path.isdir(child) and not entry.startswith("."):
                    _add_if_obra(child)
        except OSError:
            pass

    _add_if_obra(search_dir)
    _scan_children(search_dir)

    try:
        for entry in sorted(os.listdir(search_dir)):
            child = os.path.join(search_dir, entry)
            if os.path.isdir(child) and not entry.startswith("."):
                _scan_children(child)
    except OSError:
        pass

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
    """Migra nfse_extracted.json de locais antigos para .nfse-state/."""
    target = os.path.join(new_state_dir, "nfse_extracted.json")
    if os.path.exists(target):
        return

    old_locations = [
        os.path.join(os.path.dirname(new_state_dir), "nfse_extracted.json"),  # obra root
        os.path.join(SKILL_DIR, "nfse_extracted.json"),
        os.path.join(SKILL_DIR, "scripts", "nfse_extracted.json"),
        os.path.join(SKILL_DIR, "state", "nfse_extracted.json"),
    ]
    for old in old_locations:
        if os.path.exists(old):
            import shutil
            shutil.copy2(old, target)
            print(f"  Migrado nfse_extracted.json de {old} -> {target}")
            return


def init_obra(obra_dir):
    """Reinicializa todas as constantes para uma pasta de obra específica.

    Atualiza BASE_DIR, STATE_DIR, XLSX_PATH, e carrega obra.json se existir.
    Busca config em: .nfse-state/obra.json -> .sefip-state/obra.json -> skill/obra.json

    Args:
        obra_dir: Caminho para a pasta que contém as pastas de empreiteiros
    """
    global BASE_DIR, STATE_DIR, XLSX_PATH

    BASE_DIR = os.path.normpath(obra_dir)

    STATE_DIR = os.path.join(BASE_DIR, ".nfse-state")
    os.makedirs(STATE_DIR, exist_ok=True)

    _migrate_state_if_needed(STATE_DIR)

    # Buscar config: .nfse-state/ -> .sefip-state/ -> skill/
    cfg = (_load_config_from(os.path.join(STATE_DIR, "obra.json"))
           or _load_config_from(os.path.join(BASE_DIR, ".sefip-state", "obra.json"))
           or _load_config_from(_CONFIG_PATH))
    if cfg:
        _apply_config(cfg)
    else:
        _set_defaults_lagoa()

    XLSX_PATH = os.path.normpath(os.path.join(BASE_DIR, PLANILHA_FILENAME))
