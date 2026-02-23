"""
Constantes e helpers — Validador de Estrutura de Obra.

Detecta pastas de obra (estrutura 01-14) e carrega template de validação.
Diferente do constants.py do sefip (que detecta empreiteiros), este módulo
detecta a estrutura padrão de obra com seções numeradas.

Uso:
    import constants_validador as cv
    obras = cv.descobrir_obras(search_dir)
    template = cv.load_template()
"""

import json
import os
import re
import fnmatch

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, ".."))
PLUGIN_DIR = os.path.normpath(os.path.join(SKILL_DIR, "..", ".."))

# Template padrão incluso no skill
DEFAULT_TEMPLATE_PATH = os.path.join(_SCRIPT_DIR, "template_obra.json")

# Regex para pastas numeradas no padrão de obra: "NN - DESCRIÇÃO"
_OBRA_SECTION_RE = re.compile(r"^(\d{2})\s*-\s+(.+)$")

# Prefixos conhecidos de pastas de obra no nível raiz
_OBRA_PREFIX_RE = re.compile(r"^OBRA\s+", re.IGNORECASE)

# Pastas que NÃO são obras (no nível raiz da Pasta Mãe)
_NON_OBRA_NAMES = {
    "ENGENHARIA", "ARQUITETURA (PROJETOS)", "SUPRIMENTOS",
    "NOTAS FISCAIS OBRAS", "MATERIAL EMPREENDIMENTOS",
    "PLANEJAMENTO DE OBRAS", "IPTU 2026", "ECBIESEK",
    "BIE EMPREENDIMENTOS", "WALE-LUZ", "MODELO PASTA PADRÃO",
}


# ---------------------------------------------------------------------------
# Template
# ---------------------------------------------------------------------------

def load_template(path=None):
    """Carrega template de estrutura de obra.

    Args:
        path: Caminho para template JSON customizado.
              Se None, usa o template padrão incluso no skill.

    Returns:
        dict com 'version', 'description', 'structure'
    """
    template_path = path or DEFAULT_TEMPLATE_PATH
    with open(template_path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

def match_pattern(folder_name, pattern):
    """Verifica se folder_name corresponde ao pattern (case-insensitive, glob-like).

    Suporta * como wildcard (via fnmatch).
    Ex: "02 - DOCUMENTOS DA EMPRESA (SPE)" matches "02 - DOCUMENTOS DA EMPRESA*"
    """
    return fnmatch.fnmatch(folder_name.upper(), pattern.upper())


def match_file_pattern(filename, pattern):
    """Verifica se filename corresponde ao pattern de arquivo esperado.

    Case-insensitive, suporta * wildcard.
    Ex: "CNO_LAGUNAS.pdf" matches "*CNO*"
    """
    return fnmatch.fnmatch(filename.upper(), pattern.upper())


# ---------------------------------------------------------------------------
# Detecção de obras
# ---------------------------------------------------------------------------

def _count_obra_sections(path):
    """Conta quantas subpastas seguem o padrão 'NN - DESCRIÇÃO' num diretório."""
    count = 0
    try:
        for entry in os.listdir(path):
            if os.path.isdir(os.path.join(path, entry)):
                if _OBRA_SECTION_RE.match(entry):
                    count += 1
    except OSError:
        pass
    return count


def is_obra_dir(path):
    """Verifica se um diretório é uma pasta de obra.

    Uma pasta de obra contém pelo menos 5 subpastas no padrão 'NN - DESCRIÇÃO'
    (ex: '01 - DOCUMENTOS DA OBRA', '02 - DOCUMENTOS DA EMPRESA', ...).
    """
    return _count_obra_sections(path) >= 5


def descobrir_obras(search_dir):
    """Descobre pastas de obra a partir de um diretório.

    Procura em 2 níveis de profundidade por diretórios que contenham
    subpastas no padrão 'NN - DESCRIÇÃO' (ex: '01 - DOCUMENTOS DA OBRA').

    Retorna lista de dicts: [{"path": str, "name": str, "sections": int}]
    """
    obras = []
    seen = set()
    search_dir = os.path.normpath(search_dir)

    def _add_if_obra(path):
        norm = os.path.normpath(path)
        if norm in seen:
            return
        n = _count_obra_sections(path)
        if n >= 5:
            seen.add(norm)
            name = os.path.basename(norm)
            obras.append({
                "path": norm,
                "name": name,
                "sections": n,
            })

    # Nível 0: o próprio diretório
    _add_if_obra(search_dir)

    # Nível 1: filhos diretos
    try:
        for entry in sorted(os.listdir(search_dir)):
            child = os.path.join(search_dir, entry)
            if os.path.isdir(child) and not entry.startswith("."):
                _add_if_obra(child)
    except OSError:
        pass

    # Nível 2: netos (ex: PASTA MÃE / OBRA IMPERIAL / subpasta)
    try:
        for entry in sorted(os.listdir(search_dir)):
            child = os.path.join(search_dir, entry)
            if not os.path.isdir(child) or entry.startswith("."):
                continue
            for sub in sorted(os.listdir(child)):
                grandchild = os.path.join(child, sub)
                if os.path.isdir(grandchild) and not sub.startswith("."):
                    _add_if_obra(grandchild)
    except OSError:
        pass

    return obras


# ---------------------------------------------------------------------------
# Listagem de conteúdo
# ---------------------------------------------------------------------------

def list_subfolders(path):
    """Lista subpastas de um diretório (nome, não caminho completo)."""
    try:
        return sorted([
            entry for entry in os.listdir(path)
            if os.path.isdir(os.path.join(path, entry))
            and not entry.startswith(".")
        ])
    except OSError:
        return []


def list_files(path):
    """Lista arquivos (não pastas) de um diretório."""
    try:
        return sorted([
            entry for entry in os.listdir(path)
            if os.path.isfile(os.path.join(path, entry))
            and not entry.startswith(".")
        ])
    except OSError:
        return []


def is_folder_empty(path):
    """Verifica se uma pasta está efetivamente vazia (sem arquivos nem subpastas visíveis)."""
    try:
        for entry in os.listdir(path):
            if not entry.startswith("."):
                return False
    except OSError:
        pass
    return True


def count_files_recursive(path):
    """Conta arquivos recursivamente num diretório."""
    total = 0
    try:
        for root, dirs, files in os.walk(path):
            # Ignorar pastas ocultas
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            total += len([f for f in files if not f.startswith(".")])
    except OSError:
        pass
    return total
