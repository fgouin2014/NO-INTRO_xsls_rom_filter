#!/usr/bin/env python3
"""
filtrer_roms.py
===============
Extrait et filtre les URLs de téléchargement de ROMs No-Intro
depuis un fichier xlsx (copié depuis archive.org).

Usage:
    python3 filtrer_roms.py fichier.xlsx
    python3 filtrer_roms.py fichier.xlsx --output ./mes_listes

Produit un fichier .txt par onglet, une URL par ligne.
"""

import zipfile
import re
import os
import sys
import urllib.parse
import xml.etree.ElementTree as ET


# ==============================================================================
# CONFIGURATION — MODIFIER ICI SELON VOS PRÉFÉRENCES
# ==============================================================================

# Onglets à ignorer (laisser vide [] pour tous traiter)
ONGLETS_IGNORER = [
    'No-Intro ROM Sets (2024)',
    'Consoles-Portable',
]

# Régions acceptées (Japan est géré séparément, voir plus bas)
REGIONS_OK = ['USA', 'World', 'Europe','Australia']

# Tags qui entraînent l'exclusion du jeu
EXCLU_TAGS = [
    'Beta', 'Demo', 'Proto', 'Sample', 'Alpha',
    'Pirate', 'Unl', 'Aftermarket', 'Evercade',
    'Program', 'Retro-Bit Generations', 'Virtual Console'
]

# Langues acceptées (priorité dans l'ordre)
LANGUES_OK = ['En', 'Fr']

# Règle Japan : garder seulement si une de ces langues est présente
JAPAN_LANGUES_REQUISES = ['En']

# ==============================================================================
# FIN CONFIGURATION
# ==============================================================================


def get_tags(nom):
    """Extrait tous les contenus entre parenthèses d'un nom de fichier."""
    return re.findall(r'\(([^)]+)\)', nom)


def has_region(nom, regions):
    """Vérifie si une des régions est présente dans les tags."""
    for tag in get_tags(nom):
        for r in regions:
            if re.search(r'\b' + re.escape(r) + r'\b', tag):
                return True
    return False


def has_langue(nom, langues):
    """Vérifie si une des langues est présente dans un tag de langue.
    Un tag de langue = tag contenant uniquement des codes 2 lettres ex: (En,Fr,De)
    """
    for tag in get_tags(nom):
        parts = [p.strip() for p in tag.split(',')]
        if all(re.match(r'^[A-Z][a-z]$', p) for p in parts):
            for l in langues:
                if l in parts:
                    return True
    return False


def is_langue_only_tag(nom):
    """Vérifie si le nom contient un tag de langue (ex: (Ja), (En,Fr,De))."""
    for tag in get_tags(nom):
        parts = [p.strip() for p in tag.split(',')]
        if all(re.match(r'^[A-Z][a-z]$', p) for p in parts):
            return True
    return False


def has_exclu_tag(nom):
    """Vérifie si le nom contient un tag d'exclusion."""
    for tag in get_tags(nom):
        for t in EXCLU_TAGS:
            if re.search(r'\b' + re.escape(t) + r'\b', tag, re.IGNORECASE):
                return True
    return False


def nom_pur(nom):
    """Extrait le titre du jeu sans aucun tag (pour le dédoublonnage)."""
    return re.sub(r'\s*[\(\[].*', '', nom).strip()


def score_region(nom):
    """Score de priorité région — plus bas = meilleur."""
    if has_region(nom, ['USA']) and not re.search(r'\(USA,', nom): return 1  # USA seul
    if re.search(r'\(USA,', nom):                                   return 2  # USA + autres
    if has_region(nom, ['World']):                                  return 3
    if has_region(nom, ['Europe']):                                 return 4
    if has_region(nom, ['Australia']):                              return 5
    return 6


def score_langue(nom):
    """Score de priorité langue — plus bas = meilleur."""
    if has_langue(nom, ['En']): return 1
    if has_langue(nom, ['Fr']): return 2
    return 3  # pas de tag langue = on assume anglais, priorité neutre


def score_rev(nom):
    """Numéro de révision — plus haut = meilleur (on inverse dans le tri)."""
    m = re.search(r'\(Rev (\d+)\)', nom)
    return int(m.group(1)) if m else 0


def est_valide(nom):
    """
    Règles de filtrage complètes.
    Retourne True si le jeu doit être gardé.
    """
    ok_region  = has_region(nom, REGIONS_OK)
    ok_japan   = has_region(nom, ['Japan']) and has_langue(nom, JAPAN_LANGUES_REQUISES)

    # 1. Doit avoir une région valide
    if not (ok_region or ok_japan):
        return False

    # 2. Ne doit pas avoir de tag exclu
    if has_exclu_tag(nom):
        return False

    # 3. Si tag de langue présent, doit contenir En ou Fr
    if is_langue_only_tag(nom) and not has_langue(nom, LANGUES_OK):
        return False

    return True


def trier_et_dedoublonner(jeux):
    """
    Trie par: nom pur > région > langue > révision (desc)
    Puis garde le meilleur par nom pur (dédoublonnage).
    """
    jeux.sort(key=lambda x: (
        nom_pur(x[0]).lower(),
        score_region(x[0]),
        score_langue(x[0]),
        -score_rev(x[0])
    ))

    seen = {}
    for nom, url in jeux:
        np = nom_pur(nom).lower()
        if np not in seen:
            seen[np] = (nom, url)

    return sorted(seen.values(), key=lambda x: x[0].lower())


def extraire_urls_sheet(z, sheet_path, rels_path):
    """Extrait les URLs de la colonne A d'un onglet."""
    if sheet_path not in z.namelist() or rels_path not in z.namelist():
        return []

    rels_content = z.read(rels_path).decode('utf-8')
    rid_to_url = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_content))

    if not rid_to_url:
        return []

    sheet_content = z.read(sheet_path).decode('utf-8')
    hyperlinks = re.findall(r'r:id="(rId\d+)"\s+ref="([^"]+)"', sheet_content)

    col_a = []
    for rid, ref in hyperlinks:
        if re.match(r'^A\d+$', ref):
            url = rid_to_url.get(rid, '')
            if url:
                nom = urllib.parse.unquote(url.split('/')[-1])
                col_a.append((nom, url))  # nom décodé pour filtrer, url encodée pour télécharger

    return col_a


def traiter_xlsx(xlsx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml  = z.read('xl/workbook.xml').decode('utf-8')
        wb_rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        root   = ET.fromstring(wb_xml)
        sheets = root.findall('.//ns:sheet', ns)
        rid_to_file = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', wb_rels))

        for sheet_el in sheets:
            sheet_name = sheet_el.get('name')
            # Ignorer les onglets listés dans ONGLETS_IGNORER
            if sheet_name in ONGLETS_IGNORER:
                print(f"  [{sheet_name}] ignoré (liste d'exclusion)")
                continue
            rid        = sheet_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            sheet_file = rid_to_file.get(rid, '')
            sheet_path = f'xl/{sheet_file}'
            rels_path  = f'xl/worksheets/_rels/{sheet_file.split("/")[-1]}.rels'

            col_a = extraire_urls_sheet(z, sheet_path, rels_path)

            if not col_a:
                print(f"  [{sheet_name}] vide ou sans hyperliens, ignoré")
                continue

            filtres  = [(nom, url) for nom, url in col_a if est_valide(nom)]
            resultat = trier_et_dedoublonner(filtres)

            if not resultat:
                print(f"  [{sheet_name}] 0 jeux après filtrage, ignoré")
                continue

            safe_name = re.sub(r'[^\w\s-]', '', sheet_name).strip().replace(' ', '_')
            out_path  = os.path.join(output_dir, f"{safe_name}.txt")

            with open(out_path, 'w', encoding='utf-8') as f:
                for nom, url in resultat:
                    f.write(url + '\n')

            print(f"  [{sheet_name}] {len(col_a)} total -> {len(resultat)} gardés -> {safe_name}.txt")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    xlsx_path  = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else './download_lists'

    if not os.path.exists(xlsx_path):
        print(f"Erreur: fichier introuvable: {xlsx_path}")
        sys.exit(1)

    print(f"Fichier  : {xlsx_path}")
    print(f"Sortie   : {output_dir}")
    print()

    traiter_xlsx(xlsx_path, output_dir)

    print()
    print("Terminé !")
