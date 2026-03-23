# NO-INTRO xlsx ROM Filter

Filtre une liste No-Intro (xlsx copié depuis archive.org) et produit un `.txt` d'URLs de téléchargement par console.
Garde la meilleure version de chaque jeu (USA > World > Europe, anglais/français, révision la plus haute).
Exclut demos, prototypes, homebrews, et versions japonaises sans anglais.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/fgouin2014/NO-INTRO_xsls_rom_filter/blob/main/filtrer_roms_colab.ipynb)

---

## Fichiers

| Fichier | Description |
|---|---|
| `filtrer_roms.py` | Script Python local |
| `filtrer_roms.bat` | Drag & drop Windows — glisser un `.xlsx` dessus |
| `filtrer_roms_colab.ipynb` | Version Google Colab — aucune installation requise |

---

## Utilisation locale

### Prérequis
- Python 3 (aucune dépendance externe)
- `filtrer_roms.py` et `filtrer_roms.bat` dans le même dossier

### Drag & Drop
Glisser votre fichier `.xlsx` sur `filtrer_roms.bat`.
Les fichiers `.txt` sont créés dans un sous-dossier `download_lists` à côté du `.xlsx`.

### Ligne de commande
```bash
python3 filtrer_roms.py fichier.xlsx
python3 filtrer_roms.py fichier.xlsx ./mes_listes
```

---

## Utilisation Colab
1. Cliquer le bouton **Open in Colab** ci-dessus
2. Modifier la configuration si nécessaire (Étape 1)
3. Exécuter les cellules dans l'ordre
4. Upload votre `.xlsx` à l'Étape 3
5. Télécharger le `.zip` des résultats à l'Étape 5

---

## Configuration

Dans `filtrer_roms.py` ou dans la cellule **Étape 1** du notebook Colab :

```python
# Régions acceptées
REGIONS_OK = ['USA', 'World', 'Europe', 'Australia']

# Tags exclus
EXCLU_TAGS = ['Beta', 'Demo', 'Proto', 'Sample', 'Alpha',
              'Pirate', 'Unl', 'Aftermarket', 'Evercade',
              'Program', 'Retro-Bit Generations', 'Virtual Console']

# Langues acceptées
LANGUES_OK = ['En', 'Fr']

# Japan : garder seulement si anglais présent
JAPAN_LANGUES_REQUISES = ['En']

# Onglets à ignorer
ONGLETS_IGNORER = [
    # 'No-Intro ROM Sets (2024)',
    # 'Consoles-Portable',
]
```

---

## Logique de filtrage

1. **Région** — doit contenir `USA`, `World`, `Europe` ou `Australia`
2. **Japan** — accepté seulement si `(En)` est présent dans les tags
3. **Tags exclus** — élimine demos, prototypes, homebrews, etc.
4. **Langue** — si un tag de langue est présent, doit contenir `En` ou `Fr`
5. **Dédoublonnage** — garde la meilleure version par titre (USA > World > Europe > Japan, puis langue, puis révision la plus haute)
