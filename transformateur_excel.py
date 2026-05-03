#!/usr/bin/env python3
"""
Transformateur Excel
Lance avec : python3 transformateur_excel.py

Traite tous les fichiers .xlsx/.xlsm du dossier :
  - Supprime les lignes N/A (col F pour 2025/2026, col I ou J pour 2024)
  - Retire les macros
  - Sauvegarde en _officiel.xlsx dans le même dossier
"""

import openpyxl
import os
from pathlib import Path

DOSSIER = Path(__file__).parent

FEUILLE_PRINCIPALE = {2024: "Base 2024", 2025: "Feuil1", 2026: "Base 2026"}
COLONNES_NA        = {2024: [9, 10],      2025: [6],       2026: [6]}


def detecter_annee(nom):
    for a in [2024, 2025, 2026]:
        if str(a) in nom:
            return a
    return None


def derniere_ligne_reelle(ws):
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
            return r
    return 1


def traiter(chemin):
    nom    = chemin.name
    annee  = detecter_annee(nom)
    if annee is None:
        print(f"  ⚠  {nom} ignoré (année 2024/2025/2026 introuvable dans le nom)")
        return

    wb = openpyxl.load_workbook(chemin, keep_vba=False)

    nom_feuille = FEUILLE_PRINCIPALE[annee]
    if nom_feuille not in wb.sheetnames:
        nom_feuille = wb.sheetnames[0]
    ws = wb[nom_feuille]

    fin = derniere_ligne_reelle(ws)

    # Lignes à supprimer
    a_supprimer = []
    for r in range(2, fin + 1):
        for col in COLONNES_NA[annee]:
            if ws.cell(r, col).value == "N/A":
                a_supprimer.append(r)
                break

    # Suppression en ordre inverse
    for r in reversed(a_supprimer):
        ws.delete_rows(r)

    # Supprimer les lignes vides excédentaires
    nouvelle_fin = derniere_ligne_reelle(ws)
    if ws.max_row > nouvelle_fin + 5:
        ws.delete_rows(nouvelle_fin + 1, ws.max_row - nouvelle_fin)

    # Renuméroter colonne A
    num = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value is not None:
            ws.cell(r, 1).value = num
            num += 1

    sortie = DOSSIER / (chemin.stem.replace(".xlsm", "") + "_officiel.xlsx")
    wb.save(sortie)
    print(f"  ✓  {sortie.name}  ({len(a_supprimer)} lignes N/A supprimées)")


# ── Point d'entrée ─────────────────────────────────────────────────────────────

fichiers = sorted(
    f for f in DOSSIER.iterdir()
    if f.suffix.lower() in (".xlsx", ".xlsm")
    and "_officiel" not in f.name
    and not f.name.startswith("~$")
)

if not fichiers:
    print("Aucun fichier Excel trouvé dans le dossier.")
else:
    print(f"Traitement de {len(fichiers)} fichier(s)...\n")
    for f in fichiers:
        print(f"→ {f.name}")
        try:
            traiter(f)
        except Exception as e:
            print(f"  ✗  Erreur : {e}")
    print("\nTerminé.")
