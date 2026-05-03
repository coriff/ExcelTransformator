#!/usr/bin/env python3
"""
Transformateur Excel
Lance avec : python3 transformateur_excel.py

Traite tous les fichiers .xlsx/.xlsm du dossier :
  - Remplace les formules par leurs valeurs calculées (évite les références circulaires)
  - Supprime les lignes N/A (col F pour 2025/2026, col I ou J pour 2024)
  - Retire les macros
  - Sauvegarde en _officiel.xlsx dans le même dossier
"""

import openpyxl
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
    """
    Cherche la dernière ligne avec données via ws._cells (dict interne),
    sans créer de Cell objects pour les lignes vides — crucial pour les
    fichiers comme 2026 qui déclarent 17 000+ lignes dans leur XML.
    """
    max_r = 1
    for (row, col), cell in ws._cells.items():
        if cell.value is not None and row > max_r:
            max_r = row
    return max_r


def traiter(chemin):
    nom   = chemin.name
    annee = detecter_annee(nom)
    if annee is None:
        print(f"  ⚠  {nom} ignoré (2024/2025/2026 introuvable dans le nom)")
        return

    # Chargement en double :
    #   data_only=True  → valeurs calculées (sans formules)
    #   keep_vba=False  → mise en forme + structure, sans macros
    wb_vals = openpyxl.load_workbook(chemin, data_only=True)
    wb      = openpyxl.load_workbook(chemin, keep_vba=False)

    nom_feuille = FEUILLE_PRINCIPALE[annee]
    if nom_feuille not in wb.sheetnames:
        nom_feuille = wb.sheetnames[0]

    ws      = wb[nom_feuille]
    ws_vals = wb_vals[nom_feuille]

    # ── Étape 1 : vraie dernière ligne (via _cells, sans toucher aux vides) ──
    fin = derniere_ligne_reelle(ws)

    # ── Étape 2 : supprimer les lignes vides AVANT tout le reste ────────────
    # Critique pour le 2026 : réduit max_row de 17 552 à ~1 026.
    # Chaque delete_rows suivant ne reconstruit alors que ~1 026 lignes
    # au lieu de 17 552 → 17× plus rapide.
    if ws.max_row > fin + 5:
        ws.delete_rows(fin + 1, ws.max_row - fin)

    # ── Étape 3 : remplacer les formules par leurs valeurs ──────────────────
    # Limité aux vraies données (max_row=fin) pour ne pas créer
    # des Cell objects inutiles pour les lignes vides.
    for row in ws.iter_rows(max_row=fin):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = ws_vals.cell(cell.row, cell.column).value

    # ── Étape 4 : repérer les lignes N/A ────────────────────────────────────
    a_supprimer = []
    for r in range(2, fin + 1):
        for col in COLONNES_NA[annee]:
            if ws.cell(r, col).value == "N/A":
                a_supprimer.append(r)
                break

    # ── Étape 5 : supprimer les N/A (max_row ≈ 1 026 maintenant → rapide) ──
    for r in reversed(a_supprimer):
        ws.delete_rows(r)

    # ── Étape 6 : renuméroter la colonne A (#) ──────────────────────────────
    # Via _cells pour ne pas créer d'objets sur les lignes déjà supprimées.
    col_a = sorted(
        (row, cell) for (row, col), cell in ws._cells.items()
        if col == 1 and row > 1 and cell.value is not None
    )
    for num, (_, cell) in enumerate(col_a, start=1):
        cell.value = num

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
