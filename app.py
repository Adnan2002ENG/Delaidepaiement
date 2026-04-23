import io
import re
from dataclasses import dataclass
from typing import Optional, List, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Déclarations délai de paiement", layout="wide")
st.title("Application de traitement des déclarations de délai de paiement")

EXPECTED_COL_INDEX = {
    "col_a": 0,
    "date": 1,
    "journal": 2,
    "piece": 3,
    "libelle": 4,
    "debit": 5,
    "credit": 6,
    "lettrage": 7
}

# Colonnes GL PENNYLAND (indices 0-basés, fichier lu avec header=1)
PENNYLAND_COL_INDEX = {
    "account":     0,   # N° de compte  → code / clé de groupement fournisseur
    "name":        1,   # Libellé de compte → nom fournisseur
    "date":        2,   # Date
    "journal":     3,   # Journal  (AC=facture, BQ*=paiement, AA=report à nouveau)
    "libelle":     4,   # Libellé de pièce
    "num_facture": 7,   # N° de facture  (col H)
    "lettrage":    9,   # Let.  (col J)
    "debit":       10,  # Débit  (col K)
    "credit":      11,  # Crédit (col L)
}

# Colonnes GL SAGE (indices 0-basés, fichier lu avec header=1)
SAGE_COL_INDEX = {
    "compte":   0,  # A - N° COMPTE → code fournisseur (groupement par changement)
    "date":     1,  # B - DATE
    "journal":  2,  # C - CODE J
    "piece":    3,  # D - N° DE PIECE
    "libelle":  4,  # E - LIBELLE
    "lettrage": 5,  # F - LETTRAGE
    "debit":    6,  # G - DEBIT (paiement ou avoir)
    "credit":   7,  # H - CREDIT (facture)
}


# Année plancher pour la prise en compte des factures.
# Valeur par défaut : 2025 (première année obligatoire de déclaration).
# L'UI peut la descendre à 2024 via la checkbox "Inclure les factures 2024".
MIN_INVOICE_YEAR = 2025


def _min_year() -> int:
    try:
        return int(st.session_state.get("include_2024", False) and 2024 or MIN_INVOICE_YEAR)
    except Exception:
        return MIN_INVOICE_YEAR


@dataclass
class SupplierBoundary:
    supplier_code: str
    supplier_name: str
    start_row: int
    end_row: int


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_amount(value) -> float:
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, str):
        value = value.replace(" ", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return 0.0


def safe_date(value):
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def extract_supplier_info(cell_value: str) -> Optional[Tuple[str, str]]:
    text = normalize_text(cell_value)
    if not text.startswith("9"):
        return None

    first_space = text.find(" ")
    if first_space == -1:
        return None

    supplier_code = text[:first_space].strip()
    supplier_name = text[first_space + 1:].strip()

    if not supplier_code or not supplier_name:
        return None

    return supplier_code, supplier_name


def is_supplier_total(cell_value: str) -> bool:
    return normalize_text(cell_value).lower().startswith("total")


def is_opening_balance_row(text: str) -> bool:
    """Détecte une ligne solde d'ouverture : 'Total au JJ/MM/AAAA'."""
    t = normalize_text(text).lower().replace(" ", "")
    return bool(re.match(r"totalau\d{1,2}/\d{1,2}/\d{2,4}", t))


def classify_remark(delay_days):
    if delay_days is None or pd.isna(delay_days):
        return "Date facture manquante - à vérifier"
    if delay_days < 0:
        return "Avance - Rien à signaler"
    if delay_days <= 60:
        return "Rien à signaler"
    if 61 <= delay_days <= 120:
        return "Demander la convention de délai de paiement du fournisseur en cas d'existence, sinon pénalité"
    return "Pénalité à payer"


def is_internal_piece_reference(piece_text: str) -> bool:
    t = normalize_text(piece_text).upper().replace(" ", "")
    if not t:
        return False
    if len(t) >= 8 and t[:2].isalpha() and t[2:].isdigit():
        return True
    return False


def extract_invoice_from_libelle(libelle_text: str) -> str:
    lib = normalize_text(libelle_text)
    if not lib:
        return ""

    patterns = [
        r"\b(?:FAC|FACT|FACTURE|INV|INVOICE|AVOIR|BL|BC|CMD|N°|NO)?[-/\s]*([A-Z]{0,5}\d{3,}(?:[-/]\d+)*)\b",
        r"\b([A-Z]{1,5}-\d{3,}(?:-\d+)*)\b",
        r"\b(\d{4,})\b",
    ]

    for pattern in patterns:
        m = re.search(pattern, lib, flags=re.IGNORECASE)
        if m:
            if m.lastindex:
                return m.group(1).strip()
            return m.group(0).strip()

    return ""


def extract_invoice_number(piece_text: str, libelle_text: str) -> str:
    piece = normalize_text(piece_text)
    libelle = normalize_text(libelle_text)

    if piece and is_internal_piece_reference(piece):
        extracted = extract_invoice_from_libelle(libelle)
        if extracted:
            return extracted
        return libelle

    if piece:
        return piece

    extracted = extract_invoice_from_libelle(libelle)
    if extracted:
        return extracted

    return ""


def find_supplier_boundaries_pennyland(df: pd.DataFrame) -> List[SupplierBoundary]:
    """Détecte les fournisseurs dans un GL Pennyland.
    Chaque groupe de lignes consécutives partageant le même N° de compte = un fournisseur.
    Le fichier doit être lu avec header=1 (la ligne 0 blank est sautée automatiquement).
    """
    suppliers = []
    current_account = None
    current_name = None
    start_row = None

    for idx in range(len(df)):
        account = normalize_text(df.iat[idx, PENNYLAND_COL_INDEX["account"]])
        name    = normalize_text(df.iat[idx, PENNYLAND_COL_INDEX["name"]])

        # Ignorer lignes vides / NaN
        if not account or account.lower() in ("nan", ""):
            if current_account is not None:
                suppliers.append(SupplierBoundary(
                    supplier_code=current_account,
                    supplier_name=current_name,
                    start_row=start_row,
                    end_row=idx - 1,
                ))
                current_account = None
                current_name = None
                start_row = None
            continue

        if account != current_account:
            if current_account is not None:
                suppliers.append(SupplierBoundary(
                    supplier_code=current_account,
                    supplier_name=current_name,
                    start_row=start_row,
                    end_row=idx - 1,
                ))
            current_account = account
            current_name    = name
            start_row       = idx

    # Dernier fournisseur
    if current_account is not None:
        suppliers.append(SupplierBoundary(
            supplier_code=current_account,
            supplier_name=current_name,
            start_row=start_row,
            end_row=len(df) - 1,
        ))

    return suppliers


def prepare_supplier_data_pennyland(df: pd.DataFrame, boundary: SupplierBoundary,
                                     row_offset: int = 0) -> pd.DataFrame:
    """Prépare les données d'un fournisseur depuis un GL Pennyland.
    Retourne un DataFrame aux mêmes colonnes normalisées que prepare_supplier_data (COALA).
    """
    supplier_df = df.iloc[boundary.start_row: boundary.end_row + 1].copy()
    supplier_df = supplier_df.reset_index(drop=False).rename(columns={"index": "original_row"})
    supplier_df["original_row"] = supplier_df["original_row"] + row_offset

    # +1 sur chaque index car reset_index() a inséré la colonne "original_row" en position 0
    ci = PENNYLAND_COL_INDEX

    result = pd.DataFrame()
    result["original_row"] = supplier_df["original_row"]

    # Identifiant fournisseur (N° de compte) dans LibelleColA pour compatibilité
    result["LibelleColA"]  = supplier_df.iloc[:, ci["account"] + 1].apply(normalize_text)
    result["DateOperation"] = supplier_df.iloc[:, ci["date"] + 1].apply(safe_date)
    result["JournalColC"]  = supplier_df.iloc[:, ci["journal"] + 1].apply(normalize_text)
    result["LibelleColE"]  = supplier_df.iloc[:, ci["libelle"] + 1].apply(normalize_text)
    result["Debit"]        = supplier_df.iloc[:, ci["debit"] + 1].apply(normalize_amount)
    result["Credit"]       = supplier_df.iloc[:, ci["credit"] + 1].apply(normalize_amount)
    result["Lettrage"]     = supplier_df.iloc[:, ci["lettrage"] + 1].apply(normalize_text)

    # N° de facture : col H directement ; fallback sur Libellé de pièce si vide
    raw_num = supplier_df.iloc[:, ci["num_facture"] + 1].apply(normalize_text).tolist()
    result["NumeroFacture"] = [
        n if n else lib
        for n, lib in zip(raw_num, result["LibelleColE"])
    ]

    return result


def find_supplier_boundaries_sage(df: pd.DataFrame) -> List[SupplierBoundary]:
    """Détecte les fournisseurs dans un GL SAGE.
    Chaque groupe de lignes consécutives partageant le même N° COMPTE (col A) = un fournisseur.
    Pas de nom fournisseur : chaque ligne garde son propre libellé / N° pièce.
    """
    suppliers = []
    current_account = None
    start_row = None

    for idx in range(len(df)):
        account = normalize_text(df.iat[idx, SAGE_COL_INDEX["compte"]])

        if not account or account.lower() in ("nan", ""):
            if current_account is not None:
                suppliers.append(SupplierBoundary(
                    supplier_code=current_account,
                    supplier_name="",
                    start_row=start_row,
                    end_row=idx - 1,
                ))
                current_account = None
                start_row = None
            continue

        if account != current_account:
            if current_account is not None:
                suppliers.append(SupplierBoundary(
                    supplier_code=current_account,
                    supplier_name="",
                    start_row=start_row,
                    end_row=idx - 1,
                ))
            current_account = account
            start_row = idx

    if current_account is not None:
        suppliers.append(SupplierBoundary(
            supplier_code=current_account,
            supplier_name="",
            start_row=start_row,
            end_row=len(df) - 1,
        ))

    return suppliers


def prepare_supplier_data_sage(df: pd.DataFrame, boundary: SupplierBoundary,
                                row_offset: int = 0) -> pd.DataFrame:
    """Prépare les données d'un fournisseur depuis un GL SAGE."""
    supplier_df = df.iloc[boundary.start_row: boundary.end_row + 1].copy()
    supplier_df = supplier_df.reset_index(drop=False).rename(columns={"index": "original_row"})
    supplier_df["original_row"] = supplier_df["original_row"] + row_offset

    ci = SAGE_COL_INDEX

    result = pd.DataFrame()
    result["original_row"]  = supplier_df["original_row"]
    result["LibelleColA"]   = supplier_df.iloc[:, ci["compte"] + 1].apply(normalize_text)
    result["DateOperation"] = supplier_df.iloc[:, ci["date"] + 1].apply(safe_date)
    result["JournalColC"]   = supplier_df.iloc[:, ci["journal"] + 1].apply(normalize_text)
    result["PieceColD"]     = supplier_df.iloc[:, ci["piece"] + 1].apply(normalize_text)
    result["LibelleColE"]   = supplier_df.iloc[:, ci["libelle"] + 1].apply(normalize_text)
    result["Lettrage"]      = supplier_df.iloc[:, ci["lettrage"] + 1].apply(normalize_text)
    result["Debit"]         = supplier_df.iloc[:, ci["debit"] + 1].apply(normalize_amount)
    result["Credit"]        = supplier_df.iloc[:, ci["credit"] + 1].apply(normalize_amount)

    # N° facture = N° de pièce (col D) ; fallback sur libellé si vide
    result["NumeroFacture"] = [
        p if p else lib
        for p, lib in zip(result["PieceColD"], result["LibelleColE"])
    ]

    return result


def find_supplier_boundaries(df: pd.DataFrame) -> List[SupplierBoundary]:
    suppliers = []
    current_supplier = None

    for idx in range(len(df)):
        cell_a = normalize_text(df.iat[idx, EXPECTED_COL_INDEX["col_a"]])

        supplier_info = extract_supplier_info(cell_a)
        if supplier_info:
            code, name = supplier_info
            current_supplier = (code, name, idx)
            continue

        if current_supplier and is_supplier_total(cell_a):
            code, name, start_row = current_supplier
            suppliers.append(
                SupplierBoundary(
                    supplier_code=code,
                    supplier_name=name,
                    start_row=start_row,
                    end_row=idx,
                )
            )
            current_supplier = None

    return suppliers


def prepare_supplier_data(df: pd.DataFrame, boundary: SupplierBoundary, row_offset: int = 0) -> pd.DataFrame:
    supplier_df = df.iloc[boundary.start_row + 1: boundary.end_row].copy()
    supplier_df = supplier_df.reset_index(drop=False).rename(columns={"index": "original_row"})
    supplier_df["original_row"] = supplier_df["original_row"] + row_offset

    supplier_df["DateOperation"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["date"] + 1].apply(safe_date)
    supplier_df["JournalColC"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["journal"] + 1].apply(normalize_text)
    supplier_df["PieceColD"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["piece"] + 1].apply(normalize_text)
    supplier_df["LibelleColE"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["libelle"] + 1].apply(normalize_text)
    supplier_df["Debit"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["debit"] + 1].apply(normalize_amount)
    supplier_df["Credit"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["credit"] + 1].apply(normalize_amount)
    supplier_df["Lettrage"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["lettrage"] + 1].apply(normalize_text)
    supplier_df["LibelleColA"] = supplier_df.iloc[:, EXPECTED_COL_INDEX["col_a"] + 1].apply(normalize_text)

    supplier_df["NumeroFacture"] = supplier_df.apply(
        lambda row: extract_invoice_number(row["PieceColD"], row["LibelleColE"]),
        axis=1
    )

    return supplier_df


def allocate_amount_match(invoice_rows: pd.DataFrame, payment_rows: pd.DataFrame):
    """Allocation en 2 passes (factures / paiements lettrés avec plusieurs lignes).

    Passe 1 — rapprochement par montant identique :
        Pour chaque facture, on cherche un paiement avec un montant restant
        strictement égal (tolérance 0.01). Si trouvé, appariement direct.

    Passe 2 — FIFO sur le reste (paiement le plus ancien × facture la plus ancienne) :
        Les factures / paiements non appariés en passe 1 sont triés par date
        croissante puis alloués en FIFO (un paiement peut couvrir plusieurs
        factures et vice-versa).
    """
    allocations = []
    residuals   = []

    if invoice_rows.empty:
        return allocations, residuals

    invoices = [
        {"row": inv, "remaining": round(inv["Credit"], 2), "total": round(inv["Credit"], 2)}
        for _, inv in invoice_rows.iterrows()
    ]
    payments = [
        {"row": pay, "remaining": round(pay["Debit"], 2)}
        for _, pay in payment_rows.iterrows()
    ]

    # ── Passe 1 : rapprochement par montant identique ───────────────
    for inv_obj in invoices:
        if inv_obj["remaining"] <= 0:
            continue
        for pay_obj in payments:
            if pay_obj["remaining"] <= 0:
                continue
            if abs(pay_obj["remaining"] - inv_obj["remaining"]) < 0.01:
                amount = inv_obj["remaining"]
                allocations.append({
                    "invoice_row":      inv_obj["row"],
                    "payment_row":      pay_obj["row"],
                    "allocated_amount": amount,
                    "invoice_total":    inv_obj["total"],
                })
                pay_obj["remaining"] = round(pay_obj["remaining"] - amount, 2)
                inv_obj["remaining"] = 0
                break

    # ── Passe 2 : FIFO sur le reste (trié par date) ─────────────────
    def _sort_key(obj):
        d = obj["row"]["DateOperation"]
        return (
            pd.Timestamp.max if pd.isna(d) else d,
            int(obj["row"].get("original_row", 0) or 0),
        )

    remaining_invoices = sorted(
        [i for i in invoices if i["remaining"] > 0.01], key=_sort_key
    )
    remaining_payments = sorted(
        [p for p in payments if p["remaining"] > 0.01], key=_sort_key
    )

    pay_idx = 0
    for inv_obj in remaining_invoices:
        while inv_obj["remaining"] > 0.01 and pay_idx < len(remaining_payments):
            pay_obj = remaining_payments[pay_idx]
            if pay_obj["remaining"] <= 0:
                pay_idx += 1
                continue
            allocated = round(min(inv_obj["remaining"], pay_obj["remaining"]), 2)
            if allocated <= 0:
                pay_idx += 1
                continue
            allocations.append({
                "invoice_row":      inv_obj["row"],
                "payment_row":      pay_obj["row"],
                "allocated_amount": allocated,
                "invoice_total":    inv_obj["total"],
            })
            inv_obj["remaining"] = round(inv_obj["remaining"] - allocated, 2)
            pay_obj["remaining"] = round(pay_obj["remaining"] - allocated, 2)
            if pay_obj["remaining"] <= 0:
                pay_idx += 1

        if inv_obj["remaining"] > 0.01:
            residuals.append({
                "invoice_row":      inv_obj["row"],
                "remaining_amount": inv_obj["remaining"],
                "invoice_total":    inv_obj["total"],
            })

    return allocations, residuals


def allocate_fifo(invoice_rows: pd.DataFrame, payment_rows: pd.DataFrame):
    allocations = []
    residuals = []

    if invoice_rows.empty:
        return allocations, residuals

    invoices = []
    for _, inv in invoice_rows.iterrows():
        invoices.append({
            "row": inv,
            "remaining": round(inv["Credit"], 2)
        })

    payments = []
    for _, pay in payment_rows.iterrows():
        payments.append({
            "row": pay,
            "remaining": round(pay["Debit"], 2)
        })

    pay_idx = 0

    for invoice in invoices:
        inv_row = invoice["row"]
        invoice_total = round(inv_row["Credit"], 2)

        while invoice["remaining"] > 0 and pay_idx < len(payments):
            while pay_idx < len(payments) and payments[pay_idx]["remaining"] <= 0:
                pay_idx += 1

            if pay_idx >= len(payments):
                break

            pay_obj = payments[pay_idx]
            allocated = round(min(invoice["remaining"], pay_obj["remaining"]), 2)

            if allocated <= 0:
                pay_idx += 1
                continue

            allocations.append({
                "invoice_row": inv_row,
                "payment_row": pay_obj["row"],
                "allocated_amount": allocated,
                "invoice_total": invoice_total,
            })

            invoice["remaining"] = round(invoice["remaining"] - allocated, 2)
            pay_obj["remaining"] = round(pay_obj["remaining"] - allocated, 2)

            if pay_obj["remaining"] <= 0:
                pay_idx += 1

        if invoice["remaining"] > 0:
            residuals.append({
                "invoice_row": inv_row,
                "remaining_amount": round(invoice["remaining"], 2),
                "invoice_total": invoice_total,
            })

    return allocations, residuals


def process_supplier(boundary: SupplierBoundary, supplier_df: pd.DataFrame,
                     reference_date: pd.Timestamp, opening_date: pd.Timestamp,
                     mode: str = "civile", year_filter: Optional[int] = None,
                     gl_format: str = "coala",
                     invoice_journals: Optional[List[str]] = None,
                     payment_journals: Optional[List[str]] = None):
    """
    gl_format : "coala"    → détection solde ouverture via LibelleColA, allocation FIFO
                "pennyland" → détection solde ouverture via JournalColC=="AA", allocation par montant
                "sage"      → journaux factures/paiements fournis par l'utilisateur
    invoice_journals / payment_journals : pour SAGE, listes des codes journaux.
    """
    paid_rows = []
    unpaid_rows = []

    has_source_file = "source_file" in supplier_df.columns

    # --- Masque solde d'ouverture (varie selon le format) ---
    # "mixed" = un fichier COALA + un fichier PENNYLAND → on combine les deux détections
    _AN_JOURNALS = {"AA", "AD", "AN"}
    journals_upper = supplier_df["JournalColC"].str.upper()
    if gl_format == "pennyland":
        opening_mask = journals_upper.isin(_AN_JOURNALS)
    elif gl_format == "sage":
        opening_mask = journals_upper.isin(_AN_JOURNALS)
    elif gl_format == "mixed":
        opening_mask = (
            journals_upper.isin(_AN_JOURNALS)
            | supplier_df["LibelleColA"].apply(is_opening_balance_row)
        )
    else:  # coala
        opening_mask = (
            supplier_df["LibelleColA"].apply(is_opening_balance_row)
            | journals_upper.isin(_AN_JOURNALS)
        )

    # --- Masques journaux (factures / paiements) ---
    if gl_format == "sage":
        inv_j = [j.strip().upper() for j in (invoice_journals or []) if j.strip()]
        pay_j = [j.strip().upper() for j in (payment_journals or []) if j.strip()]
        # Inclure AA/AD/AN pour 2025+ même si l'utilisateur ne les a pas listés
        inv_journal_mask = journals_upper.isin(inv_j) | journals_upper.isin(_AN_JOURNALS)
        pay_journal_mask = journals_upper.isin(pay_j) if pay_j else pd.Series(True, index=supplier_df.index)
    else:
        inv_journal_mask = journals_upper.str.startswith("A", na=False)
        pay_journal_mask = pd.Series(True, index=supplier_df.index)

    # Pour min_year+ : les à nouveau (AA/AD/AN) sont inclus comme factures (montant crédit)
    _min_y = _min_year()
    _is_min_plus = supplier_df["DateOperation"].dt.year >= _min_y
    invoice_excl_mask = opening_mask & ~_is_min_plus.fillna(False)

    # --- Fonction d'allocation lettrée (varie selon le format) ---
    # Pour "mixed" et "pennyland" : matching par montant (avec fallback FIFO intégré)
    _alloc_lettered = allocate_fifo if gl_format == "coala" else allocate_amount_match

    # ============================================================
    # 1) FACTURES LETTRÉES / PAIEMENTS LETTRÉS
    # ============================================================
    def _invoice_year_ok(s):
        if year_filter is None:
            return s.dt.year >= _min_y
        return (s.dt.year >= _min_y) & (s.dt.year <= year_filter)

    invoices_lettered = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] != "")
        & inv_journal_mask
        & (~invoice_excl_mask)                             # exclure à nouveau sauf pour 2025+
        & _invoice_year_ok(supplier_df["DateOperation"])
        & (supplier_df["DateOperation"].isna() | (supplier_df["DateOperation"] <= reference_date))
    ].copy()

    payments_lettered = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] != "")
        & pay_journal_mask
    ].copy()

    if not invoices_lettered.empty:
        for letter in sorted(invoices_lettered["Lettrage"].dropna().unique()):
            inv_group = invoices_lettered[invoices_lettered["Lettrage"] == letter].copy()
            pay_group = payments_lettered[payments_lettered["Lettrage"] == letter].copy()

            if pay_group.empty:
                continue

            # ---- MODE À CHEVAL : gestion des paiements cross-année ----
            if mode == "cheval" and has_source_file:
                inv_files = set(inv_group["source_file"].unique())
                pay_group_same = pay_group[pay_group["source_file"].isin(inv_files)].copy()

                if pay_group_same.empty:
                    # Pas de paiement de la même année que la facture
                    # Vérifier s'il existe un paiement de l'année N-1 (fichier antérieur)
                    min_inv_file = min(inv_files)
                    pay_group_prev = pay_group[pay_group["source_file"] < min_inv_file].copy()

                    if not pay_group_prev.empty:
                        # Paiement de l'année N-1 : facture payée EN AVANCE
                        # Le délai sera négatif → classify_remark → "Rien à signaler"
                        pay_group = pay_group_prev
                        # On continue vers l'allocation FIFO normale ci-dessous
                    else:
                        # Aucun paiement trouvé dans les fichiers N-1 et N
                        # → le paiement est de l'année N+1 : facture non payée à la clôture
                        for _, inv in inv_group.iterrows():
                            invoice_date = inv["DateOperation"]
                            invoice_total = round(inv["Credit"], 2)
                            delay_days = (
                                None if pd.isna(invoice_date)
                                else int((reference_date - invoice_date).days + 1)
                            )
                            unpaid_rows.append({
                                "Code fournisseur": boundary.supplier_code,
                                "Nom fournisseur": boundary.supplier_name,
                                "Libellé": inv["LibelleColE"],
                                "Numero facture": inv["NumeroFacture"],
                                "Type": "Facture non payée (paiement année suivante)",
                                "Journal": inv["JournalColC"],
                                "Lettrage": letter,
                                "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                                "Date paiement": None,
                                "Montant facture": invoice_total,
                                "Montant(s) paiement(s) total": 0.0,
                                "Montant facture origine": invoice_total,
                                "Délai (jours)": delay_days,
                                "Remarque": classify_remark(delay_days),
                                "Ligne source facture": int(inv["original_row"]) + 2,
                            })
                        continue
                else:
                    pay_group = pay_group_same
            # ---- FIN MODE À CHEVAL ----

            inv_group = inv_group.sort_values(by=["DateOperation", "original_row"], na_position="last")
            pay_group = pay_group.sort_values(by=["DateOperation", "original_row"], na_position="last")

            allocations, _ = _alloc_lettered(inv_group, pay_group)

            for alloc in allocations:
                inv = alloc["invoice_row"]
                pay = alloc["payment_row"]
                allocated_amount = alloc["allocated_amount"]
                invoice_total = alloc["invoice_total"]

                invoice_date = inv["DateOperation"]
                payment_date = pay["DateOperation"]

                # ── Paiement APRÈS la date de clôture (31/12/N) ──────────────────
                # La facture est considérée NON PAYÉE à la clôture.
                # Délai = 31/12/N − date_facture + 1
                if pd.notna(payment_date) and payment_date > reference_date:
                    clot_delay = (
                        None if pd.isna(invoice_date)
                        else int((reference_date - invoice_date).days + 1)
                    )
                    unpaid_rows.append({
                        "Code fournisseur": boundary.supplier_code,
                        "Nom fournisseur": boundary.supplier_name,
                        "Libellé": inv["LibelleColE"],
                        "Numero facture": inv["NumeroFacture"],
                        "Type": "Facture non payée (paiement après clôture)",
                        "Journal": inv["JournalColC"],
                        "Lettrage": letter,
                        "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                        "Date paiement": None,
                        "Montant facture": round(allocated_amount, 2),
                        "Montant(s) paiement(s) total": 0.0,
                        "Montant facture origine": round(invoice_total, 2),
                        "Délai (jours)": clot_delay,
                        "Remarque": classify_remark(clot_delay),
                        "Ligne source facture": int(inv["original_row"]) + 2,
                    })
                    continue  # ne pas ajouter aux paid_rows
                # ── Paiement dans l'année (ou date invalide) ─────────────────────
                if pd.isna(invoice_date) or pd.isna(payment_date):
                    delay_days = None
                    remark = "Date facture ou date paiement invalide"
                else:
                    delay_days = int((payment_date - invoice_date).days + 1)
                    remark = classify_remark(delay_days)

                paid_rows.append({
                    "Code fournisseur": boundary.supplier_code,
                    "Nom fournisseur": boundary.supplier_name,
                    "Libellé": inv["LibelleColE"],
                    "Numero facture": inv["NumeroFacture"],
                    "Type": "Facture payée",
                    "Journal": inv["JournalColC"],
                    "Lettrage": letter,
                    "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                    "Date paiement": payment_date.date() if pd.notna(payment_date) else None,
                    "Montant facture": round(allocated_amount, 2),
                    "Montant(s) paiement(s) total": round(allocated_amount, 2),
                    "Montant facture origine": round(invoice_total, 2),
                    "Délai (jours)": delay_days,
                    "Remarque": remark,
                    "Ligne source facture": int(inv["original_row"]) + 2,
                })

    # ============================================================
    # 2) FACTURES NON LETTRÉES + AVANCES + PAIEMENTS NON LETTRÉS (FIFO)
    # ============================================================
    unpaid_invoices = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] == "")
        & inv_journal_mask
        & (~invoice_excl_mask)
        & _invoice_year_ok(supplier_df["DateOperation"])
        & (supplier_df["DateOperation"].isna() | (supplier_df["DateOperation"] <= reference_date))
    ].copy()

    unpaid_invoices = unpaid_invoices.sort_values(by=["DateOperation", "original_row"], na_position="last")

    opening_rows = supplier_df[opening_mask].copy()
    opening_advance_amount = round(opening_rows["Debit"].sum(), 2)

    unlettered_payments = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] == "")
        & (~opening_mask)
        & pay_journal_mask
        # Exclure les paiements postérieurs à la date de clôture
        & (supplier_df["DateOperation"].isna() | (supplier_df["DateOperation"] <= reference_date))
    ].copy()

    unlettered_payments = unlettered_payments.sort_values(by=["DateOperation", "original_row"], na_position="last")

    # 2A) Imputation de l'avance d'ouverture
    if not unpaid_invoices.empty and opening_advance_amount > 0:
        pseudo_opening_payment = pd.DataFrame([{
            "DateOperation": opening_date,
            "Debit": opening_advance_amount,
            "original_row": -1
        }])

        opening_allocations, opening_residuals = allocate_fifo(unpaid_invoices, pseudo_opening_payment)

        consumed_invoice_ids = set()

        for alloc in opening_allocations:
            inv = alloc["invoice_row"]
            allocated_amount = alloc["allocated_amount"]
            invoice_total = alloc["invoice_total"]
            invoice_date = inv["DateOperation"]

            if pd.isna(invoice_date):
                delay_days = None
            else:
                delay_days = int((opening_date - invoice_date).days + 1)

            unpaid_rows.append({
                "Code fournisseur": boundary.supplier_code,
                "Nom fournisseur": boundary.supplier_name,
                "Libellé": inv["LibelleColE"],
                "Numero facture": inv["NumeroFacture"],
                "Type": "Avance imputée",
                "Journal": inv["JournalColC"],
                "Lettrage": "",
                "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                "Date paiement": opening_date.date(),
                "Montant facture": round(allocated_amount, 2),
                "Montant(s) paiement(s) total": round(allocated_amount, 2),
                "Montant facture origine": round(invoice_total, 2),
                "Délai (jours)": delay_days,
                "Remarque": classify_remark(delay_days),
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

            consumed_invoice_ids.add(int(inv["original_row"]))

        residual_map = {}

        for res in opening_residuals:
            inv = res["invoice_row"]
            residual_map[int(inv["original_row"])] = round(res["remaining_amount"], 2)

        for _, inv in unpaid_invoices.iterrows():
            key = int(inv["original_row"])
            if key not in residual_map and key not in consumed_invoice_ids:
                residual_map[key] = round(inv["Credit"], 2)

        unpaid_working = unpaid_invoices.copy()
        unpaid_working["RemainingAfterOpening"] = unpaid_working["original_row"].apply(
            lambda x: residual_map.get(int(x), 0.0)
        )
        unpaid_working = unpaid_working[unpaid_working["RemainingAfterOpening"] > 0].copy()
        unpaid_working["Credit"] = unpaid_working["RemainingAfterOpening"]
        unpaid_working = unpaid_working.drop(columns=["RemainingAfterOpening"])
    else:
        unpaid_working = unpaid_invoices.copy()

    # 2B) Paiements non lettrés => 1) match montant identique, 2) FIFO sur le reste
    if not unpaid_working.empty and not unlettered_payments.empty:
        fifo_allocations, fifo_residuals = allocate_amount_match(unpaid_working, unlettered_payments)

        for alloc in fifo_allocations:
            inv = alloc["invoice_row"]
            pay = alloc["payment_row"]
            allocated_amount = alloc["allocated_amount"]
            invoice_total = alloc["invoice_total"]

            invoice_date = inv["DateOperation"]
            payment_date = pay["DateOperation"]

            if pd.isna(invoice_date) or pd.isna(payment_date):
                delay_days = None
                remark = "Date facture ou date paiement invalide"
            else:
                delay_days = int((payment_date - invoice_date).days + 1)
                remark = classify_remark(delay_days)

            unpaid_rows.append({
                "Code fournisseur": boundary.supplier_code,
                "Nom fournisseur": boundary.supplier_name,
                "Libellé": inv["LibelleColE"],
                "Numero facture": inv["NumeroFacture"],
                "Type": "Paiement non lettré imputé FIFO",
                "Journal": inv["JournalColC"],
                "Lettrage": "",
                "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                "Date paiement": payment_date.date() if pd.notna(payment_date) else None,
                "Montant facture": round(allocated_amount, 2),
                "Montant(s) paiement(s) total": round(allocated_amount, 2),
                "Montant facture origine": round(invoice_total, 2),
                "Délai (jours)": delay_days,
                "Remarque": remark,
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

        for res in fifo_residuals:
            inv = res["invoice_row"]
            remaining_amount = round(res["remaining_amount"], 2)
            invoice_total = round(res["invoice_total"], 2)
            invoice_date = inv["DateOperation"]

            if pd.isna(invoice_date):
                delay_days = None
            else:
                delay_days = int((reference_date - invoice_date).days + 1)

            unpaid_rows.append({
                "Code fournisseur": boundary.supplier_code,
                "Nom fournisseur": boundary.supplier_name,
                "Libellé": inv["LibelleColE"],
                "Numero facture": inv["NumeroFacture"],
                "Type": "Facture non payée",
                "Journal": inv["JournalColC"],
                "Lettrage": "",
                "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                "Date paiement": None,
                "Montant facture": round(remaining_amount, 2),
                "Montant(s) paiement(s) total": 0.0,
                "Montant facture origine": invoice_total,
                "Délai (jours)": delay_days,
                "Remarque": classify_remark(delay_days),
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

    else:
        for _, inv in unpaid_working.iterrows():
            invoice_date = inv["DateOperation"]
            remaining_amount = round(inv["Credit"], 2)

            if remaining_amount <= 0:
                continue

            if pd.isna(invoice_date):
                delay_days = None
            else:
                delay_days = int((reference_date - invoice_date).days + 1)

            unpaid_rows.append({
                "Code fournisseur": boundary.supplier_code,
                "Nom fournisseur": boundary.supplier_name,
                "Libellé": inv["LibelleColE"],
                "Numero facture": inv["NumeroFacture"],
                "Type": "Facture non payée",
                "Journal": inv["JournalColC"],
                "Lettrage": "",
                "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                "Date paiement": None,
                "Montant facture": round(remaining_amount, 2),
                "Montant(s) paiement(s) total": 0.0,
                "Montant facture origine": round(inv["Credit"], 2),
                "Délai (jours)": delay_days,
                "Remarque": classify_remark(delay_days),
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

    return paid_rows, unpaid_rows


def _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df,
                        year_filter: int = None, gl_format: str = "coala",
                        reference_date: pd.Timestamp = None,
                        invoice_journals: Optional[List[str]] = None):
    """Calcule la ligne de contrôle pour un fournisseur.
    year_filter    : si fourni, on ne comptabilise que les crédits dont la DateOperation est dans cette année.
    gl_format      : détection du solde d'ouverture adaptée au format.
    reference_date : si fourni, on exclut les crédits postérieurs à cette date (utile en mode trimestriel).
    invoice_journals : pour SAGE, liste des codes journaux facture.
    """
    _AN_JOURNALS = {"AA", "AD", "AN"}
    journals_upper = supplier_df["JournalColC"].str.upper()
    if gl_format == "pennyland":
        opening_excl = journals_upper.isin(_AN_JOURNALS)
    elif gl_format == "sage":
        opening_excl = journals_upper.isin(_AN_JOURNALS)
    elif gl_format == "mixed":
        opening_excl = (
            journals_upper.isin(_AN_JOURNALS)
            | supplier_df["LibelleColE"].apply(is_opening_balance_row)
        )
    else:  # coala
        opening_excl = (
            supplier_df["LibelleColE"].apply(is_opening_balance_row)
            | journals_upper.isin(_AN_JOURNALS)
        )

    # Les à nouveau (AA/AD/AN) sont inclus si leur date est >= min_year
    _min_y = _min_year()
    _is_min_plus = supplier_df["DateOperation"].dt.year >= _min_y
    opening_excl = opening_excl & ~_is_min_plus.fillna(False)

    # Masque journaux factures (SAGE uniquement)
    if gl_format == "sage":
        inv_j = [j.strip().upper() for j in (invoice_journals or []) if j.strip()]
        inv_journal_mask = journals_upper.isin(inv_j) | journals_upper.isin(_AN_JOURNALS)
    else:
        inv_journal_mask = pd.Series(True, index=supplier_df.index)

    mask = (
        (supplier_df["Credit"] > 0)
        & (~opening_excl)
        & inv_journal_mask
        & (supplier_df["DateOperation"].dt.year >= _min_y)
    )
    if year_filter is not None:
        mask = mask & (supplier_df["DateOperation"].dt.year <= year_filter)
    if reference_date is not None:
        mask = mask & (supplier_df["DateOperation"].isna() | (supplier_df["DateOperation"] <= reference_date))

    total_credit_source = round(supplier_df.loc[mask, "Credit"].sum(), 2)

    # Total débit (paiements) — pour détecter les fournisseurs "que paiements"
    # ou "paiements > factures". On exclut les à-nouveau (opening_excl) et,
    # pour SAGE, on se limite aux journaux de paiement déclarés par l'utilisateur.
    if gl_format == "sage":
        # supplier_df doit avoir été construit avec le même ctx, mais on ne
        # connaît pas ici les payment_journals → on prend tous les débits hors A-nouveau.
        pay_mask = (supplier_df["Debit"] > 0) & (~opening_excl)
    else:
        pay_mask = (supplier_df["Debit"] > 0) & (~opening_excl)
    total_debit_source = round(supplier_df.loc[pay_mask, "Debit"].sum(), 2)

    # Flag "que paiements" :
    #   • soit aucune facture retenue mais des paiements existent
    #   • soit les paiements excèdent strictement les factures retenues
    payments_only = (
        (total_debit_source > 0.01 and total_credit_source < 0.01)
        or (total_debit_source > total_credit_source + 0.01)
    )

    all_rows = paid_rows + unpaid_rows
    factures_reconstituees = {}
    for r in all_rows:
        key = (r.get("Code fournisseur"), r.get("Numero facture"))
        factures_reconstituees[key] = factures_reconstituees.get(key, 0.0) + float(r.get("Montant facture", 0) or 0)

    total_factures_retenues = round(sum(factures_reconstituees.values()), 2)
    ecart = round(total_credit_source - total_factures_retenues, 2)

    return {
        "Code fournisseur": boundary.supplier_code,
        "Nom fournisseur": boundary.supplier_name,
        "Nb lignes payées": len(paid_rows),
        "Nb lignes non payées": len(unpaid_rows),
        "Total crédit source": total_credit_source,
        "Total débit source": total_debit_source,
        "Total factures retenues": total_factures_retenues,
        "Écart": ecart,
        "Statut contrôle": "OK" if abs(ecart) < 0.01 else "Écart à analyser",
        "_payments_only": payments_only,
    }


def _filter_by_year(rows: list, valid_years) -> list:
    """Garde uniquement les lignes dont la date facture est dans les années valides.
    valid_years peut être :
      • un int (ex: 2026) → interprété comme "de 2025 à cette année incluse"
      • un set d'années → utilisé tel quel
    Les lignes sans date (None) sont conservées.
    """
    if isinstance(valid_years, int):
        _min_y = _min_year()
        valid_years = set(range(_min_y, valid_years + 1)) if valid_years >= _min_y else {valid_years}
    result = []
    for r in rows:
        d = r.get("Date facture")
        if d is None or d.year in valid_years:
            result.append(r)
    return result


def process_workbook(uploaded_file, sheet_name, reference_date: pd.Timestamp, opening_date: pd.Timestamp):
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)

    if df.shape[1] < 8:
        raise ValueError("Le fichier doit contenir au minimum 8 colonnes (A à H).")

    suppliers = find_supplier_boundaries(df)
    if not suppliers:
        raise ValueError("Aucun fournisseur détecté dans la colonne A.")

    all_paid = []
    all_unpaid = []
    control_rows = []
    year = reference_date.year

    for boundary in suppliers:
        supplier_df = prepare_supplier_data(df, boundary)
        paid_rows, unpaid_rows = process_supplier(boundary, supplier_df, reference_date, opening_date, mode="civile", year_filter=year)

        # Filtre final : on ne garde que les factures de l'année choisie
        paid_rows = _filter_by_year(paid_rows, year)
        unpaid_rows = _filter_by_year(unpaid_rows, year)

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(_build_control_row(boundary, paid_rows, unpaid_rows, supplier_df, reference_date=reference_date))

    return _finalize_results(all_paid, all_unpaid, control_rows)


def process_workbook_cheval(file1, sheet1, file2, sheet2,
                            reference_date: pd.Timestamp, opening_date: pd.Timestamp):
    df1 = pd.read_excel(file1, sheet_name=sheet1, header=0)
    df2 = pd.read_excel(file2, sheet_name=sheet2, header=0)

    if df1.shape[1] < 8 or df2.shape[1] < 8:
        raise ValueError("Les deux fichiers doivent contenir au minimum 8 colonnes (A à H).")

    suppliers1 = find_supplier_boundaries(df1)
    suppliers2 = find_supplier_boundaries(df2)

    if not suppliers1 and not suppliers2:
        raise ValueError("Aucun fournisseur détecté dans les deux fichiers.")

    # Index par code fournisseur
    map1 = {b.supplier_code: b for b in suppliers1}
    map2 = {b.supplier_code: b for b in suppliers2}
    all_codes = sorted(set(map1.keys()) | set(map2.keys()))

    # Offset pour éviter collision d'index entre les deux fichiers
    row_offset_file2 = len(df1) + 10000

    all_paid = []
    all_unpaid = []
    control_rows = []

    # Filtre : uniquement les factures de l'année choisie (N)
    year = reference_date.year
    _min_y = _min_year()
    valid_years = set(range(_min_y, year + 1)) if year >= _min_y else {year}

    for code in all_codes:
        b1 = map1.get(code)
        b2 = map2.get(code)

        # Récupérer le nom fournisseur (préférence au fichier 1)
        boundary = b1 if b1 else b2

        parts = []
        if b1:
            df_s1 = prepare_supplier_data(df1, b1, row_offset=0)
            df_s1["source_file"] = 1
            parts.append(df_s1)
        if b2:
            df_s2 = prepare_supplier_data(df2, b2, row_offset=row_offset_file2)
            df_s2["source_file"] = 2
            parts.append(df_s2)

        supplier_df = pd.concat(parts, ignore_index=True) if len(parts) > 1 else parts[0]

        paid_rows, unpaid_rows = process_supplier(boundary, supplier_df, reference_date, opening_date, mode="cheval", year_filter=None)

        # Filtre final : uniquement les factures de l'année choisie
        paid_rows = _filter_by_year(paid_rows, valid_years)
        unpaid_rows = _filter_by_year(unpaid_rows, valid_years)

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(_build_control_row(boundary, paid_rows, unpaid_rows, supplier_df, year_filter=year, reference_date=reference_date))

    return _finalize_results(all_paid, all_unpaid, control_rows)


def process_workbook_pennyland(uploaded_file, sheet_name,
                                reference_date: pd.Timestamp, opening_date: pd.Timestamp):
    """Traitement GL Pennyland – mode Année civile."""
    # header=1 : saute la 1ère ligne vide de l'export Pennyland et utilise la 2ème comme en-tête
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)

    if df.shape[1] < 12:
        raise ValueError("Le fichier GL Pennyland doit contenir au minimum 12 colonnes.")

    suppliers = find_supplier_boundaries_pennyland(df)
    if not suppliers:
        raise ValueError("Aucun fournisseur détecté dans le fichier GL Pennyland.")

    all_paid     = []
    all_unpaid   = []
    control_rows = []
    year         = reference_date.year

    for boundary in suppliers:
        supplier_df = prepare_supplier_data_pennyland(df, boundary)
        paid_rows, unpaid_rows = process_supplier(
            boundary, supplier_df, reference_date, opening_date,
            mode="civile", year_filter=year, gl_format="pennyland"
        )
        paid_rows   = _filter_by_year(paid_rows,   year)
        unpaid_rows = _filter_by_year(unpaid_rows, year)

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(
            _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df,
                               year_filter=year, gl_format="pennyland",
                               reference_date=reference_date)
        )

    return _finalize_results(all_paid, all_unpaid, control_rows)


def process_workbook_sage(uploaded_file, sheet_name,
                           reference_date: pd.Timestamp, opening_date: pd.Timestamp,
                           invoice_journals: List[str], payment_journals: List[str]):
    """Traitement GL SAGE – mode Année civile."""
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)

    if df.shape[1] < 8:
        raise ValueError("Le fichier GL SAGE doit contenir au minimum 8 colonnes (A à H).")

    suppliers = find_supplier_boundaries_sage(df)
    if not suppliers:
        raise ValueError("Aucun fournisseur détecté dans le fichier GL SAGE.")

    all_paid, all_unpaid, control_rows = [], [], []
    year = reference_date.year

    for boundary in suppliers:
        supplier_df = prepare_supplier_data_sage(df, boundary)
        paid_rows, unpaid_rows = process_supplier(
            boundary, supplier_df, reference_date, opening_date,
            mode="civile", year_filter=year, gl_format="sage",
            invoice_journals=invoice_journals, payment_journals=payment_journals
        )
        paid_rows   = _filter_by_year(paid_rows,   year)
        unpaid_rows = _filter_by_year(unpaid_rows, year)
        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(
            _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df,
                               year_filter=year, gl_format="sage",
                               reference_date=reference_date,
                               invoice_journals=invoice_journals)
        )

    return _finalize_results(all_paid, all_unpaid, control_rows)


def _read_file_by_format(uploaded_file, sheet_name: str, gl_format: str):
    """Lit un fichier Excel selon le format GL et retourne (df, find_boundaries_fn, prepare_fn)."""
    if gl_format == "pennyland":
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)
        if df.shape[1] < 12:
            raise ValueError("Le fichier GL Pennyland doit contenir au minimum 12 colonnes.")
        return df, find_supplier_boundaries_pennyland, prepare_supplier_data_pennyland
    elif gl_format == "sage":
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)
        if df.shape[1] < 8:
            raise ValueError("Le fichier GL SAGE doit contenir au minimum 8 colonnes (A à H).")
        return df, find_supplier_boundaries_sage, prepare_supplier_data_sage
    else:  # coala
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
        if df.shape[1] < 8:
            raise ValueError("Le fichier GL COALA doit contenir au minimum 8 colonnes (A à H).")
        return df, find_supplier_boundaries, prepare_supplier_data


def _process_one_file_cheval(df, find_fn, prep_fn, gl_fmt: str,
                              reference_date: pd.Timestamp, opening_date: pd.Timestamp,
                              row_offset: int = 0,
                              invoice_journals: Optional[List[str]] = None,
                              payment_journals: Optional[List[str]] = None):
    """Traite un seul fichier en mode cheval (année civile avec filtre + règle paiement post-clôture).
    Utilisé pour le mode MIXED où les deux fichiers ont des formats différents.
    """
    suppliers = find_fn(df)
    year       = reference_date.year
    all_paid, all_unpaid, ctrl = [], [], []

    for boundary in suppliers:
        supplier_df = prep_fn(df, boundary, row_offset=row_offset)
        paid_rows, unpaid_rows = process_supplier(
            boundary, supplier_df, reference_date, opening_date,
            mode="civile", year_filter=year, gl_format=gl_fmt,
            invoice_journals=invoice_journals, payment_journals=payment_journals
        )
        paid_rows   = _filter_by_year(paid_rows,   year)
        unpaid_rows = _filter_by_year(unpaid_rows, year)
        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        ctrl.append(
            _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df,
                               year_filter=year, gl_format=gl_fmt,
                               reference_date=reference_date)
        )
    return all_paid, all_unpaid, ctrl


def process_workbook_cheval_generic(file1, sheet1, gl_format1: str,
                                     file2, sheet2, gl_format2: str,
                                     reference_date: pd.Timestamp, opening_date: pd.Timestamp,
                                     invoice_journals1: Optional[List[str]] = None,
                                     payment_journals1: Optional[List[str]] = None,
                                     invoice_journals2: Optional[List[str]] = None,
                                     payment_journals2: Optional[List[str]] = None):
    """Traitement Année à cheval générique : chaque fichier peut être COALA ou PENNYLAND.

    • Même format (COALA+COALA ou PENNYLAND+PENNYLAND) :
      appariement fournisseurs par code, détection croisée lettrage inter-années.

    • Formats différents (COALA+PENNYLAND ou PENNYLAND+COALA) :
      les systèmes de lettrage sont incompatibles → chaque fichier est traité
      indépendamment avec filtre sur l'année choisie. Les résultats sont fusionnés.
    """
    df1, find_fn1, prep_fn1 = _read_file_by_format(file1, sheet1, gl_format1)
    df2, find_fn2, prep_fn2 = _read_file_by_format(file2, sheet2, gl_format2)

    # ── Formats identiques : appariement inter-années ────────────────────────
    if gl_format1 == gl_format2:
        suppliers1 = find_fn1(df1)
        suppliers2 = find_fn2(df2)

        if not suppliers1 and not suppliers2:
            raise ValueError("Aucun fournisseur détecté dans les deux fichiers.")

        map1 = {b.supplier_code: b for b in suppliers1}
        map2 = {b.supplier_code: b for b in suppliers2}
        all_codes = sorted(set(map1.keys()) | set(map2.keys()))

        row_offset_file2 = len(df1) + 10000
        combined_gl_format = gl_format1

        all_paid, all_unpaid, control_rows = [], [], []
        year        = reference_date.year
        _min_y = _min_year()
        valid_years = set(range(_min_y, year + 1)) if year >= _min_y else {year}

        for code in all_codes:
            b1 = map1.get(code)
            b2 = map2.get(code)
            boundary = b1 if b1 else b2

            parts = []
            if b1:
                df_s1 = prep_fn1(df1, b1, row_offset=0)
                df_s1["source_file"] = 1
                parts.append(df_s1)
            if b2:
                df_s2 = prep_fn2(df2, b2, row_offset=row_offset_file2)
                df_s2["source_file"] = 2
                parts.append(df_s2)

            supplier_df = pd.concat(parts, ignore_index=True) if len(parts) > 1 else parts[0]

            paid_rows, unpaid_rows = process_supplier(
                boundary, supplier_df, reference_date, opening_date,
                mode="cheval", year_filter=None, gl_format=combined_gl_format,
                invoice_journals=invoice_journals1, payment_journals=payment_journals1
            )
            paid_rows   = _filter_by_year(paid_rows,   valid_years)
            unpaid_rows = _filter_by_year(unpaid_rows, valid_years)

            all_paid.extend(paid_rows)
            all_unpaid.extend(unpaid_rows)
            control_rows.append(
                _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df,
                                   year_filter=year, gl_format=combined_gl_format,
                                   reference_date=reference_date,
                                   invoice_journals=invoice_journals1)
            )

        return _finalize_results(all_paid, all_unpaid, control_rows)

    # ── Formats différents (MIXED) : traitement indépendant par fichier ──────
    # Les lettrages COALA et PENNYLAND sont dans des systèmes distincts et ne
    # peuvent pas être mis en correspondance inter-fichiers.
    # Chaque fichier est traité séparément en mode civile + filtre année N.
    paid1, unpaid1, ctrl1 = _process_one_file_cheval(
        df1, find_fn1, prep_fn1, gl_format1, reference_date, opening_date, row_offset=0,
        invoice_journals=invoice_journals1, payment_journals=payment_journals1
    )
    paid2, unpaid2, ctrl2 = _process_one_file_cheval(
        df2, find_fn2, prep_fn2, gl_format2, reference_date, opening_date,
        row_offset=len(df1) + 10000,
        invoice_journals=invoice_journals2, payment_journals=payment_journals2
    )

    if not paid1 and not unpaid1 and not paid2 and not unpaid2:
        raise ValueError("Aucune facture de l'année choisie trouvée dans les deux fichiers.")

    return _finalize_results(paid1 + paid2, unpaid1 + unpaid2, ctrl1 + ctrl2)


def _finalize_results(all_paid, all_unpaid, control_rows):
    # --- Sépare les fournisseurs "que paiements / paiements > factures" -----
    payments_only_codes = {
        c.get("Code fournisseur")
        for c in control_rows
        if c.get("_payments_only")
    }
    payments_only_rows = [
        c for c in control_rows if c.get("_payments_only")
    ]
    normal_control_rows = [
        c for c in control_rows if not c.get("_payments_only")
    ]

    def _split(rows):
        in_po, out_po = [], []
        for r in rows:
            (in_po if r.get("Code fournisseur") in payments_only_codes else out_po).append(r)
        return out_po, in_po

    all_paid,    paid_po    = _split(all_paid)
    all_unpaid,  unpaid_po  = _split(all_unpaid)

    # Nettoie la clé interne avant export
    for c in normal_control_rows + payments_only_rows:
        c.pop("_payments_only", None)

    paid_df    = pd.DataFrame(all_paid)
    unpaid_df  = pd.DataFrame(all_unpaid)
    control_df = pd.DataFrame(normal_control_rows)

    payments_only_detail_df = pd.DataFrame(paid_po + unpaid_po)
    payments_only_control_df = pd.DataFrame(payments_only_rows)

    result_df = pd.concat([paid_df, unpaid_df], ignore_index=True)
    if not result_df.empty:
        result_df = result_df.sort_values(
            by=["Code fournisseur", "Nom fournisseur", "Date facture", "Type"],
            na_position="last"
        ).reset_index(drop=True)

    return (result_df, paid_df, unpaid_df, control_df,
            payments_only_detail_df, payments_only_control_df)


def to_excel_bytes(result_df: pd.DataFrame, paid_df: pd.DataFrame,
                   unpaid_df: pd.DataFrame, control_df: pd.DataFrame,
                   payments_only_detail_df: pd.DataFrame = None,
                   payments_only_control_df: pd.DataFrame = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Resultat global")
        paid_df.to_excel(writer, index=False, sheet_name="Factures payees")
        unpaid_df.to_excel(writer, index=False, sheet_name="Factures non payees")
        control_df.to_excel(writer, index=False, sheet_name="Controle fournisseurs")
        if payments_only_detail_df is not None and not payments_only_detail_df.empty:
            payments_only_detail_df.to_excel(
                writer, index=False, sheet_name="Frs paiements seuls"
            )
        if payments_only_control_df is not None and not payments_only_control_df.empty:
            payments_only_control_df.to_excel(
                writer, index=False, sheet_name="Controle paiements seuls"
            )

        for sheet in writer.book.worksheets:
            for column_cells in sheet.columns:
                max_length = 0
                col_letter = column_cells[0].column_letter
                for cell in column_cells:
                    cell_value = "" if cell.value is None else str(cell.value)
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)
                sheet.column_dimensions[col_letter].width = min(max_length + 2, 45)

    output.seek(0)
    return output.getvalue()


# ============================================================
# INTERFACE UTILISATEUR
# ============================================================

with st.expander("Format attendu du fichier Excel", expanded=False):
    tab_coala, tab_pennyland, tab_sage = st.tabs(["GL COALA", "GL PENNYLAND", "GL SAGE"])
    with tab_coala:
        st.markdown(
            """
- **Colonne A** : début/fin fournisseur (ligne avec code « 9XXXX Nom »)
- **Colonne B** : date
- **Colonne C** : journal
- **Colonne D** : pièce
- **Colonne E** : libellé
- **Colonne F** : débit
- **Colonne G** : crédit
- **Colonne H** : lettrage

**Règles :** Factures = journal col C commence par `A` · Paiements non lettrés imputés en FIFO
"""
        )
    with tab_pennyland:
        st.markdown(
            """
- **Colonne A (N° de compte)** : code fournisseur – chaque compte = un fournisseur
- **Colonne B (Libellé de compte)** : nom du fournisseur
- **Colonne C (Date)** : date de la transaction
- **Colonne D (Journal)** : `AC` = facture · `BQ*` = paiement · `AA` = Report à nouveau (exclu)
- **Colonne E (Libellé de pièce)** : libellé
- **Colonne H (N° de facture)** : numéro de facture direct
- **Colonne J (Let.)** : lettrage
- **Colonne K (Débit)** / **Colonne L (Crédit)**

**Règles :** Factures = Crédit col L + journal commence par `A` (hors `AA`) ·
Paiements lettrés → matching par montant identique · Paiements non lettrés → FIFO
"""
        )
    with tab_sage:
        st.markdown(
            """
- **Colonne A (COMPTE)** : code fournisseur – chaque N° = un fournisseur
- **Colonne B (DATE)** : date de la facture
- **Colonne C (CODE J)** : journal (factures et paiements à configurer)
- **Colonne D (N° DE PIECE)** : numéro de pièce (= numéro de facture)
- **Colonne E (LIBELLE)** : libellé
- **Colonne F (LETTRAGE)** : lettrage
- **Colonne G (DEBIT)** : paiement ou avoir
- **Colonne H (CREDIT)** : facture

**Règles :** L'utilisateur précise les journaux factures et paiements (ex: `OD` ou `OD,VE`) ·
Paiements lettrés → matching par montant identique · Paiements non lettrés → FIFO ·
`AA`/`AD`/`AN` inclus comme factures si date ≥ 2025
"""
        )

st.divider()

# --- Sélection de l'année et du mode ---
col_y, col_m = st.columns([1, 2])
with col_y:
    selected_year = st.number_input(
        "Année d'exercice", min_value=2000, max_value=2100, value=2025, step=1
    )
with col_m:
    exercise_mode = st.radio(
        "Type d'exercice",
        ["Année civile (Janvier → Décembre)", "Année à cheval (deux grands livres)"],
        horizontal=True,
    )

st.checkbox(
    "Inclure les factures de 2024",
    key="include_2024",
    help="Si coché, l'année plancher pour la prise en compte des factures descend de 2025 à 2024.",
)

QUARTER_END_DATES = {
    "T1 — 1er trimestre (31/03)":  f"{int(selected_year)}-03-31",
    "T2 — 2ème trimestre (30/06)": f"{int(selected_year)}-06-30",
    "T3 — 3ème trimestre (30/09)": f"{int(selected_year)}-09-30",
    "T4 — 4ème trimestre (31/12)": f"{int(selected_year)}-12-31",
}

col_decl, col_trim = st.columns([1, 2])
with col_decl:
    decl_type = st.radio(
        "Type de déclaration",
        ["Annuelle", "Trimestrielle"],
        horizontal=True,
    )

quarter_suffix = ""
with col_trim:
    if decl_type == "Trimestrielle":
        selected_quarter = st.selectbox("Trimestre", list(QUARTER_END_DATES.keys()))
        reference_date = pd.Timestamp(QUARTER_END_DATES[selected_quarter])
        quarter_suffix = f"_{selected_quarter[:2]}"   # "_T1", "_T2", "_T3", "_T4"
    else:
        reference_date = pd.Timestamp(f"{int(selected_year)}-12-31")

opening_date = pd.Timestamp(f"{int(selected_year)}-01-01")

st.caption(f"Date de référence (clôture) : **{reference_date.date()}** | Date d'ouverture : **{opening_date.date()}**")

st.divider()

GL_OPTIONS = ["GL COALA", "GL PENNYLAND", "GL SAGE"]

def _gl_label_to_format(label: str) -> str:
    if label == "GL PENNYLAND":
        return "pennyland"
    if label == "GL SAGE":
        return "sage"
    return "coala"

def _parse_journals(text: str) -> List[str]:
    """Parse une saisie utilisateur en liste de journaux (séparateur virgule)."""
    if not text:
        return []
    return [j.strip().upper() for j in text.split(",") if j.strip()]

def _sage_journal_inputs(key_prefix: str) -> Tuple[List[str], List[str], bool]:
    """Affiche les deux champs journaux pour SAGE et retourne (inv, pay, is_valid)."""
    st.caption("Format GL SAGE : précisez les codes journaux (séparez par virgule).")
    col_j1, col_j2 = st.columns(2)
    with col_j1:
        inv_text = st.text_input(
            "Journaux des factures (ex: `OD` ou `OD,VE,AC`)",
            key=f"{key_prefix}_inv_j",
        )
    with col_j2:
        pay_text = st.text_input(
            "Journaux des paiements (ex: `BQ` ou `BQ,CA`)",
            key=f"{key_prefix}_pay_j",
        )
    inv = _parse_journals(inv_text)
    pay = _parse_journals(pay_text)
    valid = bool(inv) and bool(pay)
    if not valid:
        st.info("Renseignez au moins un journal facture et un journal paiement.")
    return inv, pay, valid

def _show_results(result_df, paid_df, unpaid_df, control_df, year: int, suffix: str = "",
                  payments_only_detail_df=None, payments_only_control_df=None):
    st.success("Traitement terminé avec succès.")
    c1, c2, c3 = st.columns(3)
    c1.metric("Lignes payées", len(paid_df))
    c2.metric("Lignes non payées", len(unpaid_df))
    c3.metric("Total lignes résultat", len(result_df))
    st.subheader("Résultat global")
    st.dataframe(result_df, use_container_width=True)
    st.subheader("Contrôle fournisseurs détectés")
    st.dataframe(control_df, use_container_width=True)
    if payments_only_control_df is not None and not payments_only_control_df.empty:
        st.subheader("Fournisseurs avec uniquement des paiements (ou paiements > factures)")
        st.caption(
            "Ces fournisseurs sont exclus du calcul de délai car ils n'ont "
            "pas de facture correspondante (ou les paiements excèdent les factures)."
        )
        st.dataframe(payments_only_control_df, use_container_width=True)
    excel_bytes = to_excel_bytes(
        result_df, paid_df, unpaid_df, control_df,
        payments_only_detail_df, payments_only_control_df
    )
    st.download_button(
        label="Télécharger le fichier résultat",
        data=excel_bytes,
        file_name=f"resultat_delai_paiement_{year}{suffix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

is_cheval = exercise_mode.startswith("Année à cheval")

# ============================================================
# MODE ANNÉE CIVILE
# ============================================================
if not is_cheval:
    col_fmt, col_file = st.columns([1, 3])
    with col_fmt:
        gl_fmt_label = st.selectbox("Format du grand livre", GL_OPTIONS, key="fmt_civile")
    with col_file:
        uploaded_file = st.file_uploader(
            f"Grand livre fournisseurs ({gl_fmt_label}) — Excel",
            type=["xlsx", "xls"], key="up_civile"
        )

    sage_inv_j, sage_pay_j, sage_valid = [], [], True
    if gl_fmt_label == "GL SAGE":
        sage_inv_j, sage_pay_j, sage_valid = _sage_journal_inputs("civile")

    if uploaded_file is not None:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet_options = xl.sheet_names
        except Exception as e:
            st.error(f"Impossible de lire le fichier Excel : {e}")
            st.stop()

        selected_sheet = st.selectbox("Choisir la feuille à traiter", sheet_options)

        btn_disabled = (gl_fmt_label == "GL SAGE") and not sage_valid
        if st.button("Lancer le traitement", type="primary", disabled=btn_disabled):
            try:
                uploaded_file.seek(0)
                gl_fmt = _gl_label_to_format(gl_fmt_label)
                if gl_fmt == "pennyland":
                    (result_df, paid_df, unpaid_df, control_df,
                     po_detail_df, po_ctrl_df) = process_workbook_pennyland(
                        uploaded_file, selected_sheet, reference_date, opening_date
                    )
                elif gl_fmt == "sage":
                    (result_df, paid_df, unpaid_df, control_df,
                     po_detail_df, po_ctrl_df) = process_workbook_sage(
                        uploaded_file, selected_sheet, reference_date, opening_date,
                        sage_inv_j, sage_pay_j
                    )
                else:
                    (result_df, paid_df, unpaid_df, control_df,
                     po_detail_df, po_ctrl_df) = process_workbook(
                        uploaded_file, selected_sheet, reference_date, opening_date
                    )
                _show_results(result_df, paid_df, unpaid_df, control_df,
                              int(selected_year), suffix=quarter_suffix,
                              payments_only_detail_df=po_detail_df,
                              payments_only_control_df=po_ctrl_df)
            except Exception as e:
                st.error(f"Erreur pendant le traitement : {e}")
    else:
        st.info("Choisissez le format du grand livre puis chargez le fichier Excel.")

# ============================================================
# MODE ANNÉE À CHEVAL
# ============================================================
else:
    st.info(
        f"**Année à cheval** : chargez les deux grands livres. "
        f"Seules les factures de **{int(selected_year)}** apparaîtront. "
        f"Clôture : **31/12/{int(selected_year)}**. "
        f"Chaque fichier peut avoir son propre format (COALA, PENNYLAND ou SAGE)."
    )

    col_f1, col_f2 = st.columns(2)

    with col_f1:
        st.markdown(f"#### Grand livre — Année {int(selected_year) - 1}")
        gl_fmt1_label = st.selectbox("Format GL (fichier N-1)", GL_OPTIONS, key="fmt1")
        file1 = st.file_uploader(
            f"Fichier {int(selected_year) - 1} ({gl_fmt1_label})",
            type=["xlsx", "xls"], key="file1",
        )
        sheet1 = None
        if file1 is not None:
            try:
                xl1   = pd.ExcelFile(file1)
                sheet1 = st.selectbox(
                    f"Feuille — fichier {int(selected_year) - 1}",
                    xl1.sheet_names, key="sheet1",
                )
            except Exception as e:
                st.error(f"Impossible de lire le fichier N-1 : {e}")
        sage_inv_j1, sage_pay_j1, sage_valid1 = [], [], True
        if gl_fmt1_label == "GL SAGE":
            sage_inv_j1, sage_pay_j1, sage_valid1 = _sage_journal_inputs("cheval1")

    with col_f2:
        st.markdown(f"#### Grand livre — Année {int(selected_year)}")
        gl_fmt2_label = st.selectbox("Format GL (fichier N)", GL_OPTIONS, key="fmt2")
        file2 = st.file_uploader(
            f"Fichier {int(selected_year)} ({gl_fmt2_label})",
            type=["xlsx", "xls"], key="file2",
        )
        sheet2 = None
        if file2 is not None:
            try:
                xl2   = pd.ExcelFile(file2)
                sheet2 = st.selectbox(
                    f"Feuille — fichier {int(selected_year)}",
                    xl2.sheet_names, key="sheet2",
                )
            except Exception as e:
                st.error(f"Impossible de lire le fichier N : {e}")
        sage_inv_j2, sage_pay_j2, sage_valid2 = [], [], True
        if gl_fmt2_label == "GL SAGE":
            sage_inv_j2, sage_pay_j2, sage_valid2 = _sage_journal_inputs("cheval2")

    both_ready = (
        file1 is not None and file2 is not None
        and sheet1 is not None and sheet2 is not None
    )

    if both_ready:
        fmt_combo = f"{gl_fmt1_label} + {gl_fmt2_label}"
        btn_disabled = not (sage_valid1 and sage_valid2)
        if st.button(f"Lancer le traitement à cheval ({fmt_combo})", type="primary",
                     disabled=btn_disabled):
            try:
                file1.seek(0)
                file2.seek(0)
                (result_df, paid_df, unpaid_df, control_df,
                 po_detail_df, po_ctrl_df) = process_workbook_cheval_generic(
                    file1, sheet1, _gl_label_to_format(gl_fmt1_label),
                    file2, sheet2, _gl_label_to_format(gl_fmt2_label),
                    reference_date, opening_date,
                    invoice_journals1=sage_inv_j1, payment_journals1=sage_pay_j1,
                    invoice_journals2=sage_inv_j2, payment_journals2=sage_pay_j2,
                )
                _show_results(result_df, paid_df, unpaid_df, control_df,
                              int(selected_year), suffix=f"_cheval{quarter_suffix}",
                              payments_only_detail_df=po_detail_df,
                              payments_only_control_df=po_ctrl_df)
            except Exception as e:
                st.error(f"Erreur pendant le traitement : {e}")
    else:
        st.warning("Chargez les deux fichiers Excel pour activer le traitement.")
