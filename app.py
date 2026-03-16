import io
import re
from dataclasses import dataclass
from typing import Optional, List, Tuple

import pandas as pd
import streamlit as st


REFERENCE_DATE = pd.Timestamp("2025-12-31")
OPENING_DATE = pd.Timestamp("2025-01-01")

st.set_page_config(page_title="Déclarations délai de paiement", layout="wide")
st.title("Application de traitement des déclarations de délai de paiement")

EXPECTED_COL_INDEX = {
    "col_a": 0,    # A
    "date": 1,     # B
    "journal": 2,  # C
    "piece": 3,    # D
    "libelle": 4,  # E
    "debit": 5,    # F
    "credit": 6,   # G
    "lettrage": 7  # H
}


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
    t = normalize_text(text).lower().replace(" ", "")
    return t.startswith("totalau01/01/2025") or t.startswith("totalau1/1/2025")


def classify_remark(delay_days):
    if delay_days is None or pd.isna(delay_days):
        return "Date facture manquante - à vérifier"
    if delay_days < 0:
        return "Avance - Rien à signaler"
    if delay_days <= 60:
        return "Rien à signaler"
    if 61 <= delay_days <= 120:
        return "Demander la convention de délai de paiement du fournisseur en cas d’existence, sinon pénalité"
    return "Pénalité à payer"


def is_internal_piece_reference(piece_text: str) -> bool:
    t = normalize_text(piece_text).upper().replace(" ", "")
    if not t:
        return False

    # Cas à exclure : 2 lettres + chiffres uniquement
    # Exemples : AC25080025 / BQ425010004
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

    # Si D ressemble à un code interne type AC25080025 ou BQ425010004,
    # on ignore D et on cherche le numéro dans E.
    if piece and is_internal_piece_reference(piece):
        extracted = extract_invoice_from_libelle(libelle)
        if extracted:
            return extracted
        return libelle

    # Sinon on garde la pièce D
    if piece:
        return piece

    # Si D est vide, on tente dans le libellé
    extracted = extract_invoice_from_libelle(libelle)
    if extracted:
        return extracted

    return ""


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


def prepare_supplier_data(df: pd.DataFrame, boundary: SupplierBoundary) -> pd.DataFrame:
    supplier_df = df.iloc[boundary.start_row + 1: boundary.end_row].copy()
    supplier_df = supplier_df.reset_index(drop=False).rename(columns={"index": "original_row"})

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


def process_supplier(boundary: SupplierBoundary, supplier_df: pd.DataFrame):
    paid_rows = []
    unpaid_rows = []

    # FACTURES PAYEES = Credit > 0 + lettrage non vide + journal AC
    invoices_lettered = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] != "")
        & (supplier_df["JournalColC"].str.upper().str.startswith("AC", na=False))
    ].copy()

    # PAIEMENTS = Debit > 0 + lettrage non vide (sans filtre AC)
    payments_lettered = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] != "")
    ].copy()

    for _, inv in invoices_lettered.iterrows():
        letter = inv["Lettrage"]
        matching_payments = payments_lettered[payments_lettered["Lettrage"] == letter].copy()

        if matching_payments.empty:
            continue

        payment_date = matching_payments["DateOperation"].max()
        invoice_date = inv["DateOperation"]

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
            "Montant facture": round(inv["Credit"], 2),
            "Montant(s) paiement(s) total": round(matching_payments["Debit"].sum(), 2),
            "Délai (jours)": delay_days,
            "Remarque": remark,
            "Ligne source facture": int(inv["original_row"]) + 2,
        })

    # Lignes d'ouverture
    opening_rows = supplier_df[supplier_df["LibelleColA"].apply(is_opening_balance_row)].copy()
    opening_advance_amount = round(opening_rows["Debit"].sum(), 2)

    # FACTURES NON PAYEES = Credit > 0 + lettrage vide + journal AC + hors lignes d'ouverture
    unpaid_invoices = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] == "")
        & (supplier_df["JournalColC"].str.upper().str.startswith("AC", na=False))
        & (~supplier_df["LibelleColA"].apply(is_opening_balance_row))
    ].copy()

    unpaid_invoices = unpaid_invoices.sort_values(by=["DateOperation", "original_row"], na_position="last")

    for _, inv in unpaid_invoices.iterrows():
        invoice_date = inv["DateOperation"]
        remaining_amount = round(inv["Credit"], 2)

        if opening_advance_amount > 0:
            allocated_advance = min(opening_advance_amount, remaining_amount)

            if allocated_advance > 0:
                if pd.isna(invoice_date):
                    delay_days_advance = None
                else:
                    delay_days_advance = int((OPENING_DATE - invoice_date).days + 1)

                unpaid_rows.append({
                    "Code fournisseur": boundary.supplier_code,
                    "Nom fournisseur": boundary.supplier_name,
                    "Libellé": inv["LibelleColE"],
                    "Numero facture": inv["NumeroFacture"],
                    "Type": "Avance imputée",
                    "Journal": inv["JournalColC"],
                    "Lettrage": "",
                    "Date facture": invoice_date.date() if pd.notna(invoice_date) else None,
                    "Date paiement": OPENING_DATE.date(),
                    "Montant facture": round(allocated_advance, 2),
                    "Montant(s) paiement(s) total": round(allocated_advance, 2),
                    "Délai (jours)": delay_days_advance,
                    "Remarque": classify_remark(delay_days_advance),
                    "Ligne source facture": int(inv["original_row"]) + 2,
                })

                opening_advance_amount = round(opening_advance_amount - allocated_advance, 2)
                remaining_amount = round(remaining_amount - allocated_advance, 2)

        if remaining_amount <= 0:
            continue

        if pd.isna(invoice_date):
            delay_days = None
        else:
            delay_days = int((REFERENCE_DATE - invoice_date).days + 1)

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
            "Délai (jours)": delay_days,
            "Remarque": classify_remark(delay_days),
            "Ligne source facture": int(inv["original_row"]) + 2,
        })

    return paid_rows, unpaid_rows


def process_workbook(uploaded_file, sheet_name=None):
    if sheet_name:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
    else:
        df = pd.read_excel(uploaded_file, header=0)

    if df.shape[1] < 8:
        raise ValueError("Le fichier doit contenir au minimum 8 colonnes (A à H).")

    suppliers = find_supplier_boundaries(df)
    if not suppliers:
        raise ValueError("Aucun fournisseur détecté dans la colonne A.")

    all_paid = []
    all_unpaid = []

    for boundary in suppliers:
        supplier_df = prepare_supplier_data(df, boundary)
        paid_rows, unpaid_rows = process_supplier(boundary, supplier_df)
        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)

    paid_df = pd.DataFrame(all_paid)
    unpaid_df = pd.DataFrame(all_unpaid)

    return paid_df, unpaid_df


def to_excel_bytes(paid_df: pd.DataFrame, unpaid_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        paid_df.to_excel(writer, index=False, sheet_name="Factures payees")
        unpaid_df.to_excel(writer, index=False, sheet_name="Factures non payees")

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


with st.expander("Format attendu du fichier Excel", expanded=False):
    st.markdown(
        """
- **Colonne A** : début/fin fournisseur
- **Colonne B** : date
- **Colonne C** : journal
- **Colonne D** : pièce
- **Colonne E** : libellé
- **Colonne F** : débit
- **Colonne G** : crédit
- **Colonne H** : lettrage

**Règle importante :**
Seules les lignes dont le **journal en colonne C commence par `AC`** sont considérées comme des **factures**.
Les journaux `OD`, `BQ`, etc. sont exclus des factures, même s’il y a un montant au crédit.
"""
    )

uploaded_file = st.file_uploader("Importer le grand livre fournisseurs (Excel)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_options = xl.sheet_names
    except Exception as e:
        st.error(f"Impossible de lire le fichier Excel : {e}")
        st.stop()

    selected_sheet = st.selectbox("Choisir la feuille à traiter", sheet_options)

    if st.button("Lancer le traitement", type="primary"):
        try:
            uploaded_file.seek(0)
            paid_df, unpaid_df = process_workbook(uploaded_file, selected_sheet)

            st.success("Traitement terminé avec succès.")

            c1, c2 = st.columns(2)
            c1.metric("Factures payées", len(paid_df))
            c2.metric("Factures non payées", len(unpaid_df))

            st.subheader("Factures payées")
            st.dataframe(paid_df, use_container_width=True)

            st.subheader("Factures non payées")
            st.dataframe(unpaid_df, use_container_width=True)

            excel_bytes = to_excel_bytes(paid_df, unpaid_df)
            st.download_button(
                label="Télécharger le fichier résultat",
                data=excel_bytes,
                file_name="resultat_delai_paiement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Erreur pendant le traitement : {e}")
else:
    st.info("Chargez un fichier Excel pour démarrer le traitement.")



