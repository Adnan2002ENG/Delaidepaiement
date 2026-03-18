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

    # Sinon on garde D
    if piece:
        return piece

    # Si D est vide, on tente dans E
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


def process_supplier(boundary: SupplierBoundary, supplier_df: pd.DataFrame):
    paid_rows = []
    unpaid_rows = []

    # ============================================================
    # 1) FACTURES LETTRÉES / PAIEMENTS LETTRÉS
    # ============================================================
    invoices_lettered = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] != "")
        & (supplier_df["JournalColC"].str.upper().str.startswith("A", na=False))
    ].copy()

    payments_lettered = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] != "")
    ].copy()

    if not invoices_lettered.empty:
        for letter in sorted(invoices_lettered["Lettrage"].dropna().unique()):
            inv_group = invoices_lettered[invoices_lettered["Lettrage"] == letter].copy()
            pay_group = payments_lettered[payments_lettered["Lettrage"] == letter].copy()

            if pay_group.empty:
                continue

            inv_group = inv_group.sort_values(by=["DateOperation", "original_row"], na_position="last")
            pay_group = pay_group.sort_values(by=["DateOperation", "original_row"], na_position="last")

            allocations, _ = allocate_fifo(inv_group, pay_group)

            for alloc in allocations:
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
        & (supplier_df["JournalColC"].str.upper().str.startswith("A", na=False))
        & (~supplier_df["LibelleColA"].apply(is_opening_balance_row))
    ].copy()

    unpaid_invoices = unpaid_invoices.sort_values(by=["DateOperation", "original_row"], na_position="last")

    # Lignes d'ouverture : "Total au 01/01/2025" au débit = avance
    opening_rows = supplier_df[supplier_df["LibelleColA"].apply(is_opening_balance_row)].copy()
    opening_advance_amount = round(opening_rows["Debit"].sum(), 2)

    # Paiements non lettrés réels
    unlettered_payments = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] == "")
        & (~supplier_df["LibelleColA"].apply(is_opening_balance_row))
    ].copy()

    unlettered_payments = unlettered_payments.sort_values(by=["DateOperation", "original_row"], na_position="last")

    # 2A) Imputation de l'avance d'ouverture au 01/01/2025
    if not unpaid_invoices.empty and opening_advance_amount > 0:
        pseudo_opening_payment = pd.DataFrame([{
            "DateOperation": OPENING_DATE,
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
                delay_days = int((OPENING_DATE - invoice_date).days + 1)

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
                "Montant facture": round(allocated_amount, 2),
                "Montant(s) paiement(s) total": round(allocated_amount, 2),
                "Montant facture origine": round(invoice_total, 2),
                "Délai (jours)": delay_days,
                "Remarque": classify_remark(delay_days),
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

            consumed_invoice_ids.add(int(inv["original_row"]))

        remaining_invoice_rows = []
        residual_map = {}

        for res in opening_residuals:
            inv = res["invoice_row"]
            residual_map[int(inv["original_row"])] = round(res["remaining_amount"], 2)
            remaining_invoice_rows.append(inv)

        for _, inv in unpaid_invoices.iterrows():
            key = int(inv["original_row"])
            if key not in residual_map and key not in consumed_invoice_ids:
                residual_map[key] = round(inv["Credit"], 2)
                remaining_invoice_rows.append(inv)

        unpaid_working = unpaid_invoices.copy()
        unpaid_working["RemainingAfterOpening"] = unpaid_working["original_row"].apply(
            lambda x: residual_map.get(int(x), 0.0)
        )
        unpaid_working = unpaid_working[unpaid_working["RemainingAfterOpening"] > 0].copy()
        unpaid_working["Credit"] = unpaid_working["RemainingAfterOpening"]
        unpaid_working = unpaid_working.drop(columns=["RemainingAfterOpening"])
    else:
        unpaid_working = unpaid_invoices.copy()

    # 2B) Paiements non lettrés => FIFO sur factures non lettrées
    if not unpaid_working.empty and not unlettered_payments.empty:
        fifo_allocations, fifo_residuals = allocate_fifo(unpaid_working, unlettered_payments)

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
                "Montant facture origine": invoice_total,
                "Délai (jours)": delay_days,
                "Remarque": classify_remark(delay_days),
                "Ligne source facture": int(inv["original_row"]) + 2,
            })

    else:
        # Aucun paiement non lettré ou aucune facture non lettrée => rester sur la logique "non payée"
        for _, inv in unpaid_working.iterrows():
            invoice_date = inv["DateOperation"]
            remaining_amount = round(inv["Credit"], 2)

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
                "Montant facture origine": round(inv["Credit"], 2),
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
        raise ValueError("Le fichier doit contenir au minimum 8 colonnes (A ? H).")

    suppliers = find_supplier_boundaries(df)
    if not suppliers:
        raise ValueError("Aucun fournisseur d?tect? dans la colonne A.")

    all_paid = []
    all_unpaid = []
    control_rows = []

    for boundary in suppliers:
        supplier_df = prepare_supplier_data(df, boundary)
        paid_rows, unpaid_rows = process_supplier(boundary, supplier_df)

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)

        # IMPORTANT :
        # Total cr?dit source = tous les cr?dits du fournisseur
        # SAUF les lignes dont le libell? colonne E est exactement "Total au 01/01/2025"
        total_credit_source = round(
            supplier_df.loc[
                (supplier_df["Credit"] > 0)
                & (
                    supplier_df["LibelleColE"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    != "total au 01/01/2025"
                ),
                "Credit"
            ].sum(),
            2
        )

                # Reconstitution exhaustive du total factures :
        # on additionne toutes les fractions d'une m?me facture
        # (pay?e, pay?e FIFO, non pay?e) pour retrouver son montant total
        all_rows = paid_rows + unpaid_rows

        factures_reconstituees = {}

        for r in all_rows:
            key = (
                r.get("Code fournisseur"),
                r.get("Numero facture"),
            )

            montant_ligne = float(r.get("Montant facture", 0) or 0)

            if key not in factures_reconstituees:
                factures_reconstituees[key] = 0.0

            factures_reconstituees[key] += montant_ligne

        total_factures_retenues = round(
            sum(factures_reconstituees.values()),
            2
        )

        ecart = round(total_credit_source - total_factures_retenues, 2)

        control_rows.append({
            "Code fournisseur": boundary.supplier_code,
            "Nom fournisseur": boundary.supplier_name,
            "Ligne d?but": boundary.start_row + 2,
            "Ligne fin Total": boundary.end_row + 2,
            "Nb lignes pay?es": len(paid_rows),
            "Nb lignes non pay?es": len(unpaid_rows),
            "Total cr?dit source": total_credit_source,
            "Total factures retenues": total_factures_retenues,
            "?cart": ecart,
            "Statut contr?le": "OK" if abs(ecart) < 0.01 else "?cart ? analyser",
        })

    paid_df = pd.DataFrame(all_paid)
    unpaid_df = pd.DataFrame(all_unpaid)
    control_df = pd.DataFrame(control_rows)

    result_df = pd.concat([paid_df, unpaid_df], ignore_index=True)
    if not result_df.empty:
        result_df = result_df.sort_values(
            by=["Code fournisseur", "Nom fournisseur", "Date facture", "Type"],
            na_position="last"
        ).reset_index(drop=True)

    return result_df, paid_df, unpaid_df, control_df


def to_excel_bytes(result_df: pd.DataFrame, paid_df: pd.DataFrame, unpaid_df: pd.DataFrame, control_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Resultat global")
        paid_df.to_excel(writer, index=False, sheet_name="Factures payees")
        unpaid_df.to_excel(writer, index=False, sheet_name="Factures non payees")
        control_df.to_excel(writer, index=False, sheet_name="Controle fournisseurs")

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

**Règles importantes :**
- Les **factures** sont prises seulement si le **journal colonne C commence par `A`**
- Si la **Pièce D** ressemble à un code interne type `AC25080025` ou `BQ425010004`, on cherche le numéro dans le **Libellé E**
- Les paiements non lettrés sont imputés en **FIFO** sur les factures non lettrées du même fournisseur
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
            result_df, paid_df, unpaid_df, control_df = process_workbook(uploaded_file, selected_sheet)

            st.success("Traitement terminé avec succès.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Lignes payées", len(paid_df))
            c2.metric("Lignes non payées", len(unpaid_df))
            c3.metric("Total lignes résultat", len(result_df))

            st.subheader("Résultat global")
            st.dataframe(result_df, use_container_width=True)

            st.subheader("Contrôle fournisseurs détectés")
            st.dataframe(control_df, use_container_width=True)

            excel_bytes = to_excel_bytes(result_df, paid_df, unpaid_df, control_df)
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




