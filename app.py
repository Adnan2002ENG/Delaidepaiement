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
                     mode: str = "civile", year_filter: Optional[int] = None):
    paid_rows = []
    unpaid_rows = []

    has_source_file = "source_file" in supplier_df.columns

    # ============================================================
    # 1) FACTURES LETTRÉES / PAIEMENTS LETTRÉS
    # ============================================================
    def _invoice_year_ok(s):
        if year_filter is None:
            return pd.Series([True] * len(s), index=s.index)
        return s.dt.year == year_filter

    invoices_lettered = supplier_df[
        (supplier_df["Credit"] > 0)
        & (supplier_df["Lettrage"] != "")
        & (supplier_df["JournalColC"].str.upper().str.startswith("A", na=False))
        & _invoice_year_ok(supplier_df["DateOperation"])
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

            # ---- MODE À CHEVAL : exclure les paiements cross-fichier ----
            if mode == "cheval" and has_source_file:
                inv_files = set(inv_group["source_file"].unique())
                pay_group_same = pay_group[pay_group["source_file"].isin(inv_files)].copy()

                if pay_group_same.empty:
                    # Tous les paiements sont d'une autre année → facture non payée
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

                pay_group = pay_group_same
            # ---- FIN MODE À CHEVAL ----

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
        & _invoice_year_ok(supplier_df["DateOperation"])
    ].copy()

    unpaid_invoices = unpaid_invoices.sort_values(by=["DateOperation", "original_row"], na_position="last")

    opening_rows = supplier_df[supplier_df["LibelleColA"].apply(is_opening_balance_row)].copy()
    opening_advance_amount = round(opening_rows["Debit"].sum(), 2)

    unlettered_payments = supplier_df[
        (supplier_df["Debit"] > 0)
        & (supplier_df["Lettrage"] == "")
        & (~supplier_df["LibelleColA"].apply(is_opening_balance_row))
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


def _build_control_row(boundary, paid_rows, unpaid_rows, supplier_df):
    """Calcule la ligne de contrôle pour un fournisseur."""
    total_credit_source = round(
        supplier_df.loc[
            (supplier_df["Credit"] > 0)
            & (~supplier_df["LibelleColE"].apply(is_opening_balance_row)),
            "Credit"
        ].sum(),
        2
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
        "Total factures retenues": total_factures_retenues,
        "Écart": ecart,
        "Statut contrôle": "OK" if abs(ecart) < 0.01 else "Écart à analyser",
    }


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

    for boundary in suppliers:
        supplier_df = prepare_supplier_data(df, boundary)
        paid_rows, unpaid_rows = process_supplier(boundary, supplier_df, reference_date, opening_date, mode="civile", year_filter=reference_date.year)

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(_build_control_row(boundary, paid_rows, unpaid_rows, supplier_df))

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

        all_paid.extend(paid_rows)
        all_unpaid.extend(unpaid_rows)
        control_rows.append(_build_control_row(boundary, paid_rows, unpaid_rows, supplier_df))

    return _finalize_results(all_paid, all_unpaid, control_rows)


def _finalize_results(all_paid, all_unpaid, control_rows):
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


def to_excel_bytes(result_df: pd.DataFrame, paid_df: pd.DataFrame,
                   unpaid_df: pd.DataFrame, control_df: pd.DataFrame) -> bytes:
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


# ============================================================
# INTERFACE UTILISATEUR
# ============================================================

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
- En mode **Année à cheval** : un paiement lettré avec une facture d'une autre année est ignoré ; la facture est traitée comme non payée
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

reference_date = pd.Timestamp(f"{int(selected_year)}-12-31")
opening_date = pd.Timestamp(f"{int(selected_year)}-01-01")

st.caption(f"Date de référence (clôture) : **{reference_date.date()}** | Date d'ouverture : **{opening_date.date()}**")

st.divider()

is_cheval = exercise_mode.startswith("Année à cheval")

# ============================================================
# MODE ANNÉE CIVILE
# ============================================================
if not is_cheval:
    uploaded_file = st.file_uploader(
        "Importer le grand livre fournisseurs (Excel)", type=["xlsx", "xls"]
    )

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
                result_df, paid_df, unpaid_df, control_df = process_workbook(
                    uploaded_file, selected_sheet, reference_date, opening_date
                )

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
                    file_name=f"resultat_delai_paiement_{int(selected_year)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"Erreur pendant le traitement : {e}")
    else:
        st.info("Chargez un fichier Excel pour démarrer le traitement.")

# ============================================================
# MODE ANNÉE À CHEVAL
# ============================================================
else:
    st.info(
        f"**Année à cheval** : chargez les deux grands livres. "
        f"Les paiements lettrés avec des factures d'une autre période seront ignorés. "
        f"Date de clôture : **31/12/{int(selected_year)}**."
    )

    col_f1, col_f2 = st.columns(2)

    with col_f1:
        st.markdown(f"**Grand livre — Année {int(selected_year) - 1}**")
        file1 = st.file_uploader(
            f"Fichier grand livre {int(selected_year) - 1}",
            type=["xlsx", "xls"],
            key="file1",
        )
        sheet1 = None
        if file1 is not None:
            try:
                xl1 = pd.ExcelFile(file1)
                sheet1 = st.selectbox(
                    f"Feuille du fichier {int(selected_year) - 1}",
                    xl1.sheet_names,
                    key="sheet1",
                )
            except Exception as e:
                st.error(f"Impossible de lire le fichier 1 : {e}")

    with col_f2:
        st.markdown(f"**Grand livre — Année {int(selected_year)}**")
        file2 = st.file_uploader(
            f"Fichier grand livre {int(selected_year)}",
            type=["xlsx", "xls"],
            key="file2",
        )
        sheet2 = None
        if file2 is not None:
            try:
                xl2 = pd.ExcelFile(file2)
                sheet2 = st.selectbox(
                    f"Feuille du fichier {int(selected_year)}",
                    xl2.sheet_names,
                    key="sheet2",
                )
            except Exception as e:
                st.error(f"Impossible de lire le fichier 2 : {e}")

    both_ready = file1 is not None and file2 is not None and sheet1 is not None and sheet2 is not None

    if both_ready:
        if st.button("Lancer le traitement à cheval", type="primary"):
            try:
                file1.seek(0)
                file2.seek(0)
                result_df, paid_df, unpaid_df, control_df = process_workbook_cheval(
                    file1, sheet1, file2, sheet2, reference_date, opening_date
                )

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
                    file_name=f"resultat_delai_paiement_cheval_{int(selected_year)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"Erreur pendant le traitement : {e}")
    else:
        st.warning("Chargez les deux fichiers Excel pour activer le traitement.")
