import streamlit as st
import pandas as pd
import math
import zipfile
from io import BytesIO
import re
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tempfile import NamedTemporaryFile

NAN_LIKE = {"nan", "none", "null", ""}

def clean_cell(x: object) -> str:
    """Return a safe string with NAN-like values mapped to empty."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    return "" if s.lower() in NAN_LIKE else s

def to_clean_str(val: object) -> str:
    s = clean_cell(val)
    if s.startswith("'"):
        s = s[1:]
    # remove trailing .0 when it's just an integer-as-float
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

def to_intish_str(val: object) -> str:
    s = to_clean_str(val)
    if re.fullmatch(r"\d+(\.\d+)?", s):
        try:
            f = float(s)
            if f.is_integer():
                return str(int(f))
        except:
            pass
    return s

def format_phone(phone: object) -> str:
    p = clean_cell(phone).replace(" ", "").replace("+", "")
    if not p:
        return ""
    if p.startswith("61"):
        p = "0" + p[2:]
    elif p.startswith("4"):
        p = "0" + p
    return p

def run():
    st.markdown("### Clean Eats Manifest Generator")

    uploaded_file = st.file_uploader("Upload Clean Eats orders_export CSV file", type="csv")
    generate = st.button("Generate Clean Eats Manifests")

    if not (uploaded_file and generate):
        return

    # Read everything as string, then normalise ALL cells
    orders_df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    orders_df = orders_df.applymap(clean_cell)
    orders_df.columns = orders_df.columns.str.strip()

    # Ensure expected columns exist (as empty if missing)
    expected_cols = [
        "Notes","Tags","Shipping Phone","Shipping Street","Shipping City","Shipping Zip",
        "Shipping Province","Shipping Country","Shipping Name","Shipping Company","Email",
        "Name","Lineitem name","Lineitem quantity"
    ]
    for c in expected_cols:
        if c not in orders_df.columns:
            orders_df[c] = ""

    bundle_items = [
        "CARB LOVER'S FEAST",
        "SUPER CHARGED CALORIES",
        "FEED ME BEEF",
        "GIVE ME CHICKEN",
        "I WON'T PAS(TA) ON THIS MEAL",
        "THE MEGA PACK",
        "MAKE YOUR OWN MEGA PACK",
        "CARB HATERS FEAST",
        "UNDER CHARGED CALORIES",
        "VEGGIE LOVERS PACK",
        "Clean Eats Meal Plan"
    ]

    family_double_items = [
        "Family Mac and 3 Cheese Pasta Bake",
        "Baked Family Lasagna"
    ]

    manifest_rows = []
    for name, group in orders_df.groupby("Name", sort=False):
        order = group.iloc[0]
        total_qty = 0
        for _, row in group.iterrows():
            item = clean_cell(row["Lineitem name"])
            qty_raw = clean_cell(row["Lineitem quantity"])
            try:
                qty = int(float(qty_raw))
            except:
                qty = 0

            if any(bundle in item for bundle in bundle_items):
                continue
            elif item in family_double_items:
                total_qty += qty * 2
            else:
                total_qty += qty

        labels = math.ceil(total_qty / 24) if total_qty else 0

        state_map = {"VIC": "Victoria", "NSW": "New South Wales", "ACT": "Australian Capital Territory"}
        country_map = {"AU": "Australia"}
        raw_state = clean_cell(order["Shipping Province"])
        raw_country = clean_cell(order["Shipping Country"])
        state = state_map.get(raw_state, raw_state)
        country = country_map.get(raw_country, raw_country)

        # Date pulled from tags like dd/mm/yyyy if present
        tags_blob = clean_cell(order["Tags"])
        m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", tags_blob)
        delivery_date = m.group(1) if m else ""

        manifest_rows.append({
            "D.O. No.": to_clean_str(name),
            "Date": delivery_date,
            "Address 1": clean_cell(order["Shipping Street"]),
            "Address 2": clean_cell(order["Shipping City"]),
            "Postal Code": to_clean_str(order["Shipping Zip"]),
            "State": state,
            "Country": country,
            "Deliver to": clean_cell(order["Shipping Name"]),
            "Phone No.": to_clean_str(format_phone(order["Shipping Phone"])),
            "Time Window": "0600-1800",
            "Group": "Clean Eats Australia",
            "No. of Shipping Labels": labels,
            "Line Items": total_qty,
            "Email": clean_cell(order["Email"]),
            "Instructions": clean_cell(order["Notes"])
        })

    manifest_df = pd.DataFrame(manifest_rows)

    # Collect tags per order (normalised)
    tag_series = orders_df.groupby("Name")["Tags"].agg(lambda s: " ".join(map(clean_cell, s)))
    def names_with(tag):
        return tag_series[tag_series.str.contains(tag, na=False)].index.tolist()

    cm_names = names_with("CM")
    mc_names = names_with("MC")
    cx_names = names_with("CX")
    dk_names = names_with("DK")

    # Exclude DK from Other as well
    all_tagged_names = set(cm_names) | set(mc_names) | set(cx_names) | set(dk_names)

    cm_manifest = manifest_df[manifest_df["D.O. No."].isin(cm_names)]
    mc_manifest = manifest_df[manifest_df["D.O. No."].isin(mc_names)]
    cx_manifest = manifest_df[manifest_df["D.O. No."].isin(cx_names)]
    other_manifest = manifest_df[~manifest_df["D.O. No."].isin(all_tagged_names)]

    # MC: Deliver to uses Shipping Company if present, else Shipping Name — cleaned
    def company_or_name(group: pd.DataFrame) -> str:
        comp = clean_cell(group["Shipping Company"].iloc[0]) if len(group) else ""
        name = clean_cell(group["Shipping Name"].iloc[0]) if len(group) else ""
        return comp if comp else name

    if not mc_manifest.empty:
        fallback = orders_df.groupby("Name").apply(company_or_name).to_dict()
        mc_manifest = mc_manifest.copy()
        mc_manifest["Deliver to"] = mc_manifest["D.O. No."].map(fallback).fillna("")

    output = BytesIO()
    with zipfile.ZipFile(output, "w") as zipf:
        def add_to_zip_excel(df, filename):
            if df.empty:
                return
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df = df.copy()
                # keep textual columns as text
                for col in ["Phone No.","Postal Code"]:
                    if col in df.columns:
                        df[col] = df[col].astype(str)
                df.to_excel(writer, index=False, sheet_name='Manifest')
                wb = writer.book
                ws = writer.sheets['Manifest']
                text_fmt = wb.add_format({'num_format': '@'})
                for col in ["Phone No.","Postal Code"]:
                    if col in df.columns:
                        idx = df.columns.get_loc(col)
                        ws.set_column(idx, idx, None, text_fmt)
            zipf.writestr(filename, buffer.getvalue())

        def add_csv_to_zip(df, filename):
            if df.empty:
                return
            csv_buffer = df.to_csv(index=False).encode('utf-8-sig')
            zipf.writestr(filename, csv_buffer)

        add_to_zip_excel(cm_manifest, "CM_Manifest.xlsx")
        add_to_zip_excel(mc_manifest, "MC_Manifest.xlsx")
        add_to_zip_excel(cx_manifest, "CX_Manifest.xlsx")
        add_to_zip_excel(other_manifest, "Other_Manifest.xlsx")

        # CX Ready Manifest (same logic)
        if not cx_manifest.empty:
            cx_ready_body = pd.DataFrame({
                "INV NO.": cx_manifest["D.O. No."],
                "DELIVERY DATE": pd.to_datetime(cx_manifest["Date"], format="%d/%m/%Y", errors='coerce') + timedelta(days=1),
                "STORE NO": "",
                "STORE NAME": cx_manifest["Deliver to"],
                "ADDRESS": cx_manifest["Address 1"],
                "SUBURB": cx_manifest["Address 2"],
                "STATE": cx_manifest["State"],
                "POSTCODE": cx_manifest["Postal Code"],
                "CARTONS": cx_manifest["No. of Shipping Labels"],
                "PALLETS": "",
                "WEIGHT (KG)": (pd.to_numeric(cx_manifest["Line Items"], errors="coerce").fillna(0) * 0.4).round(2),
                "INV. VALUE": "",
                "COD": "",
                "TEMP": "chilled",
                "COMMENT": cx_manifest["Instructions"]
            })
            cx_ready_body["DELIVERY DATE"] = cx_ready_body["DELIVERY DATE"].dt.strftime("%d/%m/%Y")

            wb = load_workbook("cx_manifest_template.xlsx")
            ws = wb.active
            for r_idx, row in enumerate(dataframe_to_rows(cx_ready_body, index=False, header=False), start=6):
                for c_idx, value in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    if cell.coordinate in ws.merged_cells:
                        continue
                    cell.value = "" if pd.isna(value) else str(value)

            with NamedTemporaryFile() as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                zipf.writestr("CX_Ready_Manifest.xlsx", tmp.read())

        # DK Distribution Manifest (CSV)
        if len(dk_names) > 0:
            dk_src = orders_df[orders_df["Name"].isin(dk_names)]
            dk_rows = []
            dk_date_str = (datetime.now() + timedelta(days=2)).strftime("%d/%m/%Y")

            for order_name, group in dk_src.groupby("Name", sort=False):
                order_name_clean = to_clean_str(order_name)
                mrow = manifest_df[manifest_df["D.O. No."] == order_name_clean].iloc[0]

                tags_blob = " ".join(map(clean_cell, group["Tags"].unique()))
                delivery_type = "Commercial" if "CEW" in tags_blob else "Residential"

                # Location = Shipping Company if present, else Shipping Name
                ship_company = clean_cell(group["Shipping Company"].iloc[0]) if len(group) else ""
                ship_name = clean_cell(group["Shipping Name"].iloc[0]) if len(group) else ""
                location = ship_company if ship_company else ship_name

                email_vals = [clean_cell(x) for x in group["Email"].unique() if clean_cell(x)]
                email = email_vals[0] if email_vals else ""

                notes_vals = [clean_cell(x) for x in group["Notes"].unique() if clean_cell(x)]
                notes_val = notes_vals[0] if notes_vals else ""

                state_vals = [clean_cell(x) for x in group["Shipping Province"].unique() if clean_cell(x)]
                state_abbrev = state_vals[0] if state_vals else ""

                dk_rows.append({
                    "Order ID": order_name_clean,
                    "Date": dk_date_str,
                    "Time Window": "7am - 6pm",
                    "Notes": notes_val,
                    "Address 1": clean_cell(mrow["Address 1"]),
                    "Address 2": "",
                    "Address 3": "",
                    "Postal Code": to_clean_str(mrow["Postal Code"]),
                    "City": clean_cell(mrow["Address 2"]),
                    "State": state_abbrev,
                    "Country": "Australia",
                    "Location": location,
                    "Last Name": "",
                    "Phone": to_clean_str(mrow["Phone No."]),
                    "Delivery Instructions": clean_cell(mrow["Instructions"]),
                    "Email": email,
                    "DELIVERY TYPE": delivery_type,
                    "Volume": to_intish_str(mrow["No. of Shipping Labels"]),
                    "NOTES": ""
                })

            dk_df = pd.DataFrame(dk_rows, columns=[
                "Order ID","Date","Time Window","Notes","Address 1","Address 2","Address 3",
                "Postal Code","City","State","Country","Location","Last Name","Phone",
                "Delivery Instructions","Email","DELIVERY TYPE","Volume","NOTES"
            ])
            add_csv_to_zip(dk_df, "DK_Manifest.csv")

    output.seek(0)
    st.download_button(
        label="Download Manifests ZIP",
        data=output,
        file_name="CleanEats_Manifests.zip",
        mime="application/zip"
    )
