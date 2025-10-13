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

def run():
    st.markdown("### Clean Eats Manifest Generator")

    uploaded_file = st.file_uploader("Upload Clean Eats orders_export CSV file", type="csv")
    generate = st.button("Generate Clean Eats Manifests")

    if uploaded_file and generate:
        orders_df = pd.read_csv(uploaded_file)
        orders_df.columns = orders_df.columns.str.strip()
        orders_df["Notes"] = orders_df["Notes"].fillna("")
        orders_df["Tags"] = orders_df["Tags"].fillna("")

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

        def format_phone(phone):
            if pd.isna(phone):
                return ""
            phone = str(phone).strip().replace(" ", "").replace("+", "")
            if phone.startswith("61"):
                phone = "0" + phone[2:]
            elif phone.startswith("4"):
                phone = "0" + phone
            return phone

        manifest_rows = []
        grouped_orders = orders_df.groupby("Name", sort=False)

        for name, group in grouped_orders:
            order = group.iloc[0]
            total_qty = 0
            for _, row in group.iterrows():
                item = row["Lineitem name"].strip()
                qty = row["Lineitem quantity"]
                if any(bundle in item for bundle in bundle_items):
                    continue
                elif item in family_double_items:
                    total_qty += qty * 2
                else:
                    total_qty += qty
            labels = math.ceil(total_qty / 24)
            phone = order.get("Shipping Phone")
            phone = format_phone(phone)

            state_map = {
                "VIC": "Victoria",
                "NSW": "New South Wales",
                "ACT": "Australian Capital Territory"
            }
            country_map = {"AU": "Australia"}
            state = state_map.get(order["Shipping Province"], order["Shipping Province"])
            country = country_map.get(order["Shipping Country"], order["Shipping Country"])

            date_match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", order["Tags"])
            delivery_date = date_match.group(1) if date_match else ""

            manifest_rows.append({
                "D.O. No.": name,
                "Date": delivery_date,
                "Address 1": order["Shipping Street"],
                "Address 2": order["Shipping City"],
                "Postal Code": str(order["Shipping Zip"]).replace("'", ""),
                "State": state,
                "Country": country,
                "Deliver to": order["Shipping Name"],
                "Phone No.": phone,
                "Time Window": "0600-1800",
                "Group": "Clean Eats Australia",
                "No. of Shipping Labels": labels,
                "Line Items": total_qty,
                "Email": order["Email"],
                "Instructions": order["Notes"]
            })

        manifest_df = pd.DataFrame(manifest_rows)

        cm_names = orders_df[orders_df["Tags"].str.contains("CM")]["Name"].unique()
        mc_names = orders_df[orders_df["Tags"].str.contains("MC")]["Name"].unique()
        cx_names = orders_df[orders_df["Tags"].str.contains("CX")]["Name"].unique()
        dk_names = orders_df[orders_df["Tags"].str.contains("DK")]["Name"].unique()
        all_tagged_names = set(cm_names) | set(mc_names) | set(cx_names)

        cm_manifest = manifest_df[manifest_df["D.O. No."].isin(cm_names)]
        mc_manifest = manifest_df[manifest_df["D.O. No."].isin(mc_names)]
        cx_manifest = manifest_df[manifest_df["D.O. No."].isin(cx_names)]
        other_manifest = manifest_df[~manifest_df["D.O. No."].isin(all_tagged_names)]

        # For MC manifest: use Shipping Company as Deliver to, fallback to Shipping Name
        orders_df["Name"] = orders_df["Name"].astype(str).str.strip()
        mc_manifest["D.O. No."] = mc_manifest["D.O. No."].astype(str).str.strip()

        def get_valid_company_or_name(group):
            company = group["Shipping Company"].dropna().astype(str).str.strip()
            name = group["Shipping Name"].dropna().astype(str).str.strip()
            if not company.empty and company.iloc[0].lower() != "nan" and company.iloc[0] != "":
                return company.iloc[0]
            elif not name.empty:
                return name.iloc[0]
            else:
                return ""

        fallback_dict_grouped = orders_df.groupby("Name").apply(get_valid_company_or_name).to_dict()
        mc_manifest["Deliver to"] = mc_manifest["D.O. No."].map(fallback_dict_grouped).fillna("")
        mc_manifest = mc_manifest.drop(columns=["Company"], errors="ignore")

        output = BytesIO()
        with zipfile.ZipFile(output, "w") as zipf:
            def add_to_zip_excel(df, filename):
                if df.empty:
                    return
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df["Phone No."] = df["Phone No."].astype(str).str.replace(r"\.0$", "", regex=True)
                    df.to_excel(writer, index=False, sheet_name='Manifest')
                    workbook = writer.book
                    worksheet = writer.sheets['Manifest']
                    text_fmt = workbook.add_format({'num_format': '@'})
                    col_index = df.columns.get_loc("Phone No.")
                    worksheet.set_column(col_index, col_index, None, text_fmt)
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

            # CX Ready Manifest (kept as-is)
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
                    "WEIGHT (KG)": (cx_manifest["Line Items"].astype(float) * 0.4).round(2),
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
                # Delivery date = today + 2 days (labels day to +2 rule)
                dk_date = (datetime.now() + timedelta(days=2))
                dk_date_str = dk_date.strftime("%d/%m/%Y")
                for order_name, group in dk_src.groupby("Name", sort=False):
                    # find corresponding manifest entry
                    mrow = manifest_df[manifest_df["D.O. No." ] == order_name].iloc[0]
                    # delivery type based on tags: CEW -> Commercial, CEA -> Residential (default Residential)
                    tags = " ".join(group["Tags"].astype(str).unique())
                    if "CEW" in tags:
                        delivery_type = "Commercial"
                    elif "CEA" in tags:
                        delivery_type = "Residential"
                    else:
                        delivery_type = "Residential"

                    # Location: Shipping Company if present
                    ship_company = group["Shipping Company"].dropna().astype(str).str.strip()
                    location = ship_company.iloc[0] if not ship_company.empty and ship_company.iloc[0] not in ["", "nan", "NaN"] else ""

                    # Phone, email, notes
                    phone = mrow["Phone No."]
                    email_vals = group["Email"].dropna().astype(str).unique()
                    email = email_vals[0] if len(email_vals) else ""

                    notes = group["Notes"].dropna().astype(str).unique()
                    notes_val = notes[0] if len(notes) else ""
                    instr = mrow["Instructions"]

                    # State abbreviation for DK (use raw province where available)
                    state_abbrev_vals = group["Shipping Province"].dropna().astype(str).unique()
                    state_abbrev = state_abbrev_vals[0] if len(state_abbrev_vals) else ""

                    dk_rows.append({
                        "Order ID": order_name,
                        "Date": dk_date_str,
                        "Time Window": "7am - 6pm",
                        "Notes": notes_val,
                        "Address 1": mrow["Address 1"],
                        "Address 2": "",
                        "Address 3": "",
                        "Postal Code": mrow["Postal Code"],
                        "City": mrow["Address 2"],
                        "State": state_abbrev,
                        "Country": "Australia",
                        "Location": location,
                        "Last Name": "",
                        "Phone": phone,
                        "Delivery Instructions": instr,
                        "Email": email,
                        "DELIVERY TYPE": delivery_type,
                        "Volume": mrow["No. of Shipping Labels"],
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
