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

    cold_required = st.checkbox("Is there a Cold Express Pickup Required?")
    uploaded_files = st.file_uploader("Upload one or more Clean Eats CSV exports", type="csv", accept_multiple_files=True)
    generate = st.button("Generate Clean Eats Manifests")

    if uploaded_files and generate:
        orders_df = pd.concat([pd.read_csv(f) for f in uploaded_files], ignore_index=True)
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
            phone = order.get("Billing Phone") or order.get("Phone")
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

            shipping_name = order["Shipping Name"].strip()
            shipping_company = order.get("Shipping Company", "")
            shipping_company = "" if pd.isna(shipping_company) else str(shipping_company).strip()
            company_value = shipping_company if shipping_company and shipping_company.lower() != shipping_name.lower() else ""

            manifest_rows.append({
                "D.O. No.": name,
                "Date": delivery_date,
                "Address 1": order["Shipping Street"],
                "Address 2": order["Shipping City"],
                "Postal Code": str(order["Shipping Zip"]).replace("'", ""),
                "State": state,
                "Country": country,
                "Deliver to": shipping_name,
                "Company": company_value,
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
        all_tagged_names = set(cm_names) | set(mc_names) | set(cx_names)

        cm_manifest = manifest_df[manifest_df["D.O. No."].isin(cm_names)]
        mc_manifest = manifest_df[manifest_df["D.O. No."].isin(mc_names)]
        cx_manifest = manifest_df[manifest_df["D.O. No."].isin(cx_names)]
        other_manifest = manifest_df[~manifest_df["D.O. No."].isin(all_tagged_names)]

        output = BytesIO()
        with zipfile.ZipFile(output, "w") as zipf:
            def add_to_zip(df, filename):
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

            add_to_zip(cm_manifest, "CM_Manifest.xlsx")
            add_to_zip(mc_manifest, "MC_Manifest.xlsx")
            add_to_zip(cx_manifest, "CX_Manifest.xlsx")
            add_to_zip(other_manifest, "Other_Manifest.xlsx")

        output.seek(0)
        st.download_button(
            label="Download Manifests ZIP",
            data=output,
            file_name="CleanEats_Manifests.zip",
            mime="application/zip"
        )
