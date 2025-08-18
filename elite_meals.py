import streamlit as st
import pandas as pd
import math
import zipfile
from io import BytesIO
import re
from datetime import datetime


def run():
    st.markdown("### Elite Meals Manifest Generator")

    uploaded_file = st.file_uploader("Upload Elite Meals orders_export CSV file", type="csv")

    if not uploaded_file:
        return

    orders_df = pd.read_csv(uploaded_file)
    orders_df.columns = orders_df.columns.str.strip()
    orders_df["Notes"] = orders_df["Notes"].fillna("")

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
        total_qty = group["Lineitem quantity"].sum()
        labels = math.ceil(total_qty / 20)

        phone = order.get("Billing Phone") or order.get("Phone")
        phone = format_phone(phone)

        state_map = {"VIC": "Victoria", "NSW": "New South Wales"}
        country_map = {"AU": "Australia"}
        state = state_map.get(order["Shipping Province"], order["Shipping Province"])
        country = country_map.get(order["Shipping Country"], order["Shipping Country"])

        date_match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", order.get("Tags", ""))
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
                "Company": order.get("Shipping Company", ""),
            "Phone No.": phone,
            "Time Window": "0600-1800",
            "Group": "Elite Meals",
            "No. of Shipping Labels": labels,
            "Line Items": total_qty,
            "Email": order["Email"],
            "Instructions": order["Notes"]
        })

    manifest_df = pd.DataFrame(manifest_rows)

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

        add_to_zip(manifest_df, "EliteMeals_Manifest.xlsx")

    output.seek(0)
    st.download_button(
        label="Download Manifest",
        data=output,
        file_name="EliteMeals_Manifest.zip",
        mime="application/zip"
    )
