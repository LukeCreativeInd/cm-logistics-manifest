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
    uploaded_file = st.file_uploader("Upload Clean Eats orders_export CSV file", type="csv")

    if not uploaded_file:
        return

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
            if item in bundle_items:
                total_qty -= qty  # fix: subtract bundle quantity explicitly
                continue
            elif item in family_double_items:
                total_qty += qty * 2
            else:
                total_qty += qty

        labels = math.ceil(total_qty / 20)
        phone = order.get("Billing Phone") or order.get("Phone")
        phone = format_phone(phone)

        state_map = {"VIC": "Victoria", "NSW": "New South Wales"}
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
