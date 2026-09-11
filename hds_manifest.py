import re

import pandas as pd


HDS_COLUMNS = [
    "DeliveryId",
    "BoxIds",
    "LineItemSKU",
    "LineItemName",
    "LineItemBarcode",
    "LineItemQuantity",
    "LineItemCustomField1",
    "DeliveryDate",
    "BusinessName",
    "CustomerName",
    "Phone",
    "Email",
    "Line1",
    "Line2",
    "Locality",
    "Postcode",
    "State",
    "Country3",
    "Instructions",
    "DeliveryWindow",
    "Type",
    "Service",
    "CustomField1",
    "CustomField2",
    "CustomField3",
    "CustomField4",
    "BoxWeightKG",
    "BoxHeightCM",
    "BoxWidthCM",
    "BoxLengthCM",
    "BoxSize",
]


STATE_ABBREVIATIONS = {
    "new south wales": "NSW",
    "victoria": "VIC",
    "queensland": "QLD",
    "south australia": "SA",
    "western australia": "WA",
    "tasmania": "TAS",
    "northern territory": "NT",
    "australian capital territory": "ACT",
}


def clean_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return re.sub(r"[\r\n\t]+", " ", text).strip()


def state_abbreviation(value: object) -> str:
    state = clean_text(value)
    return STATE_ABBREVIATIONS.get(state.lower(), state.upper())


def carton_count(value: object) -> int:
    text = clean_text(value)
    try:
        count = int(float(text))
    except (TypeError, ValueError):
        return 0
    return max(count, 0)


def order_names_with_tag(orders_df: pd.DataFrame, tag: str) -> list[str]:
    """Return order names carrying an exact comma-separated Shopify tag."""
    wanted = clean_text(tag).casefold()
    names = []
    for order_name, group in orders_df.groupby("Name", sort=False):
        has_tag = any(
            clean_text(item).casefold() == wanted
            for cell in group["Tags"]
            for item in clean_text(cell).split(",")
        )
        if has_tag:
            names.append(order_name)
    return names


def build_hds_manifest(manifest_df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    """Expand routed orders into the HDS one-row-per-carton CSV layout."""
    prefix = clean_text(prefix).rstrip("-") + "-"
    rows = []

    for _, order in manifest_df.iterrows():
        order_id = clean_text(order.get("D.O. No.", ""))
        cartons = carton_count(order.get("No. of Shipping Labels", 0))
        if not order_id or cartons == 0:
            continue

        delivery_id = f"{prefix}{order_id}"
        delivery_type = "business" if order_id.upper().startswith("CEW") else "residential"

        shared = {
            "DeliveryId": delivery_id,
            "LineItemSKU": "",
            "LineItemName": "",
            "LineItemBarcode": "",
            "LineItemQuantity": "",
            "LineItemCustomField1": "",
            "DeliveryDate": clean_text(order.get("Date", "")),
            "BusinessName": "",
            "CustomerName": clean_text(order.get("Deliver to", "")),
            "Phone": clean_text(order.get("Phone No.", "")),
            "Email": clean_text(order.get("Email", "")),
            "Line1": clean_text(order.get("Address 1", "")),
            "Line2": "",
            "Locality": clean_text(order.get("Address 2", "")),
            "Postcode": clean_text(order.get("Postal Code", "")),
            "State": state_abbreviation(order.get("State", "")),
            "Country3": "AUS",
            "Instructions": clean_text(order.get("Instructions", "")),
            "DeliveryWindow": "STANDARD_DAY",
            "Type": delivery_type,
            "Service": "hds",
            "CustomField1": "",
            "CustomField2": "",
            "CustomField3": "",
            "CustomField4": "",
            "BoxWeightKG": "",
            "BoxHeightCM": "",
            "BoxWidthCM": "",
            "BoxLengthCM": "",
            "BoxSize": "",
        }

        for box_number in range(1, cartons + 1):
            rows.append({
                **shared,
                "BoxIds": f"{delivery_id}-{box_number:02d}",
            })

    return pd.DataFrame(rows, columns=HDS_COLUMNS)


def hds_csv_bytes(manifest_df: pd.DataFrame, prefix: str) -> bytes:
    return build_hds_manifest(manifest_df, prefix).to_csv(index=False).encode("utf-8")
