"""Shared constants for file handling and UI."""

UPLOAD_EXTENSIONS = {".csv", ".xlsx", ".xls"}
STREAMLIT_UPLOAD_TYPES = ["csv", "xlsx", "xls"]

# UTF-8 with BOM helps Excel on Windows recognize UTF-8 CSV correctly.
CSV_ENCODING = "utf-8-sig"

# Export as plain text (never scientific notation) — Excel reads long IDs as numbers otherwise.
# Keys match Excel header text after OUTPUT_HEADER_DISPLAY is applied.
LONG_NUMERIC_ID_COLUMNS: frozenset[str] = frozenset({"External ID", "UPC Code"})

# Narrow navy separator columns inserted in .xlsx after these headers (between groups).
# Anchors are NetSuite/upload column names; resolved to Excel labels via OUTPUT_HEADER_DISPLAY.
GROUP_SEPARATOR_AFTER_UPLOAD_KEYS: tuple[str, ...] = (
    "MSRP PRICE",
    "NETSUITE LINK",
    "ASIN 4.30.26",
    " Amazon: Stock",
    "AMZ TAG",
    "has_buy_box FBM ",
    "STORE SALES 4.20-4.26.26",
    "AMZ SALES 4.19-4.25.26",
    "FC ON ORDER ",
    " afn-researching-quantity (inbound inv.not avail. for sale)",
)
# Dark navy blue bar (matches template separator column).
GROUP_SEPARATOR_FILL_HEX: str = "002060"
GROUP_SEPARATOR_COLUMN_WIDTH: float = 2.25

# Header row box size (matches template): width × height in screen pixels @ 96dpi.
# Excel stores row height in points and column width in character units — see exporter conversion.
HEADER_COLUMN_WIDTH_PX: float = 87
HEADER_ROW_HEIGHT_PX: float = 146

# Upload column name → Excel row-1 header (adds snapshot dates from screenshots).
OUTPUT_HEADER_DISPLAY: dict[str, str] = {}

_BUY_BOX_5126 = (
    " Buy Box: Is FBA",
    " Buy Box Seller",
    " Buy Box: Stock",
    " Amazon: Current BUY BOX  Price",
    " Buy Box Eligible Offer Counts: New FBA",
    " Buy Box Eligible Offer Counts: New FBM",
    " Amazon: Stock",
    " Bought in past month",
)
for _k in _BUY_BOX_5126:
    OUTPUT_HEADER_DISPLAY[_k] = "5.1.26 " + _k.strip()

# AJ–AP reference: date only ACTIVE/INACTIVE (both), CURRENT SHIP TEMP; not FBM QTY/PRICE, AMZ PRICE/TAG.
OUTPUT_HEADER_DISPLAY[" ACTIVE/INACTIVE"] = "4.30.26 ACTIVE/INACTIVE"
OUTPUT_HEADER_DISPLAY[" ACTIVE/INACTIVE.1"] = "4.30.26 ACTIVE/INACTIVE"
OUTPUT_HEADER_DISPLAY[" CURRENT SHIP TEMP (MIG-STANDARD, SHIP-ADD 9.99)"] = (
    "4.30.26 CURRENT SHIP TEMP (MIG-STANDARD, SHIP-ADD 9.99)"
)

_REPRICING_43026 = (
    "repricing_enabled FBA ",
    "repricing_enabled MFN ",
    "repricing_enabled FBM ",
    "offer_position FBA ",
    "offer_position MFN ",
    "offer_position FBM ",
    "min_price FBA ",
    "min_price MFN ",
    "min_price FBM ",
    "max_price FBA ",
    "max_price MFN ",
    "max_price FBM ",
    "current_price FBA ",
    "current_price MFN ",
    "current_price FBM ",
    "current_shipping FBA ",
    "current_shipping MFN ",
    "current_shipping FBM ",
    "featured FBA ",
    "featured MFN ",
    "featured FBM ",
    "is_buy_box_winner FBA ",
    "is_buy_box_winner MFN ",
    "is_buy_box_winner FBM ",
    "has_buy_box FBA ",
    "has_buy_box MFN ",
    "has_buy_box FBM ",
)
for _k in _REPRICING_43026:
    OUTPUT_HEADER_DISPLAY[_k] = "4.30.26 " + _k

_SUPPLY_43026 = (
    " Total Days of Supply (including units from open shipments) ",
    " Recommended replenishment qty",
    "STORE INV ",
    "STORE ON ORDER ",
    "STORE MINS ",
    "FC INV ",
    "FC ON ORDER ",
)
for _k in _SUPPLY_43026:
    _trail = " " if _k.endswith(" ") else ""
    OUTPUT_HEADER_DISPLAY[_k] = "4.30.26 " + _k.strip() + _trail

_INV_AGE_43026 = (
    " inv-age-0-to-90-days",
    " inv-age-91-to-180-days",
    " inv-age-181-to-270-days",
    " inv-age-271-to-365-days",
    " inv-age-365-plus-days",
)
for _k in _INV_AGE_43026:
    OUTPUT_HEADER_DISPLAY[_k] = "4.30.26 " + _k.strip()

OUTPUT_HEADER_DISPLAY["Recomened Removel+"] = "4.30.26 Recomened Removal +"

# Single AFN block (after inv-age / removal) — matches reference layout; no duplicate metrics.
_AF_SINGLE_BLOCK = (
    " afn-total-quantity",
    " afn-fulfillable-quantity",
    " afn-unsellable-quantity",
    " afn-reserved-quantity",
    " afn-inbound-working-quantity",
    " afn-inbound-shipped-quantity",
    " afn-inbound-receiving-quantity",
    " afn-researching-quantity (inbound inv.not avail. for sale)",
)
for _k in _AF_SINGLE_BLOCK:
    OUTPUT_HEADER_DISPLAY[_k] = "4.30.26 " + _k.strip()

# Trailing NetSuite columns (UPC Code … AMZ SOURCING COST): no date prefix in Excel headers.

# Excel (.xlsx) header row fills — keys match post-display header text (ARGB-ish hex, no #).
# _RAW_HEADER_FILLS uses upload keys; HEADER_FILL_HEX_BY_COLUMN maps to dated labels.
_RAW_HEADER_FILLS: dict[str, str] = {
    # Links — yellow
    "AMAZON LINK": "FFFF00",
    "SELLER CENTRAL LINK": "FFFF00",
    "NETSUITE LINK": "FFFF00",
    # ASIN block
    "MASTER PARENT ASIN": "D9D9D9",
    "MASTER CHILD ASIN": "D9D9D9",
    "Parent ASIN 4.30.26": "FF66CC",
    "Parent Asin Match 4.30.26": "FFFFFF",
    "ASIN 4.30.26": "D9D9D9",
    # Buy Box block — light gray
    " Buy Box: Is FBA": "D9D9D9",
    " Buy Box Seller": "D9D9D9",
    " Buy Box: Stock": "D9D9D9",
    " Amazon: Current BUY BOX  Price": "D9D9D9",
    " Buy Box Eligible Offer Counts: New FBA": "D9D9D9",
    " Buy Box Eligible Offer Counts: New FBM": "D9D9D9",
    " Amazon: Stock": "D9D9D9",
    " Bought in past month": "A6A6A6",
    # Pricing / flags — light gray, AMZ TAG green
    " ACTIVE/INACTIVE": "D9D9D9",
    "FBM QTY": "D9D9D9",
    "FBM PRICE": "D9D9D9",
    " CURRENT SHIP TEMP (MIG-STANDARD, SHIP-ADD 9.99)": "D9D9D9",
    " ACTIVE/INACTIVE.1": "D9D9D9",
    "AMZ PRICE": "D9D9D9",
    "AMZ TAG": "C6EFCE",
    # Repricing grid — gray
    "repricing_enabled FBA ": "D9D9D9",
    "repricing_enabled MFN ": "D9D9D9",
    "repricing_enabled FBM ": "D9D9D9",
    "offer_position FBA ": "D9D9D9",
    "offer_position MFN ": "D9D9D9",
    "offer_position FBM ": "D9D9D9",
    "min_price FBA ": "D9D9D9",
    "min_price MFN ": "D9D9D9",
    "min_price FBM ": "D9D9D9",
    "max_price FBA ": "D9D9D9",
    "max_price MFN ": "D9D9D9",
    "max_price FBM ": "D9D9D9",
    "current_price FBA ": "D9D9D9",
    "current_price MFN ": "D9D9D9",
    "current_price FBM ": "D9D9D9",
    "current_shipping FBA ": "D9D9D9",
    "current_shipping MFN ": "D9D9D9",
    "current_shipping FBM ": "D9D9D9",
    "featured FBA ": "D9D9D9",
    "featured MFN ": "D9D9D9",
    "featured FBM ": "D9D9D9",
    "is_buy_box_winner FBA ": "D9D9D9",
    "is_buy_box_winner MFN ": "D9D9D9",
    "is_buy_box_winner FBM ": "D9D9D9",
    "has_buy_box FBA ": "D9D9D9",
    "has_buy_box MFN ": "D9D9D9",
    "has_buy_box FBM ": "D9D9D9",
    # Sales — greens
    "STORE SALES 4.20-4.26.26": "92D050",
    "AMZ SALES 4.19-4.25.26": "C6EFCE",
    # Supply / store — grays, yellow, peach tones
    " Total Days of Supply (including units from open shipments) ": "D9D9D9",
    " Recommended replenishment qty": "FFFF00",
    "STORE INV ": "FFF2CC",
    "STORE ON ORDER ": "FFE599",
    "STORE MINS ": "FFD966",
    "FC INV ": "FFF2CC",
    "FC ON ORDER ": "FFE599",
    # Inventory age ladder
    " inv-age-0-to-90-days": "375623",
    " inv-age-91-to-180-days": "C6EFCE",
    " inv-age-181-to-270-days": "F4B084",
    " inv-age-271-to-365-days": "FFC000",
    " inv-age-365-plus-days": "FF0000",
    "Recomened Removel+": "7030A0",
    " afn-total-quantity": "FFFF00",
    " afn-fulfillable-quantity": "C6EFCE",
    " afn-unsellable-quantity": "FF0000",
    " afn-reserved-quantity": "FFFFFF",
    "Reserved Customer Order": "FFFFFF",
    " afn-inbound-working-quantity": "FFFFFF",
    " afn-inbound-shipped-quantity": "FFFFFF",
    " afn-inbound-receiving-quantity": "FFFFFF",
    " afn-researching-quantity (inbound inv.not avail. for sale)": "FFFFFF",
    # Trailing NetSuite pricing block — light gray
    "UPC Code": "D9D9D9",
    "Price Level": "D9D9D9",
    "Unit Price": "D9D9D9",
    "Current Price": "D9D9D9",
    "Average Cost": "D9D9D9",
    "Purchase Price": "D9D9D9",
    "MSRP": "D9D9D9",
    "Shipping Override": "D9D9D9",
    "Location Reorder Point": "D9D9D9",
    "Last Markdown Date": "D9D9D9",
    "ITEM SID": "D9D9D9",
    "Internal ID": "D9D9D9",
    "EXCLUDE FROM AMAZON": "D9D9D9",
    "EXCLUDE FROM EBAY": "D9D9D9",
    "EXCLUDE FROM MSW.COM": "D9D9D9",
    "Last Purchase Price": "D9D9D9",
    "Amazon Minimum Price": "D9D9D9",
    "AMAZON MAX PRICE": "D9D9D9",
    "FBA STORAGE TYPE": "D9D9D9",
    "Parent": "D9D9D9",
    "SEND TO FBA": "D9D9D9",
    "FBA ONLY": "D9D9D9",
    "MAP RELEASE DATE.1": "D9D9D9",
    "ASIN": "D9D9D9",
    "AMZ SOURCING COST": "D9D9D9",
}

HEADER_FILL_HEX_BY_COLUMN: dict[str, str] = {
    OUTPUT_HEADER_DISPLAY.get(k, k): v for k, v in _RAW_HEADER_FILLS.items()
}

GROUP_SEPARATOR_AFTER_HEADERS: tuple[str, ...] = tuple(
    OUTPUT_HEADER_DISPLAY.get(k, k) for k in GROUP_SEPARATOR_AFTER_UPLOAD_KEYS
)

# Columns written to the export — NetSuite/upload header strings (114 cols); Excel labels may differ.
# Columns absent from the upload are emitted blank; order is preserved.
OUTPUT_COLUMN_ORDER: tuple[str, ...] = (
    "External ID",
    "_FBM UPC",
    "Vendor Name",
    "Netsuite Style Name",
    "Color Free Form",
    "MPN",
    "Vendor Color Code",
    "SIZE",
    "Season Code",
    "STATUS",
    "MAP RELEASE DATE",
    "AMAZON PRICE",
    "OFF MAP",
    "RETAIL PRICE",
    "MSRP PRICE",
    "AMAZON LINK",
    "SELLER CENTRAL LINK",
    "NETSUITE LINK",
    "MASTER PARENT ASIN",
    "MASTER CHILD ASIN",
    "Parent ASIN 4.30.26",
    "Parent Asin Match 4.30.26",
    "ASIN 4.30.26",
    " Buy Box: Is FBA",
    " Buy Box Seller",
    " Buy Box: Stock",
    " Amazon: Current BUY BOX  Price",
    " Buy Box Eligible Offer Counts: New FBA",
    " Buy Box Eligible Offer Counts: New FBM",
    " Amazon: Stock",
    " Bought in past month",
    " ACTIVE/INACTIVE",
    "FBM QTY",
    "FBM PRICE",
    " CURRENT SHIP TEMP (MIG-STANDARD, SHIP-ADD 9.99)",
    " ACTIVE/INACTIVE.1",
    "AMZ PRICE",
    "AMZ TAG",
    "repricing_enabled FBA ",
    "repricing_enabled MFN ",
    "repricing_enabled FBM ",
    "offer_position FBA ",
    "offer_position MFN ",
    "offer_position FBM ",
    "min_price FBA ",
    "min_price MFN ",
    "min_price FBM ",
    "max_price FBA ",
    "max_price MFN ",
    "max_price FBM ",
    "current_price FBA ",
    "current_price MFN ",
    "current_price FBM ",
    "current_shipping FBA ",
    "current_shipping MFN ",
    "current_shipping FBM ",
    "featured FBA ",
    "featured MFN ",
    "featured FBM ",
    "is_buy_box_winner FBA ",
    "is_buy_box_winner MFN ",
    "is_buy_box_winner FBM ",
    "has_buy_box FBA ",
    "has_buy_box MFN ",
    "has_buy_box FBM ",
    "STORE SALES 4.20-4.26.26",
    "AMZ SALES 4.19-4.25.26",
    " Total Days of Supply (including units from open shipments) ",
    " Recommended replenishment qty",
    "STORE INV ",
    "STORE ON ORDER ",
    "STORE MINS ",
    "FC INV ",
    "FC ON ORDER ",
    " inv-age-0-to-90-days",
    " inv-age-91-to-180-days",
    " inv-age-181-to-270-days",
    " inv-age-271-to-365-days",
    " inv-age-365-plus-days",
    "Recomened Removel+",
    " afn-total-quantity",
    " afn-fulfillable-quantity",
    " afn-unsellable-quantity",
    " afn-reserved-quantity",
    "Reserved Customer Order",
    " afn-inbound-working-quantity",
    " afn-inbound-shipped-quantity",
    " afn-inbound-receiving-quantity",
    " afn-researching-quantity (inbound inv.not avail. for sale)",
    "UPC Code",
    "Price Level",
    "Unit Price",
    "Current Price",
    "Average Cost",
    "Purchase Price",
    "MSRP",
    "Shipping Override",
    "Location Reorder Point",
    "Last Markdown Date",
    "ITEM SID",
    "Internal ID",
    "EXCLUDE FROM AMAZON",
    "EXCLUDE FROM EBAY",
    "EXCLUDE FROM MSW.COM",
    "Last Purchase Price",
    "Amazon Minimum Price",
    "AMAZON MAX PRICE",
    "FBA STORAGE TYPE",
    "Parent",
    "SEND TO FBA",
    "FBA ONLY",
    "MAP RELEASE DATE.1",
    "ASIN",
    "AMZ SOURCING COST",
)

# First match wins: coalesce legacy duplicate NetSuite field names into one output column.
OUTPUT_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    " afn-total-quantity": ("12.15 afn-total-quantity", " afn-total-quantity"),
    " afn-fulfillable-quantity": (" afn-fulfillable-quantity", "afn-fulfillable-quantity"),
    " afn-unsellable-quantity": (" afn-unsellable-quantity", "afn-unsellable-quantity"),
    " afn-reserved-quantity": (" afn-reserved-quantity", "afn-reserved-quantity"),
    " afn-inbound-working-quantity": (" afn-inbound-working-quantity", "afn-inbound-working-quantity"),
    " afn-inbound-shipped-quantity": (" afn-inbound-shipped-quantity", "afn-inbound-shipped-quantity"),
    " afn-inbound-receiving-quantity": (" afn-inbound-receiving-quantity", "afn-inbound-receiving-quantity"),
    " afn-researching-quantity (inbound inv.not avail. for sale)": (
        " afn-researching-quantity (inbound inv.not avail. for sale)",
        "afn-researching-quantity (inbound inv.not avail. for sale)",
    ),
}
