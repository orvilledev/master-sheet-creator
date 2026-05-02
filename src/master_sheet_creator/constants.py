"""Shared constants for file handling and UI."""

UPLOAD_EXTENSIONS = {".csv", ".xlsx", ".xls"}
STREAMLIT_UPLOAD_TYPES = ["csv", "xlsx", "xls"]

# UTF-8 with BOM helps Excel on Windows recognize UTF-8 CSV correctly.
CSV_ENCODING = "utf-8-sig"

# Columns written to the export (NetSuite-style upload → ``format.xlsx`` alignment).
# Order matches the template; columns absent from the upload are emitted empty.
OUTPUT_COLUMN_ORDER: tuple[str, ...] = (
    "External ID",
    "Vendor Name",
    "Netsuite Style Name",
    "Color Free Form",
    "MPN",
    "Vendor Color Code",
    "SIZE",
    "Season Code",
    "STATUS",
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
    "MAP RELEASE DATE",
    "ASIN",
    "AMZ SOURCING COST",
)
