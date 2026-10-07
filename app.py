import io
import base64
import pandas as pd
import streamlit as st

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="WMS ↔ SAP Reconciliation",
    page_icon="📦",
    layout="wide"
)

# ============================================================
# WAREHOUSE BACKGROUND
# ============================================================

try:
    with open("warehouse.png", "rb") as f:
        warehouse_img = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <style>

        .warehouse-background {{
            position: fixed;
            right: 0;
            top: 0;
            width: 50%;
            height: 100%;
            background-image: url("data:image/png;base64,{warehouse_img}");
            background-size: cover;
            background-position: center;
            opacity: 0.18;
            z-index: -1;
        }}

        .parnam-signature {{
            position: fixed;
            right: 25px;
            bottom: 20px;
            font-family: cursive;
            font-size: 22px;
            font-style: italic;
            font-weight: 600;
            opacity: 0.65;
            z-index: 10;
        }}

        </style>

        <div class="warehouse-background"></div>
        <div class="parnam-signature">PARNAM</div>
        """,
        unsafe_allow_html=True
    )

except Exception:
    pass


# ============================================================
# PAGE TITLE
# ============================================================

st.title("📊 WMS ↔ SAP Inventory Reconciliation")

st.caption(
    "Download the required templates, fill them with your data, "
    "upload them, and generate one reconciliation Excel report."
)


# ============================================================
# STANDARD EXCEL TEMPLATES
# ============================================================

WMS_COLUMNS = [
    "Material",
    "Batch",
    "WMS Quantity",
    "Plant"
]

SAP_COLUMNS = [
    "Material",
    "Batch",
    "SAP Quantity",
    "Plant"
]


# ============================================================
# CREATE EXCEL TEMPLATE
# ============================================================

def create_template(columns, sheet_name):

    df = pd.DataFrame(columns=columns)

    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

        ws = writer.book[sheet_name]

        # Freeze header
        ws.freeze_panes = "A2"
