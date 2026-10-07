
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

        # Bold headers
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)

        # Column widths
        for col in ws.columns:

            ws.column_dimensions[
                col[0].column_letter
            ].width = max(
                18,
                len(str(col[0].value)) + 3
            )

    out.seek(0)

    return out


# ============================================================
# READ AND VALIDATE EXCEL FILE
# ============================================================

def read_and_validate(uploaded, required_columns):

    df = pd.read_excel(uploaded)

    # Remove spaces from column names
    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # Find missing columns
    missing = [
        c
        for c in required_columns
        if c not in df.columns
    ]

    return df, missing


# ============================================================
# CLEAN TEXT DATA
# ============================================================

def clean_text(s):

    return (
        s.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )


# ============================================================
# RECONCILIATION ENGINE
# ============================================================

def reconcile(wms, sap):

    # Make copies so original uploaded data is not modified
    w = wms.copy()
    s = sap.copy()

    # --------------------------------------------------------
    # STANDARDIZE MATCHING COLUMNS
    # --------------------------------------------------------

    for df in (w, s):

        for col in [
            "Material",
            "Plant",
            "Batch"
        ]:

            df[col] = clean_text(
                df[col]
            )

    # --------------------------------------------------------
    # CONVERT QUANTITIES TO NUMERIC
    # --------------------------------------------------------

    w["WMS Quantity"] = pd.to_numeric(
        w["WMS Quantity"],
        errors="coerce"
    ).fillna(0)

    s["SAP Quantity"] = pd.to_numeric(
        s["SAP Quantity"],
        errors="coerce"
    ).fillna(0)

    # --------------------------------------------------------
    # MATCHING KEY
    # --------------------------------------------------------

    key_cols = [
        "Material",
        "Plant",
        "Batch"
    ]

    # --------------------------------------------------------
    # AGGREGATE WMS DATA
    # --------------------------------------------------------

    w_agg = (
        w.groupby(
            key_cols,
            dropna=False
        )["WMS Quantity"]
        .sum()
        .reset_index()
    )

    # --------------------------------------------------------
    # AGGREGATE SAP DATA
    # --------------------------------------------------------

    s_agg = (
        s.groupby(
            key_cols,
            dropna=False
        )["SAP Quantity"]
        .sum()
        .reset_index()
    )

    # --------------------------------------------------------
    # MERGE WMS AND SAP
    # --------------------------------------------------------

    result = pd.merge(
        w_agg,
        s_agg,
        on=key_cols,
        how="outer",
        indicator=True
    )

    # --------------------------------------------------------
    # REPLACE MISSING QUANTITIES WITH ZERO
    # --------------------------------------------------------

    result["WMS Quantity"] = (
        result["WMS Quantity"]
        .fillna(0)
    )

    result["SAP Quantity"] = (
        result["SAP Quantity"]
        .fillna(0)
    )

    # --------------------------------------------------------
    # CALCULATE VARIANCE
    # --------------------------------------------------------

    result["Variance"] = (
        result["WMS Quantity"]
        -
        result["SAP Quantity"]
    )

    # --------------------------------------------------------
    # DETERMINE STATUS
    # --------------------------------------------------------

    def determine_status(row):

        if (
            row["_merge"] == "both"
            and row["Variance"] == 0
        ):
            return "MATCH"

        elif (
            row["_merge"] == "both"
        ):
            return "QTY MISMATCH"

        elif (
            row["_merge"] == "left_only"
        ):
            return "WMS ONLY"

        else:
            return "SAP ONLY"

    result["Status"] = result.apply(
        determine_status,
        axis=1
    )

    # --------------------------------------------------------
    # REMOVE MERGE HELPER COLUMN
    # --------------------------------------------------------

    result = result.drop(
        columns=["_merge"]
    )

    # --------------------------------------------------------
    # SORT RESULT
    # --------------------------------------------------------

    result = result.sort_values(
        ["Status", "Material"]
    ).reset_index(drop=True)

    # ========================================================
    # SUMMARY COUNTS
    # ========================================================

    match_count = int(
        (
            result["Status"]
            ==
            "MATCH"
        ).sum()
    )

    mismatch_count = int(
        (
            result["Status"]
            ==
            "QTY MISMATCH"
        ).sum()
    )

    wms_only_count = int(
        (
            result["Status"]
            ==
            "WMS ONLY"
        ).sum()
    )

    sap_only_count = int(
        (
            result["Status"]
            ==
            "SAP ONLY"
        ).sum()
    )

    exception_count = int(
        (
            result["Status"]
            !=
            "MATCH"
        ).sum()
    )

    # ========================================================
    # SUMMARY TABLE
    # ========================================================

    summary = pd.DataFrame({

        "Metric": [

            "WMS line count",

            "SAP line count",

            "WMS total quantity",

            "SAP total quantity",

            "Correct / MATCH records",

            "Quantity mismatch records",

            "WMS-only records",

            "SAP-only records",

            "Total exception records"

        ],

        "Value": [

            len(wms),

            len(sap),

            w["WMS Quantity"].sum(),

            s["SAP Quantity"].sum(),

            match_count,

            mismatch_count,

            wms_only_count,

            sap_only_count,

            exception_count

        ]

    })

    return result, summary


# ============================================================
# CREATE FINAL EXCEL REPORT
# ============================================================

def make_excel(result, summary):

    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:

        # Reconciliation sheet
        result.to_excel(
            writer,
            index=False,
            sheet_name="Reconciliation"
        )

        # Summary sheet
        summary.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

        # ----------------------------------------------------
        # FORMAT BOTH SHEETS
        # ----------------------------------------------------

        for sheet_name in [
            "Reconciliation",
            "Summary"
        ]:

            ws = writer.book[sheet_name]

            # Freeze first row
            ws.freeze_panes = "A2"

            # Bold headers
            for cell in ws[1]:
                cell.font = cell.font.copy(
                    bold=True
                )

            # Automatic column widths
            for col in ws.columns:

                max_len = max(
                    len(str(c.value))
                    if c.value is not None
                    else 0
                    for c in col
                )

                ws.column_dimensions[
                    col[0].column_letter
                ].width = min(
                    max(
                        max_len + 2,
                        14
                    ),
                    35
                )

    out.seek(0)

    return out


# ============================================================
# STEP 1 — DOWNLOAD TEMPLATES
# ============================================================

st.header(
    "Step 1 — Download the required Excel formats"
)

st.write(
    "Do not create your own format. "
    "Download the correct template, enter your data "
    "under the provided headings, and upload it below."
)

c1, c2 = st.columns(2)


# ------------------------------------------------------------
# WMS TEMPLATE
# ------------------------------------------------------------

with c1:

    st.subheader(
        "📦 WMS Template"
    )

    st.write(
        "Required columns: "
        "Material, Plant, Batch, WMS Quantity"
    )

    st.download_button(

        "⬇️ Download WMS Excel Format",

        data=create_template(
            WMS_COLUMNS,
            "WMS Data"
        ),

        file_name="WMS_Upload_Template.xlsx",

        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ------------------------------------------------------------
# SAP TEMPLATE
# ------------------------------------------------------------

with c2:

    st.subheader(
        "🏢 SAP Template"
    )

    st.write(
        "Required columns: "
        "Material, Plant, Batch, SAP Quantity"
    )

    st.download_button(

        "⬇️ Download SAP Excel Format",

        data=create_template(
            SAP_COLUMNS,
            "SAP Data"
        ),

        file_name="SAP_Upload_Template.xlsx",

        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


st.divider()


# ============================================================
# STEP 2 — UPLOAD FILES
# ============================================================

st.header(
    "Step 2 — Upload completed templates"
)

wms_file = st.file_uploader(
    "Upload completed WMS template",
    type=["xlsx"]
)

sap_file = st.file_uploader(
    "Upload completed SAP template",
    type=["xlsx"]
)


# ============================================================
# PROCESS FILES
# ============================================================

if wms_file and sap_file:

    try:

        # ----------------------------------------------------
        # READ FILES
        # ----------------------------------------------------

        wms, wms_missing = read_and_validate(
            wms_file,
            WMS_COLUMNS
        )

        sap, sap_missing = read_and_validate(
            sap_file,
            SAP_COLUMNS
        )

        # ----------------------------------------------------
        # VALIDATE WMS COLUMNS
        # ----------------------------------------------------

        if wms_missing:

            st.error(
                "WMS file has incorrect/missing columns: "
                +
                ", ".join(wms_missing)
            )

            st.stop()

        # ----------------------------------------------------
        # VALIDATE SAP COLUMNS
        # ----------------------------------------------------

        if sap_missing:

            st.error(
                "SAP file has incorrect/missing columns: "
                +
                ", ".join(sap_missing)
            )

            st.stop()

        # ----------------------------------------------------
        # SUCCESS MESSAGE
        # ----------------------------------------------------

        st.success(
            "✅ Both files have the correct format."
        )

        # ====================================================
        # BASIC DATA QUALITY CHECKS
        # ====================================================

        warnings = []

        for name, df, qty_col in [

            (
                "WMS",
                wms,
                "WMS Quantity"
            ),

            (
                "SAP",
                sap,
                "SAP Quantity"
            )

        ]:

            # Empty file
            if df.empty:

                warnings.append(
                    f"{name} file contains no data rows."
                )

            # Material validation
            if (
                df["Material"]
                .astype(str)
                .str.strip()
                .eq("")
                .all()
            ):

                warnings.append(
                    f"{name}: Material column has "
                    "no usable values."
                )

            # Quantity validation
            if (
                pd.to_numeric(
                    df[qty_col],
                    errors="coerce"
                ).isna().any()
            ):

                warnings.append(
                    f"{name}: some quantity values "
                    "are not numeric and will be "
                    "treated as 0."
                )

        # ----------------------------------------------------
        # SHOW WARNINGS
        # ----------------------------------------------------

        for warning in warnings:

            st.warning(warning)


        # ====================================================
        # STEP 3 — RECONCILE
        # ====================================================

        st.header(
            "Step 3 — Reconcile"
        )

        if st.button(
            "🔄 Reconcile WMS vs SAP",
            type="primary",
            use_container_width=True
        ):

            # ------------------------------------------------
            # RUN RECONCILIATION
            # ------------------------------------------------

            result, summary = reconcile(
                wms,
                sap
            )

            # =================================================
            # SUMMARY METRICS
            # =================================================

            st.subheader(
                "Summary"
            )

            metrics = [

                (
                    "WMS Lines",
                    len(wms)
                ),

                (
                    "SAP Lines",
                    len(sap)
                ),

                (
                    "MATCH",
                    int(
                        (
                            result["Status"]
                            ==
                            "MATCH"
                        ).sum()
                    )
                ),

                (
                    "Exceptions",
                    int(
                        (
                            result["Status"]
                            !=
                            "MATCH"
                        ).sum()
                    )
                )

            ]

            cols = st.columns(4)

            for col, (label, value) in zip(
                cols,
                metrics
            ):

                col.metric(
                    label,
                    value
                )

            # =================================================
            # SUMMARY TABLE
            # =================================================

            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True
            )

            # =================================================
            # RECONCILIATION PREVIEW
            # =================================================

            st.subheader(
                "Reconciliation Preview"
            )

            st.dataframe(
                result.head(200),
                use_container_width=True,
                hide_index=True
            )

            # =================================================
            # CREATE EXCEL REPORT
            # =================================================

            excel = make_excel(
                result,
                summary
            )

            # =================================================
            # DOWNLOAD BUTTON
            # =================================================

            st.download_button(

                "⬇️ Download Final Reconciliation Excel",

                data=excel,

                file_name=(
                    "WMS_SAP_Reconciliation.xlsx"
                ),

                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                use_container_width=True
            )


    # ========================================================
    # ERROR HANDLING
    # ========================================================

    except Exception as e:

        st.error(
            f"Could not process the files: {e}"
        )


# ============================================================
# WAITING FOR FILES
# ============================================================

else:

    st.info(
        "Upload both completed templates to continue."
    )
        
