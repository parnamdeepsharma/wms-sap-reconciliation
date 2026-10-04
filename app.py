
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="WMS ↔ SAP Reconciliation", page_icon="📊", layout="wide")

st.title("📊 WMS ↔ SAP Inventory Reconciliation")
st.caption("Download the required templates, fill them with your data, upload them, and generate one reconciliation Excel report.")

# -----------------------------
# Standard templates
# -----------------------------
WMS_COLUMNS = ["Material", "Plant", "Storage Location", "Batch", "WMS Quantity"]
SAP_COLUMNS = ["Material", "Plant", "Storage Location", "Batch", "SAP Quantity"]

def create_template(columns, sheet_name):
    df = pd.DataFrame(columns=columns)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(18, len(str(col[0].value)) + 3)
    out.seek(0)
    return out

def read_and_validate(uploaded, required_columns):
    df = pd.read_excel(uploaded)
    # Normalize header whitespace
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in required_columns if c not in df.columns]
    return df, missing

def clean_text(s):
    return s.fillna("").astype(str).str.strip().str.upper()

def reconcile(wms, sap):
    w = wms.copy()
    s = sap.copy()

    # Standardized matching key
    for df in (w, s):
        for col in ["Material", "Plant", "Storage Location", "Batch"]:
            df[col] = clean_text(df[col])

    w["WMS Quantity"] = pd.to_numeric(w["WMS Quantity"], errors="coerce").fillna(0)
    s["SAP Quantity"] = pd.to_numeric(s["SAP Quantity"], errors="coerce").fillna(0)

    key_cols = ["Material", "Plant", "Storage Location", "Batch"]

    w_agg = w.groupby(key_cols, dropna=False)["WMS Quantity"].sum().reset_index()
    s_agg = s.groupby(key_cols, dropna=False)["SAP Quantity"].sum().reset_index()

    result = pd.merge(
        w_agg,
        s_agg,
        on=key_cols,
        how="outer",
        indicator=True
    )

    result["WMS Quantity"] = result["WMS Quantity"].fillna(0)
    result["SAP Quantity"] = result["SAP Quantity"].fillna(0)
    result["Variance"] = result["WMS Quantity"] - result["SAP Quantity"]

    result["Status"] = result.apply(
        lambda r:
            "MATCH" if r["_merge"] == "both" and r["Variance"] == 0
            else "QTY MISMATCH" if r["_merge"] == "both"
            else "WMS ONLY" if r["_merge"] == "left_only"
            else "SAP ONLY",
        axis=1
    )

    result = result.drop(columns=["_merge"])
    result = result.sort_values(["Status", "Material"])

    summary = pd.DataFrame({
        "Metric": [
            "WMS line count",
            "SAP line count",
            "WMS total quantity",
            "SAP total quantity",
            "Unique WMS material/location records",
            "Unique SAP material/location records",
            "Correct / MATCH records",
            "Quantity mismatch records",
            "WMS-only records",
            "SAP-only records",
            "Total exception records",
        ],
        "Value": [
            len(wms),
            len(sap),
            w["WMS Quantity"].sum(),
            s["SAP Quantity"].sum(),
            len(w_agg),
            len(s_agg),
            int((result["Status"] == "MATCH").sum()),
            int((result["Status"] == "QTY MISMATCH").sum()),
            int((result["Status"] == "WMS ONLY").sum()),
            int((result["Status"] == "SAP ONLY").sum()),
            int((result["Status"] != "MATCH").sum()),
        ]
    })

    return result, summary

def make_excel(result, summary):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Reconciliation")
        summary.to_excel(writer, index=False, sheet_name="Summary")

        for sheet_name in ["Reconciliation", "Summary"]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 14), 35)

    out.seek(0)
    return out

# -----------------------------
# Step 1: Templates
# -----------------------------
st.header("Step 1 — Download the required Excel formats")
st.write("Do not create your own format. Download the correct template, enter your data under the provided headings, and upload it below.")

c1, c2 = st.columns(2)

with c1:
    st.subheader("📦 WMS Template")
    st.write("Required columns: Material, Plant, Storage Location, Batch, WMS Quantity")
    st.download_button(
        "⬇️ Download WMS Excel Format",
        data=create_template(WMS_COLUMNS, "WMS Data"),
        file_name="WMS_Upload_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with c2:
    st.subheader("🏢 SAP Template")
    st.write("Required columns: Material, Plant, Storage Location, Batch, SAP Quantity")
    st.download_button(
        "⬇️ Download SAP Excel Format",
        data=create_template(SAP_COLUMNS, "SAP Data"),
        file_name="SAP_Upload_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# -----------------------------
# Step 2: Upload
# -----------------------------
st.header("Step 2 — Upload completed templates")

wms_file = st.file_uploader("Upload completed WMS template", type=["xlsx"])
sap_file = st.file_uploader("Upload completed SAP template", type=["xlsx"])

if wms_file and sap_file:
    try:
        wms, wms_missing = read_and_validate(wms_file, WMS_COLUMNS)
        sap, sap_missing = read_and_validate(sap_file, SAP_COLUMNS)

        if wms_missing:
            st.error("WMS file has incorrect/missing columns: " + ", ".join(wms_missing))
            st.stop()

        if sap_missing:
            st.error("SAP file has incorrect/missing columns: " + ", ".join(sap_missing))
            st.stop()

        st.success("✅ Both files have the correct format.")

        # Basic data quality checks
        warnings = []
        for name, df, qty_col in [
            ("WMS", wms, "WMS Quantity"),
            ("SAP", sap, "SAP Quantity")
        ]:
            if df.empty:
                warnings.append(f"{name} file contains no data rows.")
            if df["Material"].astype(str).str.strip().eq("").all():
                warnings.append(f"{name}: Material column has no usable values.")
            if pd.to_numeric(df[qty_col], errors="coerce").isna().any():
                warnings.append(f"{name}: some quantity values are not numeric and will be treated as 0.")

        for warning in warnings:
            st.warning(warning)

        st.header("Step 3 — Reconcile")

        if st.button("🔄 Reconcile WMS vs SAP", type="primary", use_container_width=True):
            result, summary = reconcile(wms, sap)

            st.subheader("Summary")

            metrics = [
                ("WMS Lines", len(wms)),
                ("SAP Lines", len(sap)),
                ("MATCH", int((result["Status"] == "MATCH").sum())),
                ("Exceptions", int((result["Status"] != "MATCH").sum())),
            ]

            cols = st.columns(4)
            for col, (label, value) in zip(cols, metrics):
                col.metric(label, value)

            st.dataframe(summary, use_container_width=True)

            st.subheader("Reconciliation Preview")
            st.dataframe(result.head(200), use_container_width=True)

            excel = make_excel(result, summary)

            st.download_button(
                "⬇️ Download Final Reconciliation Excel",
                data=excel,
                file_name="WMS_SAP_Reconciliation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Could not process the files: {e}")
else:
    st.info("Upload both completed templates to continue.")
