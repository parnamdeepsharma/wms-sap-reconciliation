# WMS ↔ SAP Inventory Reconciliation MVP

## Updated workflow

The website now forces a standard upload format.

### WMS template
- Material
- Plant
- Storage Location
- Batch
- WMS Quantity

### SAP template
- Material
- Plant
- Storage Location
- Batch
- SAP Quantity

Users first download the correct template, fill in their data, then upload the completed files.

The application validates the columns automatically and rejects files with missing/incorrect columns.

The final workbook contains:
1. Reconciliation
2. Summary

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```
