# Compound Title Processor — Streamlit App

Internal tool for VMG Repertoire Audit compound title harmonisation.

## What it does

Upload a VMG Repertoire Audit workbook (.xlsm), click Process, download the
result with compound titles filled in and edge cases flagged for review.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit web app |
| `process_compound_titles.py` | Processing logic (do not modify) |
| `requirements.txt` | Python dependencies |

## Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Community Cloud

1. Push all files to a GitHub repository (can be private)
2. Go to share.streamlit.io and sign in with GitHub
3. Click "New app" → select your repository → set main file to `app.py`
4. Click Deploy

The app will be live at a `*.streamlit.app` URL within a few minutes.
Access can be restricted to specific email addresses in the Streamlit Cloud
settings if needed.

## Deploying to Azure App Service

Hand the repository to your developer. They will need to:
- Create an Azure App Service (Python 3.11+ runtime)
- Set startup command: `streamlit run app.py --server.port 8000 --server.address 0.0.0.0`
- Deploy from the GitHub repository via Azure DevOps or GitHub Actions

## Notes

- The workbook must contain an IMPORT sheet
- Data should be refreshed (Data → Refresh All) and saved before uploading
- Processed file is named `originalname_processed.xlsm`
- The REVIEW column flags rows needing human decisions. It is added in the first column after the last used column of the IMPORT sheet (AM in the original template, AP in the Q5 Label Tranche 1 workbook), never over existing data
