# 🧾 Tax Receipt AI Agent

AI-powered tool that processes receipts and invoices using Gemini Vision,
and generates an Excel report ready for the Israeli tax authority (דף מס' 1303).

![App screenshot](docs/screenshot.png)

## Features
- 📤 Separate upload for expenses (חשבוניות) and income (קבלות)
- 🤖 Gemini Vision extracts: supplier, amount, VAT, date, category
- 📊 Applies Israeli tax rules (50% entertainment/vehicle, 100% others)
- ⚠️ Flags ambiguous documents for manual review
- ⬇️ Exports formatted Excel report (3 sheets)

## Tech Stack
- Gemini 2.5 Flash (Vision + extraction)
- Streamlit (UI)
- Python + openpyxl

## How to Run (Google Colab)
1. Open in Google Colab
2. Add `GEMINI_KEY` to Colab Secrets
3. Run all cells
4. Open the ngrok URL

## How to Run (locally)
```bash
pip install -r requirements.txt
export GEMINI_KEY="your-gemini-api-key"
streamlit run app.py
```
