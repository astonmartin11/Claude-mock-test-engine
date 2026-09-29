# Adaptive AI Mock Test Engine & Tutor
1. `pip install -r requirements.txt` (local OCR also needs the `tesseract-ocr` system package; Streamlit Cloud reads packages.txt).
2. Copy `.env.example` values into `.streamlit/secrets.toml` (as `KEY = "value"`) or environment variables. Needs a Neon `DATABASE_URL`, Gemini and Groq keys.
3. `streamlit run app.py` — the schema is created automatically on first start. The Settings page has diagnostics.
