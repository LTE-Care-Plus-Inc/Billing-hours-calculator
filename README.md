# Billable Hours Aggregator — Streamlit

This project includes the original Tkinter GUI and a Streamlit web app. The instructions below focus on running the Streamlit app (no EXE packaging).

Highlights
- Upload CSV/Excel or use the bundled `sample_data.xlsx`
- Search and sort staff summary
- Per-staff per-day breakdown with Appt Start/End
- Download filtered `summary.csv`, full `details.csv`, combined `results.xlsx`
- Sidebar theme toggle (Light/Dark)

Run (Streamlit)
1) `pip install -r requirements.txt`
2) `streamlit run streamlit_app.py`

Theme
- Default theme is defined in `.streamlit/config.toml` (light base).
- A sidebar toggle applies a simple light/dark CSS override at runtime.
