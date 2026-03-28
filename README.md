# Marketing Analytics Assistant

A local, privacy-first Streamlit app for marketing strategists to explore Adobe Analytics Excel exports using plain-English questions and AI-generated PowerPoint insights — powered by Microsoft Phi-3 running entirely on your machine via Ollama.

---

## What it does

| Step | Feature |
|------|---------|
| **1. Load** | Upload any Adobe Analytics Excel export. Auto-detects multiple sheets, named Excel Tables, and multiple tables per sheet separated by blank rows. |
| **2. Query** | Ask questions in plain English. Phi-3 writes the pandas code, runs it, and shows you the result — with an auto-chart for small result sets. |
| **3. Present** | One click generates five structured PowerPoint slide suggestions with key messages, bullet points, and chart type recommendations. |

---

## Requirements

- Python 3.10+
- [Ollama](https://ollama.com) installed and running locally
- Microsoft Phi-3 model pulled via Ollama

---

## Installation

```bash
# 1. Clone or copy the project folder
cd marketing_analytics

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Install Ollama (if not already installed)
# macOS:
brew install ollama
# or download from https://ollama.com/download

# 4. Start the Ollama server
ollama serve

# 5. Pull the Phi-3 model (one-time download, ~2.3 GB)
ollama pull phi3

# 6. Launch the app
streamlit run app.py
```

The app opens at `http://localhost:8501` in your browser.

---

## Project structure

```
marketing_analytics/
├── app.py                        # Main Streamlit application
├── requirements.txt              # Python dependencies
├── generate_sample_data.py       # Generic sample data generator
├── generate_auto_data.py         # Indian auto industry sample data generator
├── adobe_analytics_sample.xlsx   # Generic sample Excel (marketing/e-commerce)
└── varuna_motors_analytics.xlsx  # Indian auto sample Excel (7 sheets, 3,010 rows)
```

---

## Sample data

Two ready-to-use Excel files are included:

### `adobe_analytics_sample.xlsx`
Generic e-commerce brand data — 5 sheets, ~2,400 rows.

| Sheet | Contents |
|-------|----------|
| Campaign Performance | Daily campaign metrics across brands, channels, states |
| Geo Traffic | Weekly traffic by US state and brand |
| Page Behavior | Page-level engagement and conversion metrics |
| CRM Segments | Monthly customer segment data |
| Conversion Funnel | 8-step funnel from visit to order |

### `varuna_motors_analytics.xlsx`
Indian car manufacturer (Varuna Motors) digital booking funnel — 7 sheets, 3,010 rows.

| Sheet | Contents |
|-------|----------|
| Campaign Performance | Daily digital campaign data across 6 car models, 24 states, 12 channels |
| Booking Funnel | 13-step funnel from website visit to confirmed booking |
| Geo & City Performance | Weekly traffic and bookings by Indian state and city |
| Website Page Behavior | Page-level data for all key pages (model pages, EMI calculator, booking form) |
| CRM & Lead Quality | Monthly lead pipeline by buyer segment (First-Time, EV Enthusiast, Fleet, etc.) |
| Model Comparison | Weekly head-to-head across all 6 models |
| Test Drive & Dealer | Monthly test drive pipeline and dealer attribution |

---

## Configuration

All configuration is in the sidebar at runtime — no config files needed.

| Setting | Default | Description |
|---------|---------|-------------|
| Ollama model | `phi3` | Any model available in your local Ollama install |
| Ollama endpoint | `http://localhost:11434` | Change in `app.py` → `OLLAMA_BASE` if running Ollama remotely |

---

## How the AI query engine works

1. The selected table's schema (column names, types, min/max/mean/std for numeric columns) is sent to Phi-3 along with your question.
2. Phi-3 generates a pandas code snippet.
3. The code runs in an isolated namespace (`df`, `pd`, `np` available).
4. If execution fails, the error is fed back to the model for self-correction — up to **3 attempts** total.
5. The result (DataFrame, scalar, or string) is displayed, with an optional bar chart.

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI |
| `pandas` | Data manipulation |
| `openpyxl` | Excel parsing (xlsx) |
| `xlrd` | Excel parsing (xls) |
| `requests` | Ollama API calls |

---

## Privacy

All processing is local. No data is sent to any external API. Phi-3 runs entirely on your machine via Ollama.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Ollama not reachable` | Run `ollama serve` in a terminal and keep it open |
| `No models found` | Run `ollama pull phi3` |
| Query returns an error after 3 attempts | Rephrase using exact column names shown in the Data Preview |
| Excel file shows 0 tables | Ensure the file has a header row followed by at least one data row |
| Slow responses | Phi-3 Mini (`phi3:mini`) is faster; Phi-3 Medium (`phi3:medium`) is more accurate |
