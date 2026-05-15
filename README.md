# T.A.N.G.S
## Transformative AI-Powered NLP & Generation Suite

**SAKEC Mumbai — Final Year B.Tech Project 2025-26**  
*ECS Department · Powered by Groq LLaMA-3 · pdfplumber · pytesseract*

---

## What it does

T.A.N.G.S is a **query-driven document processing app** that:

1. **Extracts** tables and text from academic result PDFs, timetable PDFs,
   and financial report PDFs using a 3-tier hybrid engine
   (pdfplumber → regex → OCR with auto-rotation)
2. **Answers** natural-language queries about the document using
   Groq's free LLaMA-3 API
3. **Generates** styled Excel workbooks (3 sheets) and PowerPoint
   presentations (with embedded charts) as downloadable reports

---

## Supported PDFs

| Document Type | Example | Queries |
|---------------|---------|---------|
| Academic Results | SAKEC B.Tech result sheet (ESE marks) | "Who topped?", "List KT students", "Pass % in DBMS?" |
| Class Timetables | SAKEC semester timetable | "Workload of /MK?", "Free slots Friday?", "Which room?" |
| Financial Reports | Annual reports, P&L sheets | "What is revenue?", "Net profit?", "YoY growth?" |

---

## Local Setup (5 minutes)

### 1. Clone / download

```bash
git clone https://github.com/<your-username>/tangs.git
cd tangs
```

### 2. Create a virtual environment (recommended)

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
```

### 3. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 4. Install system dependencies

#### Tesseract OCR (required for scanned / rotated PDFs)

- **Windows:**  
  Download installer from  
  [github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)  
  Install to `C:\Program Files\Tesseract-OCR\`  
  Add that folder to your PATH environment variable.

- **macOS:**  
  ```bash
  brew install tesseract
  ```

- **Linux (Ubuntu/Debian):**  
  ```bash
  sudo apt update && sudo apt install tesseract-ocr -y
  ```

#### Poppler (required for pdf2image PDF-to-image conversion)

- **Windows:**  
  Download from  
  [github.com/oschwartz10612/poppler-windows/releases](https://github.com/oschwartz10612/poppler-windows/releases)  
  Extract, then add the `bin/` folder to your PATH.

- **macOS:**  
  ```bash
  brew install poppler
  ```

- **Linux (Ubuntu/Debian):**  
  ```bash
  sudo apt update && sudo apt install poppler-utils -y
  ```

### 5. Run the app

```bash
streamlit run app.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser.

---

## Get a Free Groq API Key

1. Go to [console.groq.com](https://console.groq.com)
2. Sign up with Google (free, no credit card)
3. Click **API Keys** → **Create API Key**
4. Copy the key (starts with `gsk_…`)
5. Paste it into the **Groq API Key** field in the app sidebar

**Rate limits (free tier):** ~30 req/min · 6,000 tokens/min — more than
enough for all T.A.N.G.S queries.

---

## Deploy to Streamlit Cloud (FREE, ~5 minutes)

Streamlit Cloud gives you a permanent public URL for free.

### Step-by-step

1. Push this folder to a **new public GitHub repository**

2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in

3. Click **New app**

4. Select your repo, branch = `main`, main file = `app.py`

5. Click **Advanced settings** → **Secrets** and add:
   ```toml
   GROQ_API_KEY = "gsk_your_key_here"
   ```
   This pre-fills the API key so users don't need to enter it.

6. Click **Deploy** — your app is live in ~3 minutes at  
   `https://<your-app-name>.streamlit.app`

> **Note:** `packages.txt` (containing `poppler-utils` and `tesseract-ocr`)
> is picked up automatically by Streamlit Cloud to install system packages
> before Python dependencies.

---

## Project Structure

```
tangs/
├── app.py                  ← Streamlit frontend (this file)
├── models.py               ← Dataclasses (StudentRecord, etc.)
├── pdf_extractor.py        ← 3-tier hybrid PDF extractor
├── academic_module.py      ← Result parsing, toppers, KT, stats
├── timetable_module.py     ← Timetable parsing, workload, schedule
├── financial_module.py     ← KPI extraction, YoY growth
├── query_router.py         ← Intent routing + TF-IDF fallback
├── output_generator.py     ← Excel, chart, and PowerPoint generators
├── requirements.txt        ← Python dependencies
├── packages.txt            ← System packages (Streamlit Cloud)
├── build_zip.sh            ← One-command ZIP for submission
├── README.md               ← This file
└── .streamlit/
    └── config.toml         ← Theme and server settings
```

---

## Extraction Engine

| Tier | Method | When Used |
|------|--------|-----------|
| **Tier 1** | `pdfplumber.extract_tables()` | Clean digital PDFs with table structure |
| **Tier 2** | Raw text + regex patterns | Semi-structured or text-heavy PDFs |
| **Tier 3** | `pdf2image` + `pytesseract` OCR with auto-rotation | Scanned or image-based PDFs |

Column headers are fuzzy-normalized with `fuzzywuzzy` (ratio ≥ 70).  
Teacher legend codes (`/MK: Manjusha Kulkarni`) are resolved automatically.

---

## Architecture

```
Upload PDF
    │
    ▼
HybridExtractor (pdf_extractor.py)
    │   Tier 1: pdfplumber tables
    │   Tier 2: regex on raw text
    │   Tier 3: OCR + auto-rotation
    ▼
detect_doc_type()  ─────────────────────────────────────────────┐
    │                                                           │
    ▼                                                           │
route_query()  ──► topper / kt / workload / kpi / llm_fallback  │
    │                                                           │
    ▼                                                           │
process_result / process_timetable / process_financial          │
    │   ├── Parse DataFrame                                     │
    │   ├── Analytics (toppers, KT, failure rates, workload)    │
    │   └── LLM answer via Groq llama3-8b-8192                 │
    ▼                                                           │
ExcelGenerator   ──► 3-sheet .xlsx (Raw / Summary / Result)     │
ChartGenerator   ──► PNG charts (bar, pie, line, failure rate)  │
PowerPointGenerator ► Themed .pptx with embedded charts         │
    │                                                           │
    ▼                                                           │
Streamlit UI  ◄─────────────────────────────────────────────────┘
```

---

## Academic Context

- **Institute:** Shah & Anchor Kutchhi Engineering College (SAKEC), Mumbai
- **Programme:** B.Tech — Electronics & Computer Science (ECS)
- **Batch:** 2021-25
- **Project Type:** Final Year Project (FYP)
- **Guide:** [Professor Name]

### Target PDFs

| File | Type | Pages | Notes |
|------|------|-------|-------|
| `FE_time_table.pdf` | Digital timetable | 1 | FE Sem-II, clean, digital |
| `Btech2.pdf` | Scanned timetable | 1 | BE Sem-VIII, rotated, OCR needed |
| `B_TECH_CYSE_MAY_2025_SEM-II.pdf` | Result sheet | 3 | 63 students, ESE/OR/GR/CR/GP |

---

*T.A.N.G.S — Making academic document analysis intelligent, one query at a time.*
