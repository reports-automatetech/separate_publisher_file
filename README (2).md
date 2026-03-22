# ✂️ Split by Publisher

A Streamlit web app that splits a multi-sheet Excel file into **one Excel file per publisher**, then packages everything into a ZIP for download — no coding required.

---

## 🚀 Live App

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-name.streamlit.app)

> Replace the link above with your deployed Streamlit URL after deployment.

---

## 📋 How It Works

1. Upload any `.xlsx` file that contains a `publishername` column
2. The app scans every sheet, detects which ones have that column, and loads the data
3. It groups rows by unique publisher name
4. Click **Split & Download ZIP** — you get one formatted `.xlsx` per publisher, all in a single ZIP file

Each output file includes:
- Bold, bordered header row
- Auto-sized columns (capped at 50 chars wide)
- Frozen header row
- Auto-filter on all columns
- Correct datetime formatting

---

## 🗂️ Repository Structure

```
├── split_by_publisher_app.py   # Main Streamlit app
├── requirements.txt            # Python dependencies
└── README.md                   # This file
```

---

## 🛠️ Run Locally

**1. Clone the repo**
```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Run the app**
```bash
streamlit run split_by_publisher_app.py
```

The app will open at `http://localhost:8501`

---

## ☁️ Deploy on Streamlit Community Cloud (Free)

1. Push this repo to GitHub (must be **public**, or a private repo on a paid plan)
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
3. Click **"New app"**
4. Fill in the fields:

   | Field | Value |
   |---|---|
   | Repository | `your-username/your-repo-name` |
   | Branch | `main` |
   | Main file path | `split_by_publisher_app.py` |

5. Click **Deploy** — Streamlit will install `requirements.txt` automatically
6. Your app will be live at `https://your-app-name.streamlit.app`

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | Web UI framework |
| `pandas` | Data reading & manipulation |
| `openpyxl` | Reading `.xlsx` files |
| `xlsxwriter` | Writing formatted `.xlsx` output |
| `numpy` | Numerical support for pandas |

---

## ⚙️ Configuration

The `publishername` column name and datetime format are defined as constants at the top of `split_by_publisher_app.py`:

```python
PUBLISHER_COL = "publishername"   # column name to split on (case-insensitive)
DATETIME_FMT  = "yyyy-mm-dd hh:mm:ss"   # Excel datetime display format
```

Change these if your column has a different name or you prefer a different date format.

---

## 📝 Notes

- Column matching is **case-insensitive** — `PublisherName`, `PUBLISHERNAME`, and `publishername` all work
- Sheets without a `publishername` column are silently skipped
- Publishers with no data rows are skipped (no empty files generated)
- Special characters (`/ \ :`) in publisher names are replaced with `_` in filenames
- All processing is done **in-memory** — no temporary files are written to disk
