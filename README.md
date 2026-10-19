# receipt-reader

I had a pile of last year's receipts I needed to get into Zoho Books. Zoho has a decent import format but you still need the data first — so I was typing them into a spreadsheet one by one. Got through maybe 50 before I thought, there has to be a better way.

The obvious move was to throw the photos at ChatGPT or Claude, but those hit usage limits fast when you're working through a backlog. So I built a small Streamlit app that runs a vision model locally through Ollama. Take a photo of a receipt (or upload one), it reads the values, you approve or fix whatever it got wrong, and it appends everything to an Excel workbook.

I kept coming back and tweaking it with each entry — added a memory system so it learns from your corrections, duplicate detection, arithmetic validation, auto-save when confidence is high enough. Turned into more of a proper tool than I expected.

The default model is `qwen2.5vl:3b`. Small enough to run on most Macs, reads receipts well enough for day-to-day use. No data leaves your machine.

Don't know who else might find it useful, but here it is.

---

**Stack:** Python · Streamlit · Ollama · openpyxl · Pillow · pandas

---

## What it produces

An `.xlsx` workbook with two main sheets:

- **Receipts** — one row per receipt (merchant, date, total, category, payment method, etc.)
- **Items** — one row per line item extracted from each receipt

Plus a sheet per category (Food, Fuel, Health, etc.) for quick filtering.

---

## Requirements

- macOS (the camera widget is browser-based; should work on other platforms but untested)
- Python 3.11+
- [Ollama for macOS](https://ollama.com/download/mac)
- A vision model pulled in Ollama — default is `qwen2.5vl:3b`

---

## Setup

**1. Install Ollama and pull the model**

Download and open Ollama, then:

```bash
ollama pull qwen2.5vl:3b
```

This downloads the model once. Everything after that runs offline. Verify it's there with `ollama list`.

If `qwen2.5vl:3b` is too slow on your machine or doesn't read your receipts well, try a larger model and swap the name in the app sidebar.

**2. Clone and set up the environment**

```bash
git clone https://github.com/YOUR_USERNAME/receipt-reader.git
cd receipt-reader
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**3. Run it**

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`. Allow camera access when the browser asks.

---

## Using it

1. Hold a receipt in decent light with the full thing in frame
2. Capture it with the camera widget or upload a photo
3. The app reads it with the local model and shows what it extracted
4. If confidence and validation checks pass, it saves automatically
5. Otherwise it holds it for review — fix anything wrong and hit **Approve and save to Excel**

The workbook lands at `~/ReceiptReader/receipts.xlsx` by default. You can change the path in the sidebar.

The feedback section (under each extracted receipt) lets you teach it things like which bank name to ignore or what a specific merchant is actually called. Those corrections get injected into every future extraction.

---

## Environment variables

| Variable | Default | What it does |
|---|---|---|
| `OLLAMA_URL` | `http://localhost:11434` | Ollama API endpoint |
| `OLLAMA_MODEL` | `qwen2.5vl:3b` | Model to use for extraction |
| `RECEIPT_READER_XLSX` | `~/ReceiptReader/receipts.xlsx` | Workbook path |

---

## Troubleshooting

**Can't connect to Ollama**
Open the Ollama app. Check that the model name in the sidebar matches what `ollama list` shows.

**Camera doesn't appear**
Use `http://localhost:8501` (not `https`), allow camera in the browser, and check macOS System Settings → Privacy & Security → Camera.

**Excel write error**
Close the workbook in Excel before the app tries to write to it.

**Poor extraction quality**
Flatten the receipt, use even lighting, avoid glare. If it's still bad, try a larger model like `qwen2.5vl:7b`.
