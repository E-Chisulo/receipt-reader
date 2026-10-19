# receipt-reader

I had a pile of last year's receipts I needed to get into Zoho Books. Zoho has a decent import format but you still need the data first, so I was typing them into a spreadsheet one by one. Got through maybe 50 before I thought, there has to be a better way.

The obvious move was to throw the photos at ChatGPT or Claude, but those hit usage limits fast when you're working through a backlog. So I built a small Streamlit app that runs a vision model locally through Ollama. Take a photo of a receipt (or upload one), it reads the values, you approve or fix whatever it got wrong, and it appends everything to an Excel workbook.

I kept coming back and tweaking it with each entry. Added a memory system so it learns from your corrections, duplicate detection, arithmetic validation, auto-save when confidence is high enough. Turned into more of a proper tool than I expected.

The default model is `qwen2.5vl:3b`. Small enough to run on most Macs, reads receipts well enough for day-to-day use. No data leaves your machine.

Don't know who else might find it useful, but here it is.

---

If you run a business that still works off paper records or till slips, this could be the missing piece between your physical receipts and whatever accounting system you use. Capture the receipts, export the sheet, import into your system. The categories and line items are all there.

---

**Stack:** Python · Streamlit · Ollama · openpyxl · Pillow · pandas

---

## What it produces

An `.xlsx` workbook with two main sheets:

- **Receipts**: one row per receipt (merchant, date, total, category, payment method, etc.)
- **Items**: one row per line item extracted from each receipt

Plus a sheet per category (Food, Fuel, Health, etc.) for quick filtering.

---

## Requirements

**Platform:** macOS, Windows, or Linux. The camera widget runs in the browser so there's nothing platform-specific about the app itself. Ollama supports all three.

**RAM:**
- 3B model (`qwen2.5vl:3b`, the default): 8GB minimum, 16GB if you want it comfortable
- 7B model (`qwen2.5vl:7b`, better accuracy): 16GB minimum

**Storage:**
- `qwen2.5vl:3b`: about 2.3GB for the model
- `qwen2.5vl:7b`: about 4.7GB
- The app and its Python environment add roughly 400MB on top

**Processor:** Any reasonably modern CPU works. Inference is just slow without a GPU. Apple Silicon (M1 and up) runs it well through Metal. On Windows, an NVIDIA GPU speeds things up a lot but is not required — a decent CPU gets the job done, just takes 10-30 seconds per receipt instead of 2-5.

**Software:**
- Python 3.11+
- [Ollama](https://ollama.com/download) (Mac, Windows, or Linux build)
- A vision model pulled in Ollama, default is `qwen2.5vl:3b`

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
5. Otherwise it holds it for review, fix anything wrong and hit **Approve and save to Excel**

The workbook lands at `~/ReceiptReader/receipts.xlsx` by default. You can change the path in the sidebar.

The feedback section (under each extracted receipt) lets you teach it things like which bank name to ignore or what a specific merchant is actually called. Those corrections get injected into every future extraction.

---

## HTTPS (needed for camera on some setups)

Browsers block camera access on non-localhost origins unless the page is served over HTTPS. This comes up if you're running the app on one machine and opening it from your phone or another device on the same network.

Generate a self-signed certificate:

```bash
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes \
  -subj "/CN=localhost"
```

Then run Streamlit with the cert:

```bash
streamlit run app.py \
  --server.sslCertFile cert.pem \
  --server.sslKeyFile key.pem
```

Open `https://<your-machine-ip>:8501` from the other device. The browser will warn about the self-signed cert; accept it and the camera will work. Keep the `.pem` files out of version control (`.gitignore` already covers them).

---

## Using a hosted model instead of Ollama

The app currently calls Ollama's local API directly. Swapping in a hosted model like GPT-4o or Claude via API key is doable but not built yet. The extraction function (`ollama_extract` in `app.py`) is the only place that would need to change. If your machine can't run a local model well enough, or you want better accuracy and don't mind the API cost, that's the place to start.

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
You're likely on a non-localhost URL without HTTPS. See the HTTPS section above, or just use `http://localhost:8501` directly on the same machine.

**Excel write error**
Close the workbook in Excel before the app tries to write to it.

**Poor extraction quality**
Flatten the receipt, use even lighting, avoid glare. If it's still bad, try a larger model like `qwen2.5vl:7b`.
