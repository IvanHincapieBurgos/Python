# Documentation

## 1. Overview

`logistica.py` is an automation tool for logistics operations that integrates several systems into a single workflow:

- Web interface automation via Selenium.
- Integration with Google Drive and Google Sheets.
- Queries against a Redshift data warehouse.
- PDF extraction and generation.
- AI-powered customs category classification.

The script exposes a graphical user interface (GUI) with credential fields and selectable actions.

---

## 2. Requirements and Credential Files

### Required Credential Files

- `credenciales_drive.json` — Service account credentials for Google Drive and Google Sheets.
- `.env` — Environment variable definitions (optional but recommended).

Both files must be placed in the same directory as `logistica.py`.

### Environment Variables

The script uses `python-dotenv` to load environment variables. The most relevant ones are:

### Python Dependencies

The script checks `requirements.txt` on startup and warns about any missing packages. Key libraries include:

- `selenium`, `webdriver-manager` — Browser automation
- `gspread`, `google-api-python-client`, `google-auth` — Google Sheets and Drive
- `pytesseract`, `pdfplumber`, `Pillow` — PDF and image processing
- `python-dotenv` — Environment variable loading
- `psycopg2`, `pandas` — Redshift connectivity and data handling

---

## 3. User Interface

The GUI is built with `tkinter` and organized into two main panels:

- **Primary Panel** — Main credential fields (username, password, etc.).
- **Secondary Panel** — Action and agent selection.

Additional controls include password visibility toggles, a **Next** button to validate credentials, and an **Execute** button to run the selected action.

---

## 4. Usage Flow

1. Run the script: `python logistica.py`
2. Fill in the credential fields.
3. Select the desired action from the panel.
4. Click **Next** to load and validate credentials.
5. Click **Execute** to start the task.

> The **Next** button updates the UI state and validates access before the selected action is run.

---

## 5. Available Actions

| Action | Description |
|---|---|
| **Login** | Validates credentials and unlocks the remaining options. |
| **Leave OMS Notes + Drive** | Uploads notes and downloads PDFs to Google Drive. |
| **Extract PDF Name from Downloads** | Extracts and processes PDF filenames from the downloads folder. |
| **Extract PDF Name and Compare in Sheet** | Extracts PDF names and cross-references them against a Google Sheet. |
| **Courier Invoices** | Generates and processes invoices for couriers. |
| **Assign Tracks** | Assigns tracking numbers to orders. |
| **Fix CUIL/CUIT** | Corrects CUIL/CUIT identifiers in the back office using data from Google Sheets. |
| **OMS Inexpress Invoices** | Processes OMS invoices through Inexpress. |
| **Company Refurbish** | Maintenance and follow-up workflow. |
| **Taxes** | Full customs tax pipeline — see Section 6 for full details. |

---

## 6. Taxes Action — Detailed Flow

The **Taxes** action is the most comprehensive workflow in the tool. It covers the full end-to-end process for customs tax extraction, consolidation, data enrichment, AI-powered product classification, and reporting.

### 6.1 Initial Validations

Before anything runs, the script performs a series of checks:

- Verifies the Redshift connection.
- Reads the start and end dates entered in the UI.
- Validates that the date range is logically correct (start before end).
- Requires a courier to be selected.

If any validation fails, the process stops and an error is shown in the UI log.

### 6.2 Data Retrieval from Redshift

Depending on the selected courier, the script queries the Redshift data warehouse to:

- Retrieve orders within the specified date range.
- Resolve identifiers such as `manifest` and `order_id` to enrich the dataset.

The enriched data is then used as input for the classification and consolidation steps.

---

## 7. AI Classification — How It Works

The most technically sophisticated part of the Taxes flow is the AI-powered customs category classification, handled by the `clasificar_productos_con_ia()` function.

### 7.1 What It Does

The goal is to assign each product a valid customs category from an official list. The process works as follows:

1. **Reads the official category list** from a designated Google Sheet.
2. **Normalizes product information** — including product name, SKU, and description — to improve matching quality.
3. **Constructs a prompt** with the product details and sends it to the Gemini AI API.
4. **Receives a structured JSON response** from the model containing:
   - `categoria` — The assigned customs category.
   - `razon` — The AI's reasoning for the classification.
5. **Applies fallback and normalization rules** to ensure the final category exists in the official list, even when the model's output is imprecise.

### 7.2 AI Models and API Modes

The script supports two modes for calling the AI:

- Direct Gemini API
- Vertex AI

The default model is `gemini-2.5-flash`. A lighter fallback model (`gemini-2.5-flash-lite`) is used when the primary model is unavailable or rate-limited.

### 7.3 Caching for Performance

To avoid reclassifying the same products repeatedly, the script maintains a **classification cache stored in Google Sheets**.

- Products already present in the cache are skipped, saving API calls and time.
- Progress is periodically saved to the cache every `TAXES_IA_SAVE_EVERY` items.

### 7.4 Classification Strategy and Robustness

The classification pipeline is built with multiple layers of reliability to handle imperfect AI outputs:

- **Lexical pre-matching** — Before calling the AI, the script attempts a direct text match against the official category list. If a match is found, the API call is skipped entirely.
- **Suggested candidates** — The prompt includes a narrowed list of the most likely categories to guide the model and reduce hallucination.
- **Structured JSON validation** — The model's output is validated to ensure it conforms to the expected format before being accepted.
- **Automatic retries** — On API errors (HTTP 429 rate limits, 500 server errors), the script retries automatically up to `TAXES_IA_MAX_RETRIES` times with appropriate backoff.

This multi-layer approach ensures reliable classification even at scale, with minimal manual intervention.

---

## 8. Outputs and Logging

- The UI displays a live log of all actions, status messages, and errors.
- Some workflows write results directly to Google Sheets.
- Completed files and reports may be saved to Google Drive.
- Critical validation results are logged both in the UI and in the corresponding Sheet when applicable.

---

## 9. Usage Recommendations

- Always complete the Google credentials setup and validate the connection before running **Taxes**.
- Use **Login** first to verify OMS access and all other credentials.
- If category classification is producing unexpected results, review the cache sheet (`TAXES_CACHE_SHEET_ID`) for stale or incorrect entries.
- For Inexpress and ULG couriers, the workflow depends heavily on the naming and content of files stored in Drive — ensure these are correct before running.