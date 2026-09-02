# Invoice Extractor POC

This project is a Proof of Concept (POC) designed to automatically extract data from "Vendor Payment Report" invoices (PDFs or images) using the GPT-4o Vision API. It provides a simple, modern Web UI to upload multiple files, processes them, and outputs a neatly formatted Excel report (`Astrya_Invoices.xlsx`).

## Features
- **Web UI:** A clean, drag-and-drop-style web interface built with Flask (light blue and white theme).
- **Automated Extraction:** Uses OpenAI's GPT-4o Vision to read facility codes, invoice numbers, and grand total amounts directly off the page images.
- **Auto-Formatting:** Generates a structured Excel file with an auto-calculated Grand Total row.
- **Local Persistence:** Stores the uploaded files in an `input/` folder and the resulting report in an `output/` folder for your records.

## Prerequisites
- Python 3.8+
- An OpenAI API Key with access to the `gpt-4o` model.

## Installation & Setup

1. **Create and activate a virtual environment** (recommended):
   ```bash
   python -m venv venv
   .\venv\Scripts\activate  # On Windows
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure Environment Variables**:
   Create a `.env` file in the root of the project (same directory as `app.py`) and add your OpenAI API key:
   ```env
   OPENAI_API_KEY="sk-your-openai-api-key-here"
   ```

## Running the Web UI

1. Start the Flask application by running:
   ```bash
   python app.py
   ```
2. Open your web browser and navigate to: **http://127.0.0.1:5000**
3. Select or drag-and-drop your invoice PDFs (e.g., `CMC_Astrya Invoice 1260077613.pdf`, `Tele_Psych_Diamond_Bar.pdf`) and click **Process Invoices**.
4. The system will process the documents and automatically download the `Extracted_Invoices.xlsx` file to your computer.

### Note on Directories
When you upload files via the Web UI:
- Files are temporarily saved to the local `input/` directory for processing (this directory is cleared before each new batch).
- The generated Excel file is saved to the local `output/` directory before being sent to your browser for download.

## Running via Command Line (Optional)
If you prefer to bypass the web UI, you can run the core extraction script directly via the CLI:

```bash
python "extract_invoices_to_excel (1).py" --pdf_dir ./input --output ./output/Astrya_Invoices.xlsx
```
