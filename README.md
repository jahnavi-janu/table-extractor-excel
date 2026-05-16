# PDF/Word Table Images to Excel

Streamlit app that uploads `PDF`, `DOCX`, or image files, uses OpenAI Vision to extract table data, then generates and downloads an Excel file.

## Features

- Upload `PDF`, `DOCX`, `PNG`, `JPG`, `BMP`, `TIFF`
- Converts PDF pages/images into extraction inputs
- Sends images to OpenAI Vision for table extraction
- Writes each extracted table to Excel sheets
- Downloads generated `.xlsx`

## Setup

1. Create and activate virtual environment.
2. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
3. Set API key in `.env` (copy from `.env.example`):
   ```env
   OPENAI_API_KEY=your_openai_api_key_here
   ```

## Run

```powershell
streamlit run app.py
```

## Notes

- Legacy `.doc` files are not supported. Convert to `.docx` first.
- Extraction quality depends on image clarity and table structure.
