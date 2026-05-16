import streamlit as st
import os
import re
import time
from pathlib import Path
from typing import Dict, Any, List, Set

import pandas as pd
from dotenv import load_dotenv
from app.file_processor import FileProcessor
from app.llm_extractor import LLMExtractor
from app.excel_writer import ExcelWriter

try:
    from PyQt5.QtWidgets import QApplication, QFileDialog
except Exception:
    QApplication = None
    QFileDialog = None

# Load environment variables
load_dotenv()


def _initialize_state() -> None:
    if "raw_tables" not in st.session_state:
        st.session_state.raw_tables = []
    if "base_filename" not in st.session_state:
        st.session_state.base_filename = "extracted"
    if "processed_at" not in st.session_state:
        st.session_state.processed_at = None


def _clear_processed_data() -> None:
    st.session_state.raw_tables = []
    st.session_state.base_filename = "extracted"
    st.session_state.processed_at = None


def _mask_sensitive_text(value: Any) -> Any:
    if value is None:
        return value

    text = str(value)
    if not text.strip():
        return value

    text = re.sub(
        r"([A-Za-z0-9._%+-])[A-Za-z0-9._%+-]*@([A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        r"\1***@\2",
        text,
    )
    text = re.sub(r"\b(\+?\d[\d\-\s]{5,}\d)\b", lambda match: _mask_digits(match.group(1)), text)
    text = re.sub(r"\b\d{8,}\b", lambda match: _mask_digits(match.group(0)), text)
    return text


def _mask_digits(text: str) -> str:
    digits = re.sub(r"\D", "", text)
    if len(digits) <= 4:
        return "*" * len(digits)
    return "*" * (len(digits) - 4) + digits[-4:]


def _auto_detect_pii_columns(columns: List[str]) -> Set[str]:
    keywords = {
        "name",
        "full name",
        "employee",
        "client",
        "customer",
        "phone",
        "mobile",
        "contact",
        "email",
        "mail",
        "address",
        "aadhaar",
        "pan",
        "ssn",
        "id",
        "dob",
    }
    pii_columns = set()
    for column in columns:
        col_lower = column.lower().strip()
        if any(keyword in col_lower for keyword in keywords):
            pii_columns.add(column)
    return pii_columns


def _build_export_tables(
    raw_tables: List[Dict[str, Any]],
    enable_masking: bool,
    selected_columns: List[str],
    auto_detect_columns: bool,
) -> List[Dict[str, Any]]:
    export_tables = []

    for table in raw_tables:
        table_copy = dict(table)
        data = table.get("data")
        if not isinstance(data, pd.DataFrame):
            export_tables.append(table_copy)
            continue

        masked_df = data.copy()
        if enable_masking:
            columns_to_mask = set(selected_columns)
            if auto_detect_columns:
                columns_to_mask |= _auto_detect_pii_columns([str(column) for column in masked_df.columns])

            for column in masked_df.columns:
                if str(column) in columns_to_mask:
                    masked_df[column] = masked_df[column].apply(_mask_sensitive_text)

        table_copy["data"] = masked_df
        export_tables.append(table_copy)

    return export_tables


def _save_excel_to_custom_path(excel_buffer: bytes, output_path: str) -> str:
    target_path = Path(output_path).expanduser()
    if target_path.is_dir():
        raise ValueError("Please provide a full file path including filename, not just a folder.")

    if target_path.suffix.lower() != ".xlsx":
        target_path = target_path.with_suffix(".xlsx")

    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_bytes(excel_buffer)
    return str(target_path)


def _browse_save_path(default_filename: str) -> str:
    """Open a native Save As dialog and return selected .xlsx path."""
    if QApplication is None or QFileDialog is None:
        raise RuntimeError("PyQt5 is not available. Install it with: pip install PyQt5")

    created_app = False
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
        created_app = True

    try:
        selected_path, _ = QFileDialog.getSaveFileName(
            None,
            "Save Excel file as",
            default_filename,
            "Excel Files (*.xlsx);;All Files (*)",
        )
    finally:
        if created_app:
            app.quit()

    return selected_path

def main():
    _initialize_state()

    st.set_page_config(
        page_title="Table Extractor to Excel",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("📊 PDF/Word to Excel Converter")
    st.markdown("Extract tables from PDF, Word, or images and convert them to Excel format")
    
    # Sidebar
    with st.sidebar:
        st.header("Settings")
        
        try:
            api_key = st.secrets["OPENAI_API_KEY"]
        except (KeyError, FileNotFoundError):
            api_key = os.getenv("OPENAI_API_KEY", "")
        
        if not api_key:
            st.error("⚠️ OpenAI API key not configured. Please set OPENAI_API_KEY in secrets or .env file.")
            st.stop()

        st.success("✓ API key loaded securely from config")

        st.divider()
        st.subheader("Privacy")
        enable_pii_masking = st.checkbox("Enable PII masking", value=True)
        auto_detect_pii_columns = st.checkbox("Auto-mask common PII columns", value=True)
        auto_delete_after_download = st.checkbox("Auto-delete processed data after download", value=True)
        
        st.divider()
        st.subheader("Export")
        export_as_multiple_sheets = st.checkbox("Multiple sheets (one table per sheet)", value=False)

        st.divider()
        with st.expander("📋 Privacy & Data Policy", expanded=False):
            st.markdown("""
**Data Handling:**
- Only table images are sent to OpenAI Vision API for extraction
- Original PDF/Word files are NOT uploaded
- Extracted tables are processed in-memory and auto-deleted based on your settings

**OpenAI Data Usage:**
- Your images may be used to improve OpenAI models (per OpenAI terms)
- To disable: Contact OpenAI support for API account data retention settings
- Images in transit: HTTPS encrypted

**Recommendation:**
- Enable PII masking for tables containing sensitive data
- Use auto-delete to clear processed data after download
- Do not upload files with highly sensitive information if concerned
            """)

    if st.session_state.processed_at:
        age_seconds = int(time.time() - st.session_state.processed_at)
        st.caption(f"Session data age: {age_seconds} seconds")
    
    # Main content
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        st.subheader("1️⃣ Upload File")
        uploaded_file = st.file_uploader(
            "Choose a file (PDF, Word, or Image)",
            type=["pdf", "docx", "png", "jpg", "jpeg", "bmp", "tiff"],
            help="Supported formats: PDF, Word (.docx), PNG, JPG, BMP, TIFF"
        )
        process_clicked = st.button("Process File", type="primary", use_container_width=True)
    
    with col2:
        st.subheader("2️⃣ Processing Status")
        status_placeholder = st.empty()
    
    if uploaded_file is not None and process_clicked:
        try:
            # Validate API key
            if not api_key or api_key == "":
                st.error("❌ API key not available")
                return
            
            with st.spinner("Processing file..."):
                # Initialize processors
                file_processor = FileProcessor()
                llm_extractor = LLMExtractor(api_key)
                excel_writer = ExcelWriter()
                
                # Step 1: Extract images from file
                status_placeholder.info("📥 Extracting images from file...")
                images = file_processor.extract_images(uploaded_file)
                
                if not images:
                    st.error("❌ No images or tables found in the file")
                    return
                
                st.success(f"✅ Found {len(images)} image(s) in the file")
                
                # Step 2: Extract table data using LLM
                status_placeholder.info("🤖 Analyzing images with OpenAI Vision...")
                all_tables = []
                
                for idx, image in enumerate(images, 1):
                    st.info(f"Processing image {idx}/{len(images)}...")
                    tables = llm_extractor.extract_tables_from_image(image)
                    all_tables.extend(tables)
                    st.write(f"Found {len(tables)} table(s) in image {idx}")
                
                if not all_tables:
                    st.warning("⚠️ No tables detected in the images")
                    return

                st.success(f"✅ Extracted {len(all_tables)} table(s) successfully")

                st.session_state.raw_tables = all_tables
                st.session_state.base_filename = Path(uploaded_file.name).stem
                st.session_state.processed_at = time.time()

                status_placeholder.success("✅ Conversion complete!")

        except ValueError as e:
            st.error(f"❌ Configuration Error: {str(e)}")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.write("Debug info:", str(type(e).__name__))

    if st.session_state.raw_tables:
        st.subheader("3️⃣ Privacy Review, Preview & Download")

        all_columns = sorted(
            {
                str(column)
                for table in st.session_state.raw_tables
                if isinstance(table, dict) and isinstance(table.get("data"), pd.DataFrame)
                for column in table["data"].columns
            }
        )

        selected_columns = st.multiselect(
            "Columns to always mask",
            options=all_columns,
            default=[],
            help="Masking applies to preview and Excel export.",
        )

        export_tables = _build_export_tables(
            raw_tables=st.session_state.raw_tables,
            enable_masking=enable_pii_masking,
            selected_columns=selected_columns,
            auto_detect_columns=auto_detect_pii_columns,
        )

        excel_writer = ExcelWriter()
        status_placeholder.info("📝 Creating Excel file...")
        excel_buffer = excel_writer.create_excel(
            export_tables,
            single_sheet=not export_as_multiple_sheets,
        )

        with st.expander("View extracted tables (post-masking):", expanded=True):
            for idx, table in enumerate(export_tables, 1):
                st.write(f"**Table {idx}:**")
                if isinstance(table, dict) and "data" in table:
                    st.dataframe(table["data"], use_container_width=True)
                else:
                    st.write(table)

        filename = st.session_state.base_filename + "_extracted.xlsx"
        st.caption("Browser downloads are controlled by browser settings. Use custom path save below for exact path.")
        downloaded = st.download_button(
            label="⬇️ Download Excel File",
            data=excel_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        #st.info("Custom path save is currently disabled. Please use the Download button above.")

        # Custom-path save flow disabled as requested.
        # if "custom_output_path" not in st.session_state:
        #     st.session_state.custom_output_path = ""
        #
        # browse_clicked = st.button("📂 Browse save location", use_container_width=True)
        # if browse_clicked:
        #     try:
        #         selected_path = _browse_save_path(filename)
        #         if selected_path:
        #             st.session_state.custom_output_path = selected_path
        #     except Exception as error:
        #         st.warning(f"Browse dialog unavailable: {str(error)}")
        #
        # custom_output_path = st.text_input(
        #     "Custom output path (save directly on this machine)",
        #     key="custom_output_path",
        #     placeholder="Click 'Browse save location' and choose where to save"
        # )
        # save_custom = st.button("💾 Save file to custom path", use_container_width=True)
        # if save_custom:
        #     try:
        #         if not str(custom_output_path).strip():
        #             raise ValueError("Please choose a save path first using Browse, or enter a full file path.")
        #         saved_path = _save_excel_to_custom_path(excel_buffer, custom_output_path)
        #         st.success(f"File saved to: {saved_path}")
        #         if auto_delete_after_download:
        #             _clear_processed_data()
        #             st.info("Processed data auto-deleted from memory after save.")
        #             st.rerun()
        #     except Exception as error:
        #         st.error(f"Could not save file: {str(error)}")

        clear_now = st.button("🗑️ Clear processed data from memory", use_container_width=True)
        if clear_now:
            _clear_processed_data()
            st.success("Processed data cleared from memory.")
            st.rerun()

        if downloaded and auto_delete_after_download:
            _clear_processed_data()
            st.success("Download completed. Processed data auto-deleted from memory.")
            st.rerun()

if __name__ == "__main__":
    main()
