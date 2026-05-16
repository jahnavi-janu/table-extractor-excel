import streamlit as st
import os
import re
import time
import io
from pathlib import Path
from typing import Dict, Any, List, Set

import pandas as pd
from dotenv import load_dotenv
from PIL import Image
from app.file_processor import FileProcessor
from app.llm_extractor import LLMExtractor
from app.excel_writer import ExcelWriter
from app.image_editor import ImageEditor

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
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = []
    if "edited_image_bytes" not in st.session_state:
        st.session_state.edited_image_bytes = None
    if "edited_image_filename" not in st.session_state:
        st.session_state.edited_image_filename = "edited_image.png"


def _clear_processed_data() -> None:
    st.session_state.raw_tables = []
    st.session_state.base_filename = "extracted"
    st.session_state.processed_at = None


def _clear_image_edit_data() -> None:
    st.session_state.edited_image_bytes = None
    st.session_state.edited_image_filename = "edited_image.png"


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


def _reset_uploaded_file(uploaded_file: Any) -> None:
    if uploaded_file is not None and hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)


def _get_file_extension(uploaded_file: Any) -> str:
    if uploaded_file is None or not getattr(uploaded_file, "name", None):
        return ""
    return uploaded_file.name.split(".")[-1].lower()


def _is_image_extension(file_ext: str) -> bool:
    return file_ext in {"png", "jpg", "jpeg", "bmp", "tiff"}


def _infer_chat_action(prompt: str, uploaded_file: Any) -> str:
    prompt_lower = (prompt or "").lower()
    file_ext = _get_file_extension(uploaded_file)

    table_keywords = {
        "table", "tables", "excel", "sheet", "spreadsheet", "extract", "rows", "columns"
    }
    image_edit_keywords = {
        "edit", "change", "remove", "replace", "background", "color", "style", "retouch", "enhance"
    }

    if any(keyword in prompt_lower for keyword in table_keywords):
        return "table_extract"
    if any(keyword in prompt_lower for keyword in image_edit_keywords):
        return "image_edit"

    if file_ext in {"pdf", "docx"}:
        return "table_extract"
    if _is_image_extension(file_ext):
        return "image_edit"

    return "unknown"


def _extract_tables_from_file(uploaded_file: Any, api_key: str, status_placeholder: Any = None) -> List[Dict[str, Any]]:
    _reset_uploaded_file(uploaded_file)
    file_processor = FileProcessor()
    llm_extractor = LLMExtractor(api_key)

    if status_placeholder is not None:
        status_placeholder.info("📥 Extracting images from file...")

    images = file_processor.extract_images(uploaded_file)
    if not images:
        return []

    all_tables: List[Dict[str, Any]] = []
    for idx, image in enumerate(images, 1):
        if status_placeholder is not None:
            status_placeholder.info(f"🤖 Processing image {idx}/{len(images)} with OpenAI Vision...")
        tables = llm_extractor.extract_tables_from_image(image)
        all_tables.extend(tables)

    return all_tables

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

    tab_chat, tab_tables, tab_image = st.tabs(["🤖 Assistant Chat", "📊 Table Extractor", "🖼️ Image Edit"])

    with tab_chat:
        st.subheader("Assistant")
        st.caption("Ask me to extract tables to Excel or edit an image. Upload a file, then type your request.")

        chat_uploaded_file = st.file_uploader(
            "Upload file for chat action",
            type=["pdf", "docx", "png", "jpg", "jpeg", "bmp", "tiff"],
            key="chat_uploaded_file",
        )

        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        user_prompt = st.chat_input("Example: Extract tables to Excel OR make background white")
        if user_prompt:
            st.session_state.chat_messages.append({"role": "user", "content": user_prompt})

            with st.chat_message("assistant"):
                action = _infer_chat_action(user_prompt, chat_uploaded_file)
                try:
                    if action == "table_extract":
                        if chat_uploaded_file is None:
                            assistant_text = "Please upload a PDF/DOCX/image first, then ask to extract tables."
                        else:
                            with st.spinner("Extracting tables..."):
                                tables = _extract_tables_from_file(chat_uploaded_file, api_key)

                            if not tables:
                                assistant_text = "I could not find tables in this file."
                            else:
                                st.session_state.raw_tables = tables
                                st.session_state.base_filename = Path(chat_uploaded_file.name).stem
                                st.session_state.processed_at = time.time()
                                assistant_text = (
                                    f"Done ✅ I extracted {len(tables)} table(s). "
                                    "Open the 'Table Extractor' tab to review and download Excel."
                                )

                    elif action == "image_edit":
                        file_ext = _get_file_extension(chat_uploaded_file)
                        if chat_uploaded_file is None:
                            assistant_text = "Please upload an image file first (PNG/JPG/BMP/TIFF)."
                        elif not _is_image_extension(file_ext):
                            assistant_text = "For image editing, upload an image (PNG/JPG/BMP/TIFF)."
                        else:
                            _reset_uploaded_file(chat_uploaded_file)
                            image = Image.open(io.BytesIO(chat_uploaded_file.read())).convert("RGBA")
                            editor = ImageEditor(api_key)
                            with st.spinner("Editing image..."):
                                edited_bytes = editor.edit_image(image=image, prompt=user_prompt)

                            st.session_state.edited_image_bytes = edited_bytes
                            st.session_state.edited_image_filename = f"{Path(chat_uploaded_file.name).stem}_edited.png"
                            st.session_state.processed_at = time.time()
                            assistant_text = "Done ✅ Image edited successfully. Preview and download are shown below."
                    else:
                        assistant_text = (
                            "I can currently do two actions: (1) extract tables to Excel and "
                            "(2) edit images from your prompt. Please upload a file and ask one of these."
                        )

                except Exception as error:
                    assistant_text = f"Sorry, I couldn't complete that action: {str(error)}"

                st.markdown(assistant_text)
                st.session_state.chat_messages.append({"role": "assistant", "content": assistant_text})

        if st.session_state.edited_image_bytes:
            st.markdown("### Latest Edited Image")
            st.image(st.session_state.edited_image_bytes, use_container_width=True)
            image_downloaded_from_chat = st.download_button(
                "⬇️ Download Edited Image",
                data=st.session_state.edited_image_bytes,
                file_name=st.session_state.edited_image_filename,
                mime="image/png",
                key="chat_download_edited_image",
            )
            if image_downloaded_from_chat and auto_delete_after_download:
                _clear_image_edit_data()
                st.success("Image download completed. Edited image auto-deleted from memory.")
                st.rerun()

        if st.button("Clear chat history", key="clear_chat_history"):
            st.session_state.chat_messages = []
            st.success("Chat history cleared.")
            st.rerun()

    with tab_tables:
        st.subheader("1️⃣ Upload File")
        table_uploaded_file = st.file_uploader(
            "Choose a file (PDF, Word, or Image)",
            type=["pdf", "docx", "png", "jpg", "jpeg", "bmp", "tiff"],
            help="Supported formats: PDF, Word (.docx), PNG, JPG, BMP, TIFF",
            key="table_uploaded_file",
        )
        process_clicked = st.button("Process File", type="primary", use_container_width=True, key="table_process")
        table_status_placeholder = st.empty()

        if table_uploaded_file is not None and process_clicked:
            try:
                with st.spinner("Processing file..."):
                    all_tables = _extract_tables_from_file(table_uploaded_file, api_key, table_status_placeholder)

                if not all_tables:
                    st.warning("⚠️ No tables detected in the file")
                else:
                    st.success(f"✅ Extracted {len(all_tables)} table(s) successfully")
                    st.session_state.raw_tables = all_tables
                    st.session_state.base_filename = Path(table_uploaded_file.name).stem
                    st.session_state.processed_at = time.time()
                    table_status_placeholder.success("✅ Conversion complete!")
            except ValueError as e:
                st.error(f"❌ Configuration Error: {str(e)}")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.write("Debug info:", str(type(e).__name__))

        if st.session_state.raw_tables:
            st.subheader("2️⃣ Privacy Review, Preview & Download")

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
                key="table_mask_columns",
            )

            export_tables = _build_export_tables(
                raw_tables=st.session_state.raw_tables,
                enable_masking=enable_pii_masking,
                selected_columns=selected_columns,
                auto_detect_columns=auto_detect_pii_columns,
            )

            excel_writer = ExcelWriter()
            table_status_placeholder.info("📝 Creating Excel file...")
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
            downloaded = st.download_button(
                label="⬇️ Download Excel File",
                data=excel_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="table_download_excel",
            )

            clear_now = st.button("🗑️ Clear processed data from memory", use_container_width=True, key="table_clear_data")
            if clear_now:
                _clear_processed_data()
                st.success("Processed data cleared from memory.")
                st.rerun()

            if downloaded and auto_delete_after_download:
                _clear_processed_data()
                st.success("Download completed. Processed data auto-deleted from memory.")
                st.rerun()

    with tab_image:
        st.subheader("Prompt-based Image Editing")
        st.caption("Upload an image and describe exactly what should change.")

        image_uploaded_file = st.file_uploader(
            "Upload image",
            type=["png", "jpg", "jpeg", "bmp", "tiff"],
            key="image_edit_uploaded_file",
        )
        image_prompt = st.text_area(
            "Describe the edit",
            placeholder="Example: Remove background and make it plain white. Keep subject unchanged.",
            key="image_edit_prompt",
        )
        apply_edit = st.button("✨ Apply Edit", use_container_width=True, key="image_edit_apply")

        if image_uploaded_file is not None:
            _reset_uploaded_file(image_uploaded_file)
            preview_bytes = image_uploaded_file.read()
            st.markdown("**Original image**")
            st.image(preview_bytes, use_container_width=True)

        if apply_edit:
            try:
                if image_uploaded_file is None:
                    raise ValueError("Please upload an image first.")
                if not image_prompt.strip():
                    raise ValueError("Please enter an edit prompt.")

                _reset_uploaded_file(image_uploaded_file)
                source_image = Image.open(io.BytesIO(image_uploaded_file.read())).convert("RGBA")

                with st.spinner("Applying image edit..."):
                    editor = ImageEditor(api_key)
                    edited_bytes = editor.edit_image(source_image, image_prompt)

                st.session_state.edited_image_bytes = edited_bytes
                st.session_state.edited_image_filename = f"{Path(image_uploaded_file.name).stem}_edited.png"
                st.session_state.processed_at = time.time()
                st.success("✅ Image edited successfully")
            except Exception as error:
                st.error(f"❌ Could not edit image: {str(error)}")

        if st.session_state.edited_image_bytes:
            st.markdown("**Edited image**")
            st.image(st.session_state.edited_image_bytes, use_container_width=True)
            image_downloaded_from_tab = st.download_button(
                "⬇️ Download Edited Image",
                data=st.session_state.edited_image_bytes,
                file_name=st.session_state.edited_image_filename,
                mime="image/png",
                key="tab_download_edited_image",
            )
            if image_downloaded_from_tab and auto_delete_after_download:
                _clear_image_edit_data()
                st.success("Image download completed. Edited image auto-deleted from memory.")
                st.rerun()
            if st.button("🗑️ Clear edited image", use_container_width=True, key="clear_edited_image"):
                _clear_image_edit_data()
                st.success("Edited image cleared from memory.")
                st.rerun()

if __name__ == "__main__":
    main()
