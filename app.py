import streamlit as st
import os
import re
import time
import io
from pathlib import Path
from typing import Dict, Any, List, Set, Tuple

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
    if "edited_file_bytes" not in st.session_state:
        st.session_state.edited_file_bytes = None
    if "edited_file_filename" not in st.session_state:
        st.session_state.edited_file_filename = "edited_file.png"
    if "edited_file_mime" not in st.session_state:
        st.session_state.edited_file_mime = "image/png"
    if "edited_preview_images" not in st.session_state:
        st.session_state.edited_preview_images = []


def _clear_processed_data() -> None:
    st.session_state.raw_tables = []
    st.session_state.base_filename = "extracted"
    st.session_state.processed_at = None


def _clear_edit_data() -> None:
    st.session_state.edited_file_bytes = None
    st.session_state.edited_file_filename = "edited_file.png"
    st.session_state.edited_file_mime = "image/png"
    st.session_state.edited_preview_images = []


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


def _is_pdf_extension(file_ext: str) -> bool:
    return file_ext == "pdf"


def _render_preview_png(image: Image.Image) -> bytes:
    preview_buffer = io.BytesIO()
    image.convert("RGB").save(preview_buffer, format="PNG")
    return preview_buffer.getvalue()


def _edit_file_from_prompt(uploaded_file: Any, prompt: str, api_key: str) -> Tuple[bytes, str, str, List[bytes]]:
    _reset_uploaded_file(uploaded_file)
    file_ext = _get_file_extension(uploaded_file)
    editor = ImageEditor(api_key)

    if _is_image_extension(file_ext):
        image = Image.open(io.BytesIO(uploaded_file.read())).convert("RGBA")
        edited_bytes = editor.edit_image(image=image, prompt=prompt)
        return (
            edited_bytes,
            f"{Path(uploaded_file.name).stem}_edited.png",
            "image/png",
            [edited_bytes],
        )

    if _is_pdf_extension(file_ext):
        file_processor = FileProcessor()
        page_images = file_processor.extract_images(uploaded_file)
        if not page_images:
            raise ValueError("No pages found in PDF.")

        edited_pages: List[Image.Image] = []
        preview_images: List[bytes] = []
        for page_image in page_images:
            edited_page_bytes = editor.edit_image(image=page_image, prompt=prompt)
            edited_page = Image.open(io.BytesIO(edited_page_bytes)).convert("RGB")
            edited_pages.append(edited_page)
            preview_images.append(_render_preview_png(edited_page))

        pdf_buffer = io.BytesIO()
        first_page, other_pages = edited_pages[0], edited_pages[1:]
        first_page.save(pdf_buffer, format="PDF", save_all=True, append_images=other_pages)
        return (
            pdf_buffer.getvalue(),
            f"{Path(uploaded_file.name).stem}_edited.pdf",
            "application/pdf",
            preview_images,
        )

    raise ValueError("For editing, upload an image (PNG/JPG/BMP/TIFF) or PDF.")


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

    tab_chat, tab_tables, tab_edit = st.tabs(["🤖 Assistant Chat", "📊 Table Extractor", "✏️ Edit"])

    with tab_chat:
        st.subheader("Assistant")
        st.caption("Ask me to extract tables to Excel or edit an image/PDF. Upload a file, then type your request.")

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
                        if chat_uploaded_file is None:
                            assistant_text = "Please upload an image or PDF first (PNG/JPG/BMP/TIFF/PDF)."
                        else:
                            with st.spinner("Applying edit..."):
                                edited_bytes, edited_filename, edited_mime, preview_images = _edit_file_from_prompt(
                                    uploaded_file=chat_uploaded_file,
                                    prompt=user_prompt,
                                    api_key=api_key,
                                )

                            st.session_state.edited_file_bytes = edited_bytes
                            st.session_state.edited_file_filename = edited_filename
                            st.session_state.edited_file_mime = edited_mime
                            st.session_state.edited_preview_images = preview_images
                            st.session_state.processed_at = time.time()
                            assistant_text = "Done ✅ Edit completed successfully. Preview and download are shown below."
                    else:
                        assistant_text = (
                            "I can currently do two actions: (1) extract tables to Excel and "
                            "(2) edit images/PDFs from your prompt. Please upload a file and ask one of these."
                        )

                except Exception as error:
                    assistant_text = f"Sorry, I couldn't complete that action: {str(error)}"

                st.markdown(assistant_text)
                st.session_state.chat_messages.append({"role": "assistant", "content": assistant_text})

        if st.session_state.edited_file_bytes:
            st.markdown("### Latest Edited Output")
            if st.session_state.edited_preview_images:
                st.image(st.session_state.edited_preview_images[0], use_container_width=True)
                if len(st.session_state.edited_preview_images) > 1:
                    st.caption(f"Previewing page 1 of {len(st.session_state.edited_preview_images)} edited pages.")

            file_kind = "File"
            if st.session_state.edited_file_mime == "image/png":
                file_kind = "Image"
            elif st.session_state.edited_file_mime == "application/pdf":
                file_kind = "PDF"

            image_downloaded_from_chat = st.download_button(
                f"⬇️ Download Edited {file_kind}",
                data=st.session_state.edited_file_bytes,
                file_name=st.session_state.edited_file_filename,
                mime=st.session_state.edited_file_mime,
                key="chat_download_edited_file",
            )
            if image_downloaded_from_chat and auto_delete_after_download:
                _clear_edit_data()
                st.success("Download completed. Edited output auto-deleted from memory.")
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

    with tab_edit:
        st.subheader("Prompt-based Editing")
        st.caption("Upload an image or PDF and describe exactly what should change.")

        edit_uploaded_file = st.file_uploader(
            "Upload image or PDF",
            type=["pdf", "png", "jpg", "jpeg", "bmp", "tiff"],
            key="edit_uploaded_file",
        )
        edit_prompt = st.text_area(
            "Describe the edit",
            placeholder="Example: Remove background and make it plain white. Keep subject unchanged.",
            key="edit_prompt",
        )
        apply_edit = st.button("✨ Apply Edit", use_container_width=True, key="edit_apply")

        if edit_uploaded_file is not None:
            file_ext = _get_file_extension(edit_uploaded_file)
            if _is_image_extension(file_ext):
                _reset_uploaded_file(edit_uploaded_file)
                preview_bytes = edit_uploaded_file.read()
                st.markdown("**Original image**")
                st.image(preview_bytes, use_container_width=True)
            elif _is_pdf_extension(file_ext):
                _reset_uploaded_file(edit_uploaded_file)
                pdf_pages = FileProcessor().extract_images(edit_uploaded_file)
                st.markdown("**Original PDF (first page preview)**")
                if pdf_pages:
                    st.image(_render_preview_png(pdf_pages[0]), use_container_width=True)
                    if len(pdf_pages) > 1:
                        st.caption(f"Previewing page 1 of {len(pdf_pages)} pages.")

        if apply_edit:
            try:
                if edit_uploaded_file is None:
                    raise ValueError("Please upload an image or PDF first.")
                if not edit_prompt.strip():
                    raise ValueError("Please enter an edit prompt.")

                with st.spinner("Applying edit..."):
                    edited_bytes, edited_filename, edited_mime, preview_images = _edit_file_from_prompt(
                        uploaded_file=edit_uploaded_file,
                        prompt=edit_prompt,
                        api_key=api_key,
                    )

                st.session_state.edited_file_bytes = edited_bytes
                st.session_state.edited_file_filename = edited_filename
                st.session_state.edited_file_mime = edited_mime
                st.session_state.edited_preview_images = preview_images
                st.session_state.processed_at = time.time()
                st.success("✅ Edit completed successfully")
            except Exception as error:
                st.error(f"❌ Could not apply edit: {str(error)}")

        if st.session_state.edited_file_bytes:
            st.markdown("**Edited output**")
            if st.session_state.edited_preview_images:
                st.image(st.session_state.edited_preview_images[0], use_container_width=True)
                if len(st.session_state.edited_preview_images) > 1:
                    st.caption(f"Previewing page 1 of {len(st.session_state.edited_preview_images)} edited pages.")

            file_kind = "File"
            if st.session_state.edited_file_mime == "image/png":
                file_kind = "Image"
            elif st.session_state.edited_file_mime == "application/pdf":
                file_kind = "PDF"

            image_downloaded_from_tab = st.download_button(
                f"⬇️ Download Edited {file_kind}",
                data=st.session_state.edited_file_bytes,
                file_name=st.session_state.edited_file_filename,
                mime=st.session_state.edited_file_mime,
                key="tab_download_edited_file",
            )
            if image_downloaded_from_tab and auto_delete_after_download:
                _clear_edit_data()
                st.success("Download completed. Edited output auto-deleted from memory.")
                st.rerun()
            if st.button("🗑️ Clear edited image", use_container_width=True, key="clear_edited_image"):
                _clear_edit_data()
                st.success("Edited output cleared from memory.")
                st.rerun()

if __name__ == "__main__":
    main()
