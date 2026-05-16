import io
from typing import List
from PIL import Image
import pymupdf  # fitz
from docx import Document


class FileProcessor:
    """Handles extraction of images from various file formats"""
    
    SUPPORTED_FORMATS = {
        'pdf': 'application/pdf',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'bmp': 'image/bmp',
        'tiff': 'image/tiff'
    }
    
    def extract_images(self, uploaded_file) -> List[Image.Image]:
        """
        Extract images from uploaded file based on file type
        
        Args:
            uploaded_file: Streamlit uploaded file object
            
        Returns:
            List of PIL Image objects
        """
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'pdf':
            return self._extract_from_pdf(uploaded_file)
        elif file_extension == 'docx':
            return self._extract_from_word(uploaded_file)
        elif file_extension == 'doc':
            raise ValueError("Legacy .doc files are not supported. Please save as .docx and upload again.")
        elif file_extension in ['png', 'jpg', 'jpeg', 'bmp', 'tiff']:
            return self._extract_from_image(uploaded_file)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")
    
    def _extract_from_pdf(self, uploaded_file) -> List[Image.Image]:
        """Render each PDF page as an image for consistent table extraction."""
        images = []
        
        try:
            pdf_bytes = uploaded_file.read()
            pdf_document = pymupdf.open(stream=pdf_bytes, filetype="pdf")

            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2))
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data)).convert("RGB")
                images.append(img)
            
            pdf_document.close()
            
        except Exception as e:
            raise ValueError(f"Error processing PDF: {str(e)}")
        
        return images
    
    def _extract_from_word(self, uploaded_file) -> List[Image.Image]:
        """Extract images from Word document"""
        images = []
        
        try:
            doc_bytes = uploaded_file.read()
            doc = Document(io.BytesIO(doc_bytes))
            
            # Extract images from document
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    image_data = rel.target_part.blob
                    img = Image.open(io.BytesIO(image_data)).convert("RGB")
                    images.append(img)
            
            # Also extract images from tables within document
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for rel in cell._element.part.rels.values():
                            if "image" in rel.target_ref:
                                image_data = rel.target_part.blob
                                img = Image.open(io.BytesIO(image_data)).convert("RGB")
                                images.append(img)
        
        except Exception as e:
            raise ValueError(f"Error processing Word document: {str(e)}")
        
        return images
    
    def _extract_from_image(self, uploaded_file) -> List[Image.Image]:
        """Extract/load image file"""
        try:
            img = Image.open(io.BytesIO(uploaded_file.read())).convert("RGB")
            return [img]
        except Exception as e:
            raise ValueError(f"Error processing image: {str(e)}")
