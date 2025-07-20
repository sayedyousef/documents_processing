# document_processor.py
"""Core document processing functionality - refactored."""

import logging
from pathlib import Path
from typing import List
from docx import Document as DocxDocument

from models import Document
from document_processor.text_extractor import TextExtractor
from document_processor.image_analyzer import ImageAnalyzer
from document_processor.section_extractor import SectionExtractor

class DocumentProcessor:
    """Handles processing of Word documents."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.doc_counter = 0
        
        # Initialize extractors
        self.text_extractor = TextExtractor()
        self.image_analyzer = ImageAnalyzer()
        self.section_extractor = SectionExtractor()
    
    def process_folder(self, folder_path: Path) -> List[Document]:
        """Process all Word documents in folder and subfolders."""
        documents = []
        
        # Find all .docx files recursively
        for doc_path in folder_path.rglob("*.docx"):
            # Skip temporary files
            if doc_path.name.startswith("~"):
                continue
            
            try:
                self.logger.info(f"Processing: {doc_path}")
                doc = self.process_document(doc_path)
                documents.append(doc)
            except Exception as e:
                self.logger.error(f"Error processing {doc_path}: {e}")
        
        return documents
    
    def process_document(self, file_path: Path) -> Document:
        """Process a single Word document."""
        self.doc_counter += 1
        docx = DocxDocument(file_path)
        
        # Extract sections first for analysis
        sections = self.section_extractor.extract_sections(docx)
        
        # Check style compliance
        uses_proper_styles = self.section_extractor.check_style_compliance(sections)
        
        # Log style summary
        self.section_extractor.log_style_summary(file_path, sections)
        
        # Create document object with all extracted data
        doc = Document(
            id=self.doc_counter,
            file_path=file_path,
            name=file_path.name,
            parent_folder=file_path.parent.name,
            title=self.text_extractor.extract_title(docx),
            author=self.text_extractor.extract_author(docx),
            word_count=self.text_extractor.count_words(docx),
            image_count=self.image_analyzer.count_images_total(docx),
            unique_image_count=self.image_analyzer.count_unique_images(docx, file_path),
            sections=sections,
            uses_proper_styles=uses_proper_styles
        )
        
        return doc