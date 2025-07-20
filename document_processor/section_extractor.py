# section_extractor.py
"""Module for extracting document sections and analyzing styles."""

import logging
import re
from typing import List
from docx import Document as DocxDocument
from docx.shared import Pt
from models import DocumentSection

class SectionExtractor:
    """Handles section extraction and style analysis."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_sections(self, docx: DocxDocument) -> List[DocumentSection]:
        """Extract sections with their styles and content, including image captions."""
        sections = []
        current_section = None
        
        for paragraph in docx.paragraphs:
            text = paragraph.text.strip()
            
            # Skip empty paragraphs
            if not text:
                continue
            
            # Check if this is an image or table caption
            is_image_caption = self._is_image_or_table_caption(text)
            
            # Check if this is a heading
            is_heading = (paragraph.style.name.startswith('Heading') or 
                         self._has_special_formatting(paragraph) or 
                         is_image_caption)
            
            if is_heading:
                # Save previous section if exists
                if current_section:
                    sections.append(current_section)
                
                # Determine section type
                section_type = "image" if is_image_caption else "text"
                
                # Create new section
                current_section = DocumentSection(
                    heading=text,
                    style_name=paragraph.style.name,
                    section_type=section_type
                )
            elif current_section:
                # Add text to current section
                current_section.text += paragraph.text + "\n"
        
        # Don't forget the last section
        if current_section:
            sections.append(current_section)
        
        return sections
    
    def check_style_compliance(self, sections: List[DocumentSection]) -> bool:
        """Check if document uses proper heading styles."""
        proper_styles = {'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 
                        'Heading 5', 'Heading 6', 'Title', 'Subtitle'}
        
        # Check text sections (excluding image captions)
        text_sections = [s for s in sections if s.section_type == "text"]
        
        if not text_sections:
            return True
        
        # Count sections using proper styles
        proper_count = sum(1 for s in text_sections if s.style_name in proper_styles)
        
        # Consider compliant if at least 80% of sections use proper styles
        compliance_ratio = proper_count / len(text_sections)
        
        return compliance_ratio >= 0.8
    
    def log_style_summary(self, file_path, sections: List[DocumentSection]):
        """Log summary of styles used in the document."""
        style_count = {}
        image_count = 0
        
        for section in sections:
            if section.section_type == "image":
                image_count += 1
            else:
                style = section.style_name or "Normal"
                style_count[style] = style_count.get(style, 0) + 1
        
        self.logger.info(f"\n========== Style Summary for {file_path.name} ==========")
        self.logger.info(f"Text sections by style:")
        for style, count in sorted(style_count.items()):
            self.logger.info(f"  {style}: {count}")
        self.logger.info(f"Image/Table captions: {image_count}")
        self.logger.info("=" * 50)
    
    def _is_image_or_table_caption(self, text: str) -> bool:
        """Check if text is likely an image or table caption."""
        caption_patterns = [
            r'^\[الشكل\s*\d+\]', r'^\[الصورة\s*\d+\]', r'^\[الجدول\s*\d+\]',
            r'^الشكل\s*\(\d+\)', r'^الصورة\s*\(\d+\)', r'^الجدول\s*\(\d+\)',
            r'^شكل\s*\d+', r'^صورة\s*\d+', r'^جدول\s*\d+',
            r'^\[Figure\s*\d+\]', r'^\[Table\s*\d+\]', r'^\[Image\s*\d+\]',
            r'^Figure\s*\d+', r'^Table\s*\d+', r'^Image\s*\d+'
        ]
        
        for pattern in caption_patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return True
        
        return False
    
    def _has_special_formatting(self, paragraph) -> bool:
        """Check if paragraph has special formatting (bold, larger font)."""
        if not paragraph.runs:
            return False
        
        first_run = paragraph.runs[0]
        if first_run.bold:
            return True
        
        if first_run.font.size and first_run.font.size > Pt(14):
            return True
        
        return False