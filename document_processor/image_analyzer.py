# image_analyzer.py
"""Module for analyzing images in documents."""

import logging
import re
from pathlib import Path
from docx import Document as DocxDocument

class ImageAnalyzer:
    """Handles image counting and analysis."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def count_images_total(self, docx: DocxDocument) -> int:
        """Count total images in document (including all formats/duplicates)."""
        unique_images = set()
        
        # Check document relationships
        for rel_id, rel in docx.part.rels.items():
            if "image" in rel.reltype:
                unique_images.add(rel.target_ref)
        
        # Check inline shapes
        try:
            from lxml import etree
            nsmap = {'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}
            
            for paragraph in docx.paragraphs:
                for pic in paragraph._element.xpath('.//pic:pic', namespaces=nsmap):
                    blip = pic.xpath('.//a:blip/@r:embed', 
                                   namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                                             'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
                    if blip:
                        unique_images.add(blip[0])
        except Exception as e:
            self.logger.warning(f"Error counting inline images: {e}")
        
        return len(unique_images)
    
    def count_unique_images(self, docx: DocxDocument, file_path: Path) -> int:
        """Count unique images by detecting PNG/SVG duplicate patterns."""
        # Collect all image references
        all_images = {}
        
        for rel_id, rel in docx.part.rels.items():
            if "image" in rel.reltype:
                filename = rel.target_ref.split('/')[-1]
                all_images[filename] = rel_id
        
        if not all_images:
            return 0
        
        self.logger.info(f"====================Processing: {file_path.name}====================")
        self.logger.info(f"  Found images: {sorted(all_images.keys())}")
        
        # Identify SVG duplicates to ignore
        images_to_ignore = set()
        
        for filename in all_images:
            if filename.endswith('.svg'):
                match = re.search(r'image(\d+)\.svg', filename)
                if match:
                    svg_num = int(match.group(1))
                    png_equivalent = f"image{svg_num - 1}.png"
                    
                    if png_equivalent in all_images:
                        images_to_ignore.add(filename)
                        self.logger.info(f"  Ignoring {filename} (duplicate of {png_equivalent})")
        
        unique_images = set(all_images.keys()) - images_to_ignore
        unique_count = len(unique_images)
        
        self.logger.info(f"  Unique images: {sorted(unique_images)}")
        self.logger.info(f"  Total files: {len(all_images)}, Unique images: {unique_count}")
        
        return unique_count