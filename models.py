# models.py
"""Data models for document processing."""

from dataclasses import dataclass, field
from typing import List, Optional, Dict
from pathlib import Path

@dataclass
class DocumentSection:
    """Represents a section within a document."""
    heading: str
    style_name: Optional[str] = None
    section_type: str = "text"  # "heading", "image", or "table"
    text: str = ""
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    suggested_style: Optional[str] = None  # Suggested Word style
    has_caption: bool = True  # For images/tables
    
@dataclass
class Document:
    """Represents a processed document."""
    # Required fields (no defaults)
    id: int
    file_path: Path
    name: str
    parent_folder: str
    title: str
    author: str  # From document properties
    word_count: int
    image_count: int
    unique_image_count: int
    
    # Optional fields (with defaults)
    author_from_text: str = "Unknown"  # From 2nd paragraph
    sections: List[DocumentSection] = field(default_factory=list)
    uses_proper_styles: bool = True
    
    # Reference and footnote counts
    arabic_reference_count: int = 0
    english_reference_count: int = 0
    footnote_count: int = 0
    
    # Enhanced format tracking attributes
    format_quality: str = "Unknown"  # 'Excellent', 'Good', 'Fair', 'Poor'
    format_issues: List[str] = field(default_factory=list)
    images_missing_captions: List[str] = field(default_factory=list)
    heading_hierarchy_issues: List[Dict] = field(default_factory=list)
    heading_stats: Dict[str, int] = field(default_factory=dict)
    
    @property
    def filename(self) -> str:
        """Get the filename without extension."""
        return self.file_path.stem
    
    @property
    def total_format_issues(self) -> int:
        """Get total count of all format issues."""
        return (len(self.format_issues) + 
                len(self.images_missing_captions) + 
                len(self.heading_hierarchy_issues))
    
    @property
    def total_references(self) -> int:
        """Get total reference count."""
        return self.arabic_reference_count + self.english_reference_count