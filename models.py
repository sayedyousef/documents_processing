# models.py
"""Data models for document processing."""

from dataclasses import dataclass, field
from typing import List, Optional
from pathlib import Path

@dataclass
class DocumentSection:
    """Represents a section within a document."""
    heading: str
    style_name: Optional[str] = None
    section_type: str = "text"  # "text", "image", or "table"
    text: str = ""
    
@dataclass
class Document:
    """Represents a processed document."""
    id: int
    file_path: Path
    name: str
    parent_folder: str
    title: str
    author: str
    word_count: int
    image_count: int
    unique_image_count: int
    sections: List[DocumentSection] = field(default_factory=list)
    uses_proper_styles: bool = True
    
    @property
    def filename(self) -> str:
        """Get the filename without extension."""
        return self.file_path.stem