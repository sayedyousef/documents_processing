# models.py
"""Data models for document processing."""

from dataclasses import dataclass, field
from typing import List, Optional
from pathlib import Path

@dataclass
class DocumentSection:
    """Represents a section within a document."""
    heading: str
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    text: str = ""
    
@dataclass
class Document:
    """Represents a processed document."""
    id: int
    file_path: Path
    name: str
    parent_folder: str
    title: str
    word_count: int
    image_count: int
    unique_image_count: int
    sections: List[DocumentSection] = field(default_factory=list)
    
    @property
    def filename(self) -> str:
        """Get the filename without extension."""
        return self.file_path.stem
