# config.py - Updated with console encoding option
"""Configuration settings for document processor."""

from pathlib import Path

class Config:
    """Application configuration."""
    
    # Paths
    INPUT_FOLDER = Path("./documents")
    INPUT_FOLDER = Path("D:\\Work 3 (20-Oct-24)\\2 Side projects May 25\\Encyclopedia\\articles\\مقالات بعد الاخراج")
    OUTPUT_FILE = Path("document_analysis.xlsx")
    
    # Processing settings
    MAX_TEXT_PREVIEW_LENGTH = 200
    MAX_SHEET_NAME_LENGTH = 31
    
    # Font detection thresholds
    HEADING_MIN_FONT_SIZE = 14  # Points
    
    # Excel formatting
    HEADER_BG_COLOR = "366092"
    HEADER_FONT_COLOR = "FFFFFF"
    SECTION_HEADER_BG_COLOR = "CCCCCC"
    
    # Logging
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOG_LEVEL = 'INFO'
    
    # Console output settings
    SIMPLE_CONSOLE_OUTPUT = True  # Set to True to avoid Unicode issues in Windows console
    VERBOSE_FILE_LOGGING = True   # Full details in log files