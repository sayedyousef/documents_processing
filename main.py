# main.py
"""Main entry point for document processing application."""

import logging
import sys
from pathlib import Path
import sys
import io

# Force UTF-8 encoding for Windows console
if sys.platform == 'win32':
    # Set console code page to UTF-8
    import os
    os.system('chcp 65001 >nul 2>&1')
    
    # Reconfigure stdout and stderr
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Then your regular imports...

#import sys
#from pathlib import Path
#sys.path.append(str(Path(__file__).parent))  # ← Add current directory to path
from datetime import datetime
from document_processor.document_processor import DocumentProcessor
from excel_writer import ExcelWriter
from config import Config

def setup_logging():
    """Configure logging for the application."""
    # Create logs directory if it doesn't exist
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Setup file and console logging
    log_filename = log_dir / f"document_processor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    # Create handlers with UTF-8 encoding
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(getattr(logging, Config.LOG_LEVEL))
    file_handler.setFormatter(logging.Formatter(Config.LOG_FORMAT))
    
    # Console handler with UTF-8
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, Config.LOG_LEVEL))
    console_handler.setFormatter(logging.Formatter(Config.LOG_FORMAT))
    
    # Configure root logger
    logging.basicConfig(
        level=getattr(logging, Config.LOG_LEVEL),
        handlers=[file_handler, console_handler]
    )

def validate_environment():
    """Validate that required directories exist."""
    if not Config.INPUT_FOLDER.exists():
        Config.INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        logging.warning(f"Created input folder: {Config.INPUT_FOLDER}")
        return False
    return True

def main():
    """Main function to orchestrate document processing."""
    setup_logging()
    logger = logging.getLogger(__name__)
    
    logger.info("=" * 60)
    logger.info("Document Processing Application Started")
    logger.info("=" * 60)
    
    # Validate environment
    if not validate_environment():
        logger.warning(f"No documents found. Please add .docx files to: {Config.INPUT_FOLDER}")
        return
    
    try:
        # Initialize processors
        doc_processor = DocumentProcessor()
        excel_writer = ExcelWriter(Config.OUTPUT_FILE)
        
        # Process documents
        logger.info(f"Scanning for documents in: {Config.INPUT_FOLDER}")
        documents = doc_processor.process_folder(Config.INPUT_FOLDER)
        
        if not documents:
            logger.warning("No documents found to process!")
            return
        
        # Write results to Excel
        logger.info(f"Found {len(documents)} documents. Writing to Excel...")
        excel_writer.write_summary(documents)
        excel_writer.write_sections(documents)
        excel_writer.save()
        
        # Summary statistics
        total_words = sum(doc.word_count for doc in documents)
        total_images = sum(doc.image_count for doc in documents)
        total_sections = sum(len(doc.sections) for doc in documents)
        total_headings = sum(len([s for s in doc.sections if s.section_type == 'heading']) for doc in documents)
        total_arabic_refs = sum(getattr(doc, 'arabic_reference_count', 0) for doc in documents)
        total_english_refs = sum(getattr(doc, 'english_reference_count', 0) for doc in documents)
        total_footnotes = sum(getattr(doc, 'footnote_count', 0) for doc in documents)
        
        # Format quality statistics
        quality_counts = {}
        for doc in documents:
            quality = getattr(doc, 'format_quality', 'Unknown')
            quality_counts[quality] = quality_counts.get(quality, 0) + 1
        
        poor_docs = [doc for doc in documents if getattr(doc, 'format_quality', '') == 'Poor']
        docs_with_issues = [doc for doc in documents if not doc.uses_proper_styles]
        docs_missing_captions = [doc for doc in documents if len(getattr(doc, 'images_missing_captions', [])) > 0]
        
        logger.info("=" * 60)
        logger.info("Processing Summary:")
        logger.info(f"  Documents processed: {len(documents)}")
        logger.info(f"  Total words: {total_words:,}")
        logger.info(f"  Total headings: {total_headings}")
        logger.info(f"  Total images: {total_images}")
        logger.info(f"  Total references: {total_arabic_refs + total_english_refs} (Arabic: {total_arabic_refs}, English: {total_english_refs})")
        logger.info(f"  Total footnotes: {total_footnotes}")
        logger.info(f"  Output file: {Config.OUTPUT_FILE.absolute()}")
        logger.info("=" * 60)
        
        # Format quality summary
        logger.info("\nFormat Quality Summary:")
        for quality in ['Excellent', 'Good', 'Fair', 'Poor']:
            count = quality_counts.get(quality, 0)
            if count > 0:
                logger.info(f"  {quality}: {count} documents ({count/len(documents)*100:.1f}%)")
        
        if docs_with_issues:
            logger.info(f"\nWARNING: {len(docs_with_issues)} documents not using proper heading styles")
        
        if docs_missing_captions:
            logger.info(f"\nIMAGES: {len(docs_missing_captions)} documents have images without captions")
        
        if poor_docs:
            logger.info(f"\nCRITICAL: {len(poor_docs)} documents need immediate formatting attention:")
            for doc in poor_docs[:5]:
                issues_count = getattr(doc, 'total_format_issues', 0)
                logger.info(f"    - {doc.name} ({issues_count} issues)")
            if len(poor_docs) > 5:
                logger.info(f"    ... and {len(poor_docs) - 5} more")
        
        logger.info("\nTIP: Check the 'Format Issues' sheet in the Excel file for detailed recommendations.")
        
    except Exception as e:
        logger.error(f"Error during processing: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()