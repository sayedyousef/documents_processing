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
    
    logging.basicConfig(
        level=getattr(logging, Config.LOG_LEVEL),
        format=Config.LOG_FORMAT,
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(sys.stdout)
        ]
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
        
        logger.info("=" * 60)
        logger.info("Processing Summary:")
        logger.info(f"  Documents processed: {len(documents)}")
        logger.info(f"  Total words: {total_words:,}")
        logger.info(f"  Total images: {total_images}")
        logger.info(f"  Total sections: {total_sections}")
        logger.info(f"  Output file: {Config.OUTPUT_FILE.absolute()}")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error during processing: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
