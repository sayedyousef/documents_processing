# batch_processor.py
"""Enhanced batch processing with progress tracking and resume capability."""

import json
import time
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional
import logging

from document_processor.document_processor import DocumentProcessor
from excel_writer import ExcelWriter
from models import Document
from config import Config

class BatchProcessor:
    """Enhanced batch processor with progress tracking."""
    
    def __init__(self, checkpoint_file: Path = Path("processing_checkpoint.json")):
        self.checkpoint_file = checkpoint_file
        self.processed_files = set()
        self.failed_files = {}
        self.logger = logging.getLogger(__name__)
        
    def process_with_progress(self, input_folder: Path, output_file: Path):
        """Process documents with progress tracking and resume capability."""
        # Load checkpoint if exists
        self.load_checkpoint()
        
        # Get all documents
        all_files = list(input_folder.rglob("*.docx"))
        all_files = [f for f in all_files if not f.name.startswith("~")]
        
        # Filter out already processed
        pending_files = [f for f in all_files if str(f) not in self.processed_files]
        
        self.logger.info(f"Total files: {len(all_files)}")
        self.logger.info(f"Already processed: {len(self.processed_files)}")
        self.logger.info(f"Pending: {len(pending_files)}")
        
        if not pending_files:
            self.logger.info("All files already processed!")
            return
        
        # Process documents
        doc_processor = DocumentProcessor()
        documents = []
        start_time = time.time()
        
        for idx, file_path in enumerate(pending_files, 1):
            # Progress indicator
            progress = (idx + len(self.processed_files)) / len(all_files) * 100
            elapsed = time.time() - start_time
            eta = (elapsed / idx) * (len(pending_files) - idx) if idx > 0 else 0
            
            self.logger.info(f"Processing [{idx}/{len(pending_files)}] "
                           f"({progress:.1f}% total) - {file_path.name} "
                           f"- ETA: {self.format_time(eta)}")
            
            try:
                doc = doc_processor.process_document(file_path)
                documents.append(doc)
                self.processed_files.add(str(file_path))
                
                # Save checkpoint every 10 documents
                if idx % 10 == 0:
                    self.save_checkpoint()
                    
            except Exception as e:
                self.logger.error(f"Failed to process {file_path}: {e}")
                self.failed_files[str(file_path)] = str(e)
        
        # Write to Excel
        if documents:
            self.logger.info("Writing results to Excel...")
            excel_writer = ExcelWriter(output_file)
            excel_writer.write_summary(documents)
            excel_writer.write_sections(documents)
            excel_writer.save()
        
        # Save final checkpoint
        self.save_checkpoint()
        
        # Report summary
        self.print_summary(len(all_files), len(documents), time.time() - start_time)
    
    def load_checkpoint(self):
        """Load processing checkpoint if exists."""
        if self.checkpoint_file.exists():
            try:
                with open(self.checkpoint_file, 'r') as f:
                    data = json.load(f)
                    self.processed_files = set(data.get('processed', []))
                    self.failed_files = data.get('failed', {})
                self.logger.info(f"Loaded checkpoint: {len(self.processed_files)} files processed")
            except Exception as e:
                self.logger.warning(f"Failed to load checkpoint: {e}")
    
    def save_checkpoint(self):
        """Save current processing state."""
        try:
            data = {
                'processed': list(self.processed_files),
                'failed': self.failed_files,
                'timestamp': datetime.now().isoformat()
            }
            with open(self.checkpoint_file, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.logger.error(f"Failed to save checkpoint: {e}")
    
    def reset_checkpoint(self):
        """Reset processing checkpoint."""
        self.processed_files.clear()
        self.failed_files.clear()
        if self.checkpoint_file.exists():
            self.checkpoint_file.unlink()
        self.logger.info("Checkpoint reset")
    
    def format_time(self, seconds: float) -> str:
        """Format seconds into human-readable time."""
        if seconds < 60:
            return f"{seconds:.0f}s"
        elif seconds < 3600:
            return f"{seconds/60:.1f}m"
        else:
            return f"{seconds/3600:.1f}h"
    
    def print_summary(self, total: int, processed: int, elapsed: float):
        """Print processing summary."""
        self.logger.info("=" * 60)
        self.logger.info("BATCH PROCESSING SUMMARY")
        self.logger.info("=" * 60)
        self.logger.info(f"Total files found: {total}")
        self.logger.info(f"Successfully processed: {processed}")
        self.logger.info(f"Failed: {len(self.failed_files)}")
        self.logger.info(f"Processing time: {self.format_time(elapsed)}")
        self.logger.info(f"Average time per document: {elapsed/processed:.1f}s")
        
        if self.failed_files:
            self.logger.info("\nFailed files:")
            for file, error in self.failed_files.items():
                self.logger.info(f"  - {Path(file).name}: {error}")


# batch_main.py
"""Entry point for batch processing with resume capability."""

import argparse
import logging
import sys
from pathlib import Path
from batch_processor import BatchProcessor
from config import Config

def main():
    """Main function for batch processing."""
    parser = argparse.ArgumentParser(description="Batch process Word documents")
    parser.add_argument('--input', type=Path, default=Config.INPUT_FOLDER,
                       help='Input folder containing documents')
    parser.add_argument('--output', type=Path, default=Config.OUTPUT_FILE,
                       help='Output Excel file')
    parser.add_argument('--reset', action='store_true',
                       help='Reset checkpoint and start fresh')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Setup logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format=Config.LOG_FORMAT,
        handlers=[
            logging.FileHandler(f"batch_processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Run batch processor
    processor = BatchProcessor()
    
    if args.reset:
        processor.reset_checkpoint()
    
    processor.process_with_progress(args.input, args.output)

if __name__ == "__main__":
    main()
