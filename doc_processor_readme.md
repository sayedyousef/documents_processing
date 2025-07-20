# Document Processor

A modular Python application for extracting and analyzing content from Word documents (.docx files), with support for both Arabic and English documents.

## Features

- **Batch Processing**: Process multiple Word documents from folders and subfolders
- **Comprehensive Analysis**: 
  - Extract document metadata (title, word count, image count)
  - Identify document sections with font information
  - Handle both Arabic and English text
- **Excel Export**: 
  - Summary sheet with all documents
  - Individual sheets for each document's sections
- **Modular Design**: Clean, object-oriented architecture following KISS principle
- **Robust Error Handling**: Detailed logging and error recovery

## Project Structure

```
document_processor/
├── main.py                 # Application entry point
├── models.py              # Data models (Document, Section)
├── document_processor.py   # Core processing logic
├── excel_writer.py        # Excel export functionality
├── config.py              # Configuration settings
├── utils.py               # Utility functions
├── test_setup.py          # Setup verification script
├── requirements.txt       # Python dependencies
├── documents/             # Input folder for .docx files
└── logs/                  # Processing logs (auto-created)
```

## Installation

1. **Clone or download the project files**

2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   
   # Activate virtual environment:
   # Windows:
   venv\Scripts\activate
   # Mac/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Verify installation**:
   ```bash
   python test_setup.py
   ```

## Usage

1. **Place your Word documents** in the `documents/` folder
   - Supports nested subfolders
   - Only processes `.docx` files (not `.doc`)

2. **Run the processor**:
   ```bash
   python main.py
   ```

3. **Check the output**:
   - `document_analysis.xlsx` - Contains all extracted data
   - `logs/` folder - Contains detailed processing logs

## Output Format

### Summary Sheet
Contains overview of all processed documents:
- **ID**: Sequential document identifier
- **Document Name**: Original filename
- **Document Title**: Extracted title (from properties or first heading)
- **Word Count**: Total words in document
- **Image Count**: Number of images found

### Individual Document Sheets
Each document gets its own sheet containing:
- Document metadata (ID, name, title)
- Section details:
  - **Section Heading**: Text of the section header
  - **Font Name**: Font used for the heading
  - **Font Size**: Size of the heading font
  - **Text Preview**: First 200 characters of section content

## Configuration

Edit `config.py` to customize:

```python
# Input/Output paths
INPUT_FOLDER = Path("./documents")
OUTPUT_FILE = Path("document_analysis.xlsx")

# Processing settings
MAX_TEXT_PREVIEW_LENGTH = 200
HEADING_MIN_FONT_SIZE = 14  # Points

# Excel formatting colors
HEADER_BG_COLOR = "366092"
```

## Architecture

The application follows clean architecture principles:

- **Separation of Concerns**: Each module has a single responsibility
- **KISS Principle**: Simple, straightforward implementations
- **Object-Oriented**: Clean class structures without over-engineering
- **Modular Design**: Easy to extend or modify individual components

### Key Components

1. **Models** (`models.py`):
   - `Document`: Represents a processed document
   - `DocumentSection`: Represents a section within a document

2. **DocumentProcessor** (`document_processor.py`):
   - Handles Word document parsing
   - Extracts text, images, and sections
   - Identifies headings based on styles and formatting

3. **ExcelWriter** (`excel_writer.py`):
   - Creates formatted Excel workbooks
   - Handles sheet creation and data writing
   - Auto-adjusts column widths

4. **Utilities** (`utils.py`):
   - Text cleaning and processing
   - Language detection (Arabic/English)
   - Helper functions

## Troubleshooting

### Common Issues

1. **No documents found**:
   - Ensure files have `.docx` extension (not `.doc`)
   - Check that files are in the `documents/` folder
   - Look for error messages in the logs

2. **Image count seems incorrect**:
   - The processor counts unique images
   - Embedded vs linked images may affect count
   - Check logs for image counting warnings

3. **Sections not detected properly**:
   - Adjust `HEADING_MIN_FONT_SIZE` in config.py
   - Ensure headings use Word's heading styles
   - Check if headings are bold or have larger fonts

4. **Excel file won't open**:
   - Ensure the output file isn't already open
   - Check for special characters in document names
   - Review logs for Excel writing errors

### Debug Mode

For more detailed logging, edit `config.py`:
```python
LOG_LEVEL = 'DEBUG'  # Instead of 'INFO'
```

## Extending the Application

### Adding New Features

1. **New document properties**: 
   - Add fields to the `Document` class in `models.py`
   - Update extraction logic in `document_processor.py`
   - Add columns in `excel_writer.py`

2. **Additional export formats**:
   - Create new writer classes (e.g., `CSVWriter`, `JSONWriter`)
   - Follow the same interface as `ExcelWriter`

3. **Custom section detection**:
   - Modify `_has_special_formatting()` method
   - Add new detection criteria

## Performance

- Processes approximately 10-20 documents per minute
- Memory efficient - processes one document at a time
- Suitable for batches up to 1000+ documents

## License

This project is provided as-is for educational and commercial use.

## Support

For issues or questions:
1. Check the logs in the `logs/` folder
2. Run `test_setup.py` to verify installation
3. Review the troubleshooting section above
