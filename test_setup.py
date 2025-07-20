# test_setup.py
"""Test script to verify installation and setup."""

import sys

def test_imports():
    """Test if all required packages are installed."""
    print("Testing imports...")
    
    required_packages = [
        ('docx', 'python-docx'),
        ('openpyxl', 'openpyxl'),
        ('pandas', 'pandas'),
        ('lxml', 'lxml')
    ]
    
    all_good = True
    for module_name, package_name in required_packages:
        try:
            __import__(module_name)
            print(f"✓ {package_name} is installed")
        except ImportError:
            print(f"✗ {package_name} is NOT installed")
            all_good = False
    
    return all_good

def test_modules():
    """Test if all application modules are present."""
    print("\nTesting application modules...")
    
    modules = [
        'models',
        'document_processor',
        'excel_writer',
        'config',
        'utils'
    ]
    
    all_good = True
    for module in modules:
        try:
            __import__(module)
            print(f"✓ {module}.py found")
        except ImportError:
            print(f"✗ {module}.py NOT found")
            all_good = False
    
    return all_good

def main():
    """Run all tests."""
    print("Document Processor Setup Test")
    print("=" * 40)
    
    imports_ok = test_imports()
    modules_ok = test_modules()
    
    print("\n" + "=" * 40)
    if imports_ok and modules_ok:
        print("✓ All tests passed! Ready to process documents.")
        print("\nUsage:")
        print("1. Place your .docx files in the 'documents' folder")
        print("2. Run: python main.py")
        print("3. Check 'document_analysis.xlsx' for results")
    else:
        print("✗ Some tests failed. Please check the errors above.")
        sys.exit(1)

if __name__ == "__main__":
    main()


# === USAGE GUIDE ===
"""
Document Processor - Usage Guide
================================

SETUP:
1. Create project directory structure:
   document_processor/
   ├── main.py
   ├── models.py
   ├── document_processor.py
   ├── excel_writer.py
   ├── config.py
   ├── utils.py
   ├── test_setup.py
   ├── requirements.txt
   └── documents/          (create this folder)

2. Install dependencies:
   pip install -r requirements.txt

3. Test setup:
   python test_setup.py

USAGE:
1. Place Word documents (.docx) in the 'documents' folder
   - Supports subfolders
   - Handles Arabic and English documents

2. Run the processor:
   python main.py

3. Check output:
   - 'document_analysis.xlsx' will be created
   - 'logs' folder will contain processing logs

OUTPUT:
- Summary sheet: Document list with ID, name, title, word count, images
- Individual sheets: Section details for each document

CUSTOMIZATION:
- Edit config.py to change settings
- Modify font size thresholds for section detection
- Adjust output formatting

TROUBLESHOOTING:
- Check logs folder for detailed error messages
- Ensure documents are not corrupted
- Verify all .docx files are closed before processing
"""
