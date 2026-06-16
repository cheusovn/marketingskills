#!/usr/bin/env python3
"""
Document processor: Convert PDFs to DOCX with name replacements
Maintains logical page structure and sequential order
"""
import os
import re
from pathlib import Path
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
import pytesseract
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Configuration
UPLOAD_DIR = '/root/.claude/uploads/3e540a7b-5ab9-5cc8-9677-e3d80f77fa1e'
OUTPUT_DIR = '/home/user/marketingskills'

# Replacement patterns
REPLACEMENTS = [
    (r'ИП Чеусова Николая ИГоревича', 'ИП Пугачеву Ольгу Николаевну'),
    (r'ИП Чеусова Николая Игоревича', 'ИП Пугачеву Ольгу Николаевну'),
    (r'Чеусова Николая ИГоревича', 'Пугачеву Ольгу Николаевну'),
    (r'Чеусова Николая Игоревича', 'Пугачеву Ольгу Николаевну'),
    (r'Чеусова', 'Пугачева'),
    (r'Николай Игоревич', 'Ольга Николаевна'),
    (r'Николай ИГоревич', 'Ольга Николаевна'),
]

def extract_text_direct(pdf_path):
    """Try to extract text directly from PDF."""
    text_by_page = []
    try:
        pdf_reader = PdfReader(pdf_path)
        for page_num, page in enumerate(pdf_reader.pages, 1):
            page_text = page.extract_text()
            text_by_page.append({
                'page': page_num,
                'text': page_text if page_text else '',
                'source': 'direct'
            })
    except Exception as e:
        print(f"  Warning: Direct extraction failed: {e}")
    return text_by_page

def extract_text_ocr(pdf_path):
    """Extract text from PDF using OCR, page by page."""
    text_by_page = []
    try:
        images = convert_from_path(pdf_path, dpi=200)
        for page_num, image in enumerate(images, 1):
            print(f"    Processing page {page_num}/{len(images)} with OCR...")
            page_text = pytesseract.image_to_string(image, lang='rus+eng')
            text_by_page.append({
                'page': page_num,
                'text': page_text if page_text else '',
                'source': 'ocr'
            })
    except Exception as e:
        print(f"  OCR extraction failed: {e}")
    return text_by_page

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file, maintaining page structure."""
    print(f"Extracting text from: {os.path.basename(pdf_path)}")

    # Try direct extraction first
    text_by_page = extract_text_direct(pdf_path)

    # Check if we got sufficient text
    total_chars = sum(len(p['text']) for p in text_by_page)

    # If direct extraction failed or returned minimal text, use OCR
    if total_chars < 100:
        print(f"  Direct extraction insufficient ({total_chars} chars), using OCR...")
        text_by_page = extract_text_ocr(pdf_path)
        total_chars = sum(len(p['text']) for p in text_by_page)

    if text_by_page:
        print(f"✓ Extracted {total_chars} characters from {len(text_by_page)} page(s)")
        for page_info in text_by_page:
            if page_info['text'].strip():
                print(f"  Page {page_info['page']} ({page_info['source']}): {len(page_info['text'])} chars")
    else:
        print(f"⚠ No text could be extracted")

    return text_by_page

def replace_names_in_text(text):
    """Replace all instances of the names."""
    for old, new in REPLACEMENTS:
        text = re.sub(old, new, text)
    return text

def add_page_break(doc):
    """Add a page break to the document."""
    doc.add_page_break()

def create_word_document_structured(text_by_page, output_path):
    """Create a Word document maintaining page structure and logical order."""
    print(f"Creating structured Word document...")

    doc = Document()

    # Set default font and styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0, 0, 0)

    # Process each page
    for page_idx, page_data in enumerate(text_by_page):
        page_num = page_data['page']
        page_text = page_data['text'].strip()

        if not page_text:
            continue

        # Replace names in this page
        page_text = replace_names_in_text(page_text)

        # Split into logical blocks (paragraphs separated by double newlines)
        # First try splitting by double newlines
        blocks = page_text.split('\n\n')

        for block_idx, block in enumerate(blocks):
            if not block.strip():
                continue

            # Clean up the block
            lines = [line.strip() for line in block.split('\n') if line.strip()]

            # Combine lines that are part of the same paragraph
            block_text = ' '.join(lines)

            if block_text:
                # Add paragraph with proper formatting
                p = doc.add_paragraph()
                p.text = block_text

                # Set paragraph formatting
                p.paragraph_format.space_after = Pt(12)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.right_indent = Inches(0)

        # Add page break after each page (except the last one)
        if page_idx < len(text_by_page) - 1:
            # Check if there's content on the next page
            if page_idx + 1 < len(text_by_page) and text_by_page[page_idx + 1]['text'].strip():
                add_page_break(doc)

    # Save
    try:
        doc.save(output_path)
        print(f"✓ Saved: {os.path.basename(output_path)}")
        return True
    except Exception as e:
        print(f"Error saving document: {e}")
        return False

def main():
    """Main processing function."""
    print("=" * 75)
    print("Document Converter - PDF to DOCX (with page structure preservation)")
    print("=" * 75)

    # Find PDF files
    pdf_files = sorted([f for f in os.listdir(UPLOAD_DIR) if f.endswith('.pdf')])

    if not pdf_files:
        print(f"Error: No PDF files found in {UPLOAD_DIR}")
        return

    print(f"\nFound {len(pdf_files)} PDF file(s):\n")
    for i, f in enumerate(pdf_files, 1):
        size = os.path.getsize(os.path.join(UPLOAD_DIR, f)) / 1024
        print(f"  {i}. {f}")
        print(f"     Size: {size:.1f} KB")

    results = []

    # Process each PDF
    for i, pdf_file in enumerate(pdf_files, 1):
        pdf_path = os.path.join(UPLOAD_DIR, pdf_file)

        print(f"\n{'=' * 75}")
        print(f"Processing ({i}/{len(pdf_files)}): {pdf_file}")
        print('=' * 75)

        # Extract text maintaining page structure
        text_by_page = extract_text_from_pdf(pdf_path)

        if not text_by_page or all(not p['text'].strip() for p in text_by_page):
            print(f"Skipping {pdf_file} - no text could be extracted")
            continue

        # Determine output filename
        if 'ekvizit' in pdf_file.lower():
            output_name = 'Реквизиты_Пугачева.docx'
        else:
            output_name = 'Договор_Пугачева.docx'

        output_path = os.path.join(OUTPUT_DIR, output_name)

        # Remove old file if exists
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass

        # Create Word document with page structure
        if create_word_document_structured(text_by_page, output_path):
            results.append({
                'path': output_path,
                'name': output_name,
                'pages': len([p for p in text_by_page if p['text'].strip()])
            })

    # Summary
    print("\n" + "=" * 75)
    if results:
        print(f"✅ Conversion complete!")
        print(f"\nCreated {len(results)} document(s):\n")
        for doc_info in results:
            size = os.path.getsize(doc_info['path']) / 1024
            print(f"  ✓ {doc_info['name']}")
            print(f"    Size: {size:.1f} KB | Pages: {doc_info['pages']}")
            print()
    else:
        print("⚠ No documents were created")
    print("=" * 75)

if __name__ == '__main__':
    main()
