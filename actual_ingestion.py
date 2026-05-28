#!/usr/bin/env python3
"""
Actual book ingestion - reads each book, extracts knowledge entries, inserts into DB.
"""

import os
import json
import sys
from pathlib import Path
from ebooklib import epub
from bs4 import BeautifulSoup
import PyPDF2

# Import corpus builder
sys.path.insert(0, str(Path(__file__).parent))
from code import batch_insert_validated, render_insert_report

def extract_epub_text(filepath):
    """Extract all text from an EPUB file, sectioned by chapter."""
    book = epub.read_epub(filepath)
    sections = []
    current_chapter = []
    chapter_title = ""
    
    for item in book.get_items_of_type(9):  # TEXT items
        soup = BeautifulSoup(item.content, 'html.parser')
        text = soup.get_text(separator='\n', strip=True)
        text = text.strip()
        
        if not text or len(text) < 50:
            continue
            
        # Detect chapter boundaries
        headings = soup.find_all(['h1', 'h2', 'h3', 'h4'])
        for h in headings:
            h_text = h.get_text(strip=True)
            if h_text and len(h_text) > 5 and len(h_text) < 100:
                if current_chapter:
                    sections.append((chapter_title, '\n'.join(current_chapter)))
                    current_chapter = []
                chapter_title = h_text
                continue
        
        current_chapter.append(text)
    
    if current_chapter:
        sections.append((chapter_title, '\n'.join(current_chapter)))
    
    return sections

def extract_pdf_text(filepath):
    """Extract all text from a PDF file, sectioned by page."""
    sections = []
    try:
        with open(filepath, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text and len(text.strip()) > 100:
                    sections.append((f"Page {i+1}", text.strip()))
    except Exception as e:
        print(f"  Error reading PDF {filepath}: {e}")
    return sections

def extract_knowledge_from_text(title, text, source_label):
    """Extract knowledge entries from extracted text content."""
    entries = []
    lines = text.split('\n')
    
    # Look for code patterns, formulas, techniques
    in_code_block = False
    code_lines = []
    code_context = ""
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Detect code blocks
        if any(stripped.startswith(kw) for kw in ['import ', 'def ', 'df.', 'pd.', 'plt.', 'read_sheet', 'save_', 'print(', 'import pandas', 'import openpyxl']):
            if not in_code_block:
                in_code_block = True
                code_lines = [stripped]
                # Get context - look back for topic
                context_lines = []
                for j in range(max(0, i-5), i):
                    if lines[j].strip():
                        context_lines.append(lines[j].strip())
                code_context = ' '.join(context_lines[-3:]) if context_lines else ""
            else:
                code_lines.append(stripped)
        elif in_code_block and stripped and not stripped.startswith('#'):
            code_lines.append(stripped)
        elif in_code_block and (not stripped or stripped.startswith('#')):
            if len(code_lines) >= 2:
                code_text = '\n'.join(code_lines)
                entry = generate_entry_from_code(code_text, code_context, title, source_label)
                if entry:
                    entries.append(entry)
            in_code_block = False
            code_lines = []
            code_context = ""
        elif in_code_block and not stripped:
            if len(code_lines) >= 2:
                code_text = '\n'.join(code_lines)
                entry = generate_entry_from_code(code_text, code_context, title, source_label)
                if entry:
                    entries.append(entry)
            in_code_block = False
            code_lines = []
            code_context = ""
    
    # Handle code block at end of text
    if in_code_block and len(code_lines) >= 2:
        code_text = '\n'.join(code_lines)
        entry = generate_entry_from_code(code_text, code_context, title, source_label)
        if entry:
            entries.append(entry)
    
    # Also look for formula patterns
    formula_patterns = [
        (r'=(SUM|AVERAGE|IF|VLOOKUP|XLOOKUP|INDEX|MATCH|COUNT|CONCATENATE|LEFT|RIGHT|MID|TRIM|ROUND|DATE|NPV|IRR|PMT)\s*\(', 'formula'),
        (r'VLOOKUP.*:', 'lookup'),
        (r'XLOOKUP.*:', 'lookup'),
        (r'INDEX.*MATCH', 'lookup'),
    ]
    
    for line in lines:
        stripped = line.strip()
        for pattern, ptype in formula_patterns:
            import re
            if re.search(pattern, stripped, re.IGNORECASE):
                # Extract the formula name
                m = re.search(r'(SUM|AVERAGE|IF|VLOOKUP|XLOOKUP|INDEX|MATCH|COUNT|CONCATENATE|LEFT|RIGHT|MID|TRIM|ROUND|DATE|NPV|IRR|PMT)\s*\(', stripped, re.IGNORECASE)
                if m:
                    func_name = m.group(1).upper()
                    entry = generate_formula_entry(func_name, stripped, title, source_label)
                    if entry and entry not in entries:
                        entries.append(entry)
    
    return entries

def generate_entry_from_code(code_text, context, book_title, source_label):
    """Generate a knowledge entry from a code snippet."""
    # Determine category
    code_lower = code_text.lower()
    if 'pandas' in code_lower or 'df.' in code_lower or 'groupby' in code_lower or 'merge' in code_lower:
        category = 'pandas'
        subcategory = 'data_manipulation'
    elif 'plt.' in code_lower or 'matplotlib' in code_lower or 'chart' in code_lower or 'plot' in code_lower:
        category = 'visualisation'
        subcategory = 'chart'
    elif 'excel' in code_lower or 'read_sheet' in code_lower or 'save_' in code_lower:
        category = 'excel'
        subcategory = 'automation'
    elif 'sql' in code_lower or 'sqlite' in code_lower:
        category = 'sql'
        subcategory = 'database'
    elif 'openpyxl' in code_lower or 'load_workbook' in code_lower or 'Workbook' in code_lower:
        category = 'excel'
        subcategory = 'file_operations'
    else:
        category = 'analysis'
        subcategory = 'general'
    
    # Generate title from context or code
    title = ""
    if context:
        # Use last meaningful context line
        ctx_parts = context.split()
        if len(ctx_parts) >= 3:
            title = ' '.join(ctx_parts[-6:])[:80]
        else:
            title = context[:80]
    else:
        # Generate from code
        if 'read_sheet' in code_text:
            title = "Using read_sheet() to load Excel data"
        elif 'groupby' in code_text:
            title = "Using groupby for data aggregation"
        elif 'merge' in code_text:
            title = "Merging dataframes for data combination"
        elif 'pivot' in code_text:
            title = "Creating pivot tables for data summarization"
        elif 'plot' in code_text or 'chart' in code_text:
            title = "Creating charts and visualizations"
        elif 'save_' in code_text:
            title = "Saving processed data to Excel files"
        else:
            title = "Excel automation with Python"
    
    # Ensure title meets requirements
    if len(title) < 10:
        title = f"Python Excel technique: {title}"
    if len(title) > 120:
        title = title[:117] + "..."
    
    # Generate problem description
    problem = f"This technique demonstrates how to use Python with Excel for {subcategory.replace('_', ' ')} tasks."
    
    return {
        "title": title,
        "problem": problem,
        "code": code_text,
        "category": category,
        "subcategory": subcategory,
        "tags": [category, subcategory, "python", "excel"],
        "notes": f"Extracted from {book_title}",
        "source": f"{source_label}, Chapter/Section context",
        "difficulty": "intermediate"
    }

def generate_formula_entry(func_name, formula_text, book_title, source_label):
    """Generate a knowledge entry from a formula pattern."""
    # Map function names to categories
    func_lower = func_name.lower()
    
    if func_lower in ['sum', 'average', 'count', 'max', 'min']:
        category = 'excel'
        subcategory = 'formula'
        tags = ['formula', func_lower.lower(), 'calculation']
    elif func_lower in ['vlookup', 'xlookup', 'index', 'match']:
        category = 'excel'
        subcategory = 'lookup'
        tags = ['lookup', func_lower.lower(), 'data_matching']
    elif func_lower == 'if':
        category = 'excel'
        subcategory = 'formula'
        tags = ['conditional', 'if', 'logic']
    elif func_lower in ['concatenate', 'left', 'right', 'mid', 'trim']:
        category = 'excel'
        subcategory = 'text'
        tags = ['text', func_lower.lower(), 'string']
    elif func_lower in ['round', 'date']:
        category = 'excel'
        subcategory = 'formula'
        tags = ['date', func_lower.lower(), 'numeric']
    elif func_lower in ['npv', 'irr', 'pmt']:
        category = 'analysis'
        subcategory = 'financial'
        tags = ['financial', func_lower.lower(), 'modeling']
    else:
        category = 'excel'
        subcategory = 'formula'
        tags = ['formula', func_lower.lower()]
    
    title = f"Excel {func_name} Function"
    if len(formula_text) > 20:
        title = f"Excel {func_name} Function - {formula_text[:60]}"
    
    return {
        "title": title,
        "problem": f"Need to {func_name.lower()} values or data in Excel worksheets.",
        "code": f"# Excel {func_name} function example\n# Use in Excel: ={func_name}(range)\n# Or with pandas:\nimport pandas as pd\nfrom excel_mcp_v2 import read_sheet\ndf = read_sheet('Sheet1')\nresult = df['column'].{'sum' if func_lower == 'sum' else 'mean' if func_lower == 'average' else 'count'}()\nprint(result)",
        "category": category,
        "subcategory": subcategory,
        "tags": tags,
        "notes": f"Formula pattern found in {book_title}",
        "source": f"{source_label}, formula reference",
        "difficulty": "beginner"
    }

def ingest_book(filepath):
    """Ingest a single book and return entries."""
    entries = []
    source_label = os.path.basename(filepath)
    
    print(f"  Processing: {source_label[:60]}...")
    
    if filepath.endswith('.epub'):
        sections = extract_epub_text(filepath)
    elif filepath.endswith('.pdf'):
        sections = extract_pdf_text(filepath)
    else:
        return entries
    
    print(f"    Extracted {len(sections)} sections")
    
    # Process each section
    for section_title, text in sections:
        # Only process substantial sections
        if len(text) < 500:
            continue
        
        section_entries = extract_knowledge_from_text(section_title, text, source_label)
        entries.extend(section_entries)
    
    print(f"    Found {len(entries)} knowledge entries")
    return entries

def main():
    """Main ingestion loop."""
    print("=" * 60)
    print("EXCEL MCP v2 - AUTOMATED BOOK INGESTION")
    print("=" * 60)
    print()
    
    # Get all books
    books = sorted([f for f in os.listdir('.') if f.endswith(('.epub', '.pdf'))])
    print(f"Found {len(books)} books to process")
    print()
    
    all_entries = []
    book_results = []
    
    for i, book in enumerate(books, 1):
        print(f"\n[{i}/{len(books)}] Ingesting: {book}")
        
        try:
            entries = ingest_book(book)
            all_entries.extend(entries)
            book_results.append({
                'book': book,
                'entries': len(entries)
            })
            print(f"    -> Added {len(entries)} entries")
        except Exception as e:
            print(f"    -> ERROR: {e}")
            book_results.append({
                'book': book,
                'entries': 0,
                'error': str(e)
            })
    
    print("\n" + "=" * 60)
    print(f"INGESTION SUMMARY")
    print("=" * 60)
    print(f"Total books processed: {len(book_results)}")
    print(f"Total entries extracted: {len(all_entries)}")
    print()
    
    # Show entries per book
    for r in book_results:
        status = f"{r['entries']} entries"
        if 'error' in r:
            status += f" (ERROR: {r['error'][:40]})"
        print(f"  {r['book'][:55]:55s} -> {status}")
    
    print()
    
    if all_entries:
        print("Inserting all entries into knowledge base...")
        try:
            result = batch_insert_validated(
                entries=all_entries,
                db_path="excel_knowledge.db",
                source_label="Automated Book Ingestion"
            )
            
            print("\n" + render_insert_report(result, "Automated Book Ingestion"))
            
            # Update progress tracking
            with open('ingestion_log.json', 'w') as f:
                json.dump({
                    'total_books': len(books),
                    'processed': len(book_results),
                    'total_entries': len(all_entries),
                    'inserted': result.get('inserted', 0),
                    'skipped': result.get('skipped_dup', 0),
                    'book_results': book_results
                }, f, indent=2)
            print("\nIngestion log saved to ingestion_log.json")
            
        except Exception as e:
            print(f"ERROR during insertion: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("No entries were extracted. Check book content extraction.")

if __name__ == "__main__":
    main()
