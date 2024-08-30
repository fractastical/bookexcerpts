import re
import json
from docx import Document
from collections import defaultdict

def extract_citations(doc):
    # Pattern for matching "author (date) title" format
    pattern = r'^\s*(.*?)\s*\((\d{4})\)\s*(.*?)\.?\s*$'
    citations = []
    chapter_citation_count = defaultdict(int)
    current_chapter_number = 1  # Start with Chapter 1
    headings = {
        "chapter": f"Chapter {current_chapter_number}: Start",  # Initialize as Chapter 1
        "heading2": "",
        "heading3": "",
        "heading4": ""
    }
    
    for para in doc.paragraphs:
        # Track Heading 2 as chapters (but do not change the chapter number here)
        if para.style.name == 'Heading 2':
            match = re.match(r'(\d+)\.\s*(.*)', para.text.strip())
            if match:
                headings['heading2'] = f"{match.group(1)}: {match.group(2)}"
                print(f"Processing {headings['heading2']}")  # Debugging print
        
        # Track Heading 3 for sub-sections and detect "Bibliography"
        elif para.style.name == 'Heading 3':
            if para.text.strip().lower() == 'bibliography':
                print(f"Detected Bibliography, ending {headings['chapter']}")
                current_chapter_number += 1
                headings['chapter'] = f"Chapter {current_chapter_number}: Next Chapter"
                headings['heading2'] = headings['chapter']  # Update heading2 for the new chapter
                headings['heading3'] = ""
                headings['heading4'] = ""
                print(f"Starting {headings['chapter']}")  # Debugging print
            else:
                headings['heading3'] = para.text.strip()
        
        # Track Heading 4 if present
        elif para.style.name == 'Heading 4':
            headings['heading4'] = para.text.strip()
        
        # Check if the paragraph matches the citation pattern
        match = re.match(pattern, para.text.strip())
        if match:
            author = match.group(1)
            date = match.group(2)
            title = match.group(3)
            citation_info = {
                "chapter": headings['chapter'],
                "heading2": headings['heading2'],
                "heading3": headings['heading3'],
                "heading4": headings['heading4'],
                "author": author,
                "date": date,
                "title": title
            }
            citations.append(citation_info)
            chapter_citation_count[headings['chapter']] += 1
            print(f"Extracted citation: {citation_info}")  # Debugging print
    
    return citations, chapter_citation_count

def process_document(doc):
    all_citations = []
    chapter_citation_count = defaultdict(int)

    # Extract all citations in the document
    citations, chapter_citation_count = extract_citations(doc)
    all_citations.extend(citations)
    
    return all_citations, chapter_citation_count

# Load the .docx file
doc_path = 'cyberutopias-r.docx'
doc = Document(doc_path)

# Verify document loaded
if not doc:
    print("Failed to load the document.")
else:
    print("Document loaded successfully.")

# Process the document
citations_json, chapter_citation_count = process_document(doc)

# Check if any citations were extracted
if not citations_json:
    print("No citations were extracted.")
else:
    # Save the extracted citations as JSON
    output_json_path = 'extracted_citations_with_headings.json'
    with open(output_json_path, 'w') as outfile:
        json.dump(citations_json, outfile, indent=2)

    # Print the total number of citations extracted
    total_citations = len(citations_json)
    print(f"Total number of citations: {total_citations}")
    
    # Print the number of citations per chapter
    print("\nNumber of citations per chapter:")
    for chapter, count in chapter_citation_count.items():
        print(f"{chapter}: {count} citations")
