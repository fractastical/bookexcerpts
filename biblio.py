import re
import json
from docx import Document
from collections import defaultdict

def extract_citations(doc):
    main_text_pattern = r'^\s*(.*?)\s*\((\d{4}(?:-\d{4})?)\)\.?\s*(.*?)(?:,\s*p\.\s*(\d+))?\.?\s*$'
    footnote_patterns = [
        r'(\w+),\s+(.*?)\s*\((\d{4}(?:-\d{4})?)\)\.\s*(.*?)\.',  # Author, Title (Year). Publisher.
        r'(\w+)\s+(\w+)\s+(?:&|and)\s+(\w+)\s+(\w+)\s*\((\d{4}(?:-\d{4})?)\)\.\s*(.*?)\.',  # Author1 Lastname1 & Author2 Lastname2 (Year). Title.
        r'(\w+),\s+(.*?)\s*(\d{4}(?:-\d{4})?)\.',  # Author, Title Year.
        r'(\w+),\s+(\w+)\s*\((\d{4}(?:-\d{4})?)\)\s*(.*?)\.',  # Lastname, Firstname (Year) Title.
    ]
    bibliography_pattern = r'(.*?)\s*\((\d{4}(?:-\d{4})?)\)\.?\s*(.*?)\.\s*'
    
    citations = defaultdict(lambda: {'main_text': [], 'footnotes': [], 'extracts': [], 'references': []})
    current_chapter_number = 0
    headings = {
        "chapter": "",
        "heading2": "",
        "heading3": "",
        "heading4": ""
    }

    in_bibliography = False
    bibliography_section = ""
    bibliography_headings = {"extracts", "references", "additional readings"}

    def process_footnote(footnote_text):
        standard_format = False
        for pattern in footnote_patterns:
            match = re.search(pattern, footnote_text)
            if match:
                groups = match.groups()
                if len(groups) == 4:  # First pattern
                    author, title, date, publisher = groups
                elif len(groups) == 6:  # Second pattern
                    author = f"{groups[0]} {groups[1]} and {groups[2]} {groups[3]}"
                    date, title = groups[4], groups[5]
                elif len(groups) == 3:  # Third pattern
                    author, title, date = groups
                elif len(groups) == 4:  # Fourth pattern
                    author = f"{groups[1]} {groups[0]}"
                    date, title = groups[2], groups[3]
                
                citation_info = {
                    "author": author,
                    "date": date,
                    "title": title,
                    "standard_format": True
                }
                standard_format = True
                break
        
        if not standard_format:
            citation_info = {
                "full_text": footnote_text,
                "standard_format": False
            }
        
        return citation_info

    # Process footnotes
    footnotes = doc.part._element.findall('.//w:footnote', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    for footnote in footnotes:
        if footnote.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id'] not in ['0', '-1']:  # Skip separator and continuation separator footnotes
            footnote_text = ''.join(paragraph.text for paragraph in footnote.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            citation_info = process_footnote(footnote_text)
            citations[headings['chapter']]['footnotes'].append(citation_info)
            print(f"Extracted footnote: {citation_info}")

    for para in doc.paragraphs:
        if para.style.name == 'Heading 2':
            match = re.match(r'(\d+)\.\s*(.*)', para.text.strip())
            if match:
                current_chapter_number = int(match.group(1))
                chapter_title = match.group(2)
            else:
                current_chapter_number += 1
                chapter_title = para.text.strip()
            
            headings['chapter'] = f"Chapter {current_chapter_number}: {chapter_title}"
            headings['heading2'] = headings['chapter']
            headings['heading3'] = ""
            headings['heading4'] = ""
            in_bibliography = False
            print(f"Processing {headings['chapter']}")

        elif para.style.name in ['Heading 3', 'Heading 4']:
            lower_text = para.text.strip().lower()
            if lower_text in bibliography_headings:
                print(f"Detected {para.text.strip()} in {headings['chapter']}")
                in_bibliography = True
                bibliography_section = para.text.strip().lower()
            else:
                headings[para.style.name.lower()] = para.text.strip()
                if not in_bibliography:
                    in_bibliography = False

        if in_bibliography and para.text.strip() and para.style.name not in ['Heading 3', 'Heading 4']:
            bib_match = re.search(bibliography_pattern, para.text.strip())
            if bib_match:
                author = bib_match.group(1).strip()
                date = bib_match.group(2).strip()
                title = bib_match.group(3).strip()
                entry = {
                    "author": author,
                    "date": date,
                    "title": title
                }
                if bibliography_section == 'extracts':
                    citations[headings['chapter']]['extracts'].append(entry)
                elif bibliography_section == 'references':
                    citations[headings['chapter']]['references'].append(entry)
        elif not in_bibliography:
            match = re.match(main_text_pattern, para.text.strip())
            if match:
                author = match.group(1).strip()
                date = match.group(2).strip()
                title = match.group(3).strip()
                page_number = match.group(4)

                citation_info = {
                    "author": author,
                    "date": date,
                    "title": title,
                    "page_number": page_number if page_number else None
                }

                citations[headings['chapter']]['main_text'].append(citation_info)
                print(f"Extracted main text citation: {citation_info}")

    return citations

# The rest of the script (process_document function and main execution) remains the same

def process_document(doc):
    citations = extract_citations(doc)
    output_json = {}

    for chapter, chapter_data in citations.items():
        chapter_json = {
            "citations": [],
            "footnotes": [],
            "extracts": chapter_data['extracts'],
            "references": chapter_data['references']
        }

        missing_extracts = []
        missing_references = []

        # Process main text citations
        for citation in chapter_data['main_text']:
            if citation not in chapter_data['extracts']:
                citation['not_in_extracts'] = True
                missing_extracts.append(citation)
            chapter_json['citations'].append(citation)

        # Process footnotes
        for footnote in chapter_data['footnotes']:
            if footnote.get('standard_format', False):
                if footnote not in chapter_data['references']:
                    footnote['not_in_references'] = True
                    missing_references.append(footnote)
            chapter_json['footnotes'].append(footnote)

        # Generate summary
        summary = {
            "citations_in_text": len(chapter_data['main_text']),
            "citations_in_footnotes": len(chapter_data['footnotes']),
            "listings_in_extracts": len(chapter_data['extracts']),
            "missing_listings_in_extracts": len(missing_extracts),
            "listings_in_references": len(chapter_data['references']),
            "missing_listings_in_references": len(missing_references)
        }

        chapter_json['summary'] = summary
        output_json[chapter] = chapter_json

    # Save all chapters to a single JSON file
    with open('all_chapters_citations.json', 'w') as outfile:
        json.dump(output_json, outfile, indent=2)

    print("Saved all_chapters_citations.json")

    # Print summary for each chapter to console
    print("\nSummary for each chapter:")
    for chapter, chapter_data in output_json.items():
        print(f"\n{chapter}")
        summary = chapter_data['summary']
        print(f"  Citations in text: {summary['citations_in_text']}")
        print(f"  Citations in footnotes: {summary['citations_in_footnotes']}")
        print(f"  Listings in extracts: {summary['listings_in_extracts']}")
        print(f"  Missing listings in extracts: {summary['missing_listings_in_extracts']}")
        print(f"  Listings in references: {summary['listings_in_references']}")
        print(f"  Missing listings in references: {summary['missing_listings_in_references']}")

    return citations

# Load the .docx file
doc_path = 'cyberutopias-r.docx'
doc = Document(doc_path)

if not doc:
    print("Failed to load the document.")
else:
    print("Document loaded successfully.")

process_document(doc)
