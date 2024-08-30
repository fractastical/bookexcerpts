import re
import json
from docx import Document
from collections import defaultdict

def extract_indented_quotes(doc, headings):
    # Pattern for matching "author (date) title" format
    pattern = r'^\s*(.*?)\s*\((\d{4})\)\s*(.*?)\.?\s*$'
    quotes = []
    paragraphs = iter(doc.paragraphs)
    current_quote = []  # List to accumulate multi-paragraph quotes
    
    for para in paragraphs:
        # Check if the current paragraph is likely part of a quote (indented or not a heading)
        if para.paragraph_format.left_indent or not para.style.name.startswith('Heading'):
            current_quote.append(para.text.strip())
            
            # Look ahead to see if the next paragraph is an attribution
            next_para = None
            try:
                next_para = next(paragraphs)
            except StopIteration:
                pass  # End of document reached
            
            if next_para and re.match(pattern, next_para.text.strip()):
                # Combine all collected paragraphs into one quote
                full_quote = " ".join(current_quote)
                match = re.match(pattern, next_para.text.strip())
                author = match.group(1)
                date = match.group(2)
                title = match.group(3)
                quote_info = {
                    "chapter": headings.get('Heading 2', 'Unknown Chapter'),
                    "heading2": headings.get('Heading 2', ''),
                    "heading3": headings.get('Heading 3', ''),
                    "heading4": headings.get('Heading 4', ''),
                    "author": author,
                    "date": date,
                    "title": title,
                    "quote": full_quote
                }
                quotes.append(quote_info)
                print(f"Extracted multi-paragraph quote: {quote_info}")  # Debugging print
                current_quote = []  # Reset after processing
            elif next_para:
                # If the next paragraph isn't an attribution, put it back into the iteration
                current_quote.append(next_para.text.strip())
            else:
                # If there's no next paragraph, treat it as an orphaned attribution
                if current_quote:
                    quote_info = {
                        "chapter": headings.get('Heading 2', 'Unknown Chapter'),
                        "heading2": headings.get('Heading 2', ''),
                        "heading3": headings.get('Heading 3', ''),
                        "heading4": headings.get('Heading 4', ''),
                        "quote": " ".join(current_quote),
                        "author": "",
                        "date": "",
                        "title": "Orphaned quote without attribution"
                    }
                    quotes.append(quote_info)
                    print(f"Orphaned quote: {quote_info}")  # Debugging print
                current_quote = []  # Reset after processing
        else:
            # If it's a heading or other paragraph, check for an attribution on its own
            if current_quote:
                quote_info = {
                    "chapter": headings.get('Heading 2', 'Unknown Chapter'),
                    "heading2": headings.get('Heading 2', ''),
                    "heading3": headings.get('Heading 3', ''),
                    "heading4": headings.get('Heading 4', ''),
                    "quote": " ".join(current_quote),
                    "author": "",
                    "date": "",
                    "title": "Orphaned quote without attribution"
                }
                quotes.append(quote_info)
                print(f"Orphaned quote: {quote_info}")  # Debugging print
                current_quote = []  # Reset after processing
            
            # Check if this paragraph is an orphaned attribution
            match = re.match(pattern, para.text.strip())
            if match:
                author = match.group(1)
                date = match.group(2)
                title = match.group(3)
                quote_info = {
                    "chapter": headings.get('Heading 2', 'Unknown Chapter'),
                    "heading2": headings.get('Heading 2', ''),
                    "heading3": headings.get('Heading 3', ''),
                    "heading4": headings.get('Heading 4', ''),
                    "quote": "Orphaned attribution without preceding quote",
                    "author": author,
                    "date": date,
                    "title": title
                }
                quotes.append(quote_info)
                print(f"Extracted orphaned attribution: {quote_info}")  # Debugging print

    return quotes

def process_document(doc):
    all_quotes = []
    chapter_quotes_count = defaultdict(int)
    
    headings = {}

    for para in doc.paragraphs:
        # Track Heading 2, Heading 3, Heading 4, etc.
        if para.style.name == 'Heading 2':
            match = re.match(r'(\d+)\.\s*(.*)', para.text.strip())
            if match:
                headings['Heading 2'] = f"{match.group(1)}: {match.group(2)}"
                print(f"Processing {headings['Heading 2']}")  # Debugging print
        elif para.style.name == 'Heading 3':
            headings['Heading 3'] = para.text.strip()
        elif para.style.name == 'Heading 4':
            headings['Heading 4'] = para.text.strip()
        
        # Extract quotes within the current headings context
        quotes = extract_indented_quotes(doc, headings)
        all_quotes.extend(quotes)
        chapter_quotes_count[headings.get('Heading 2', 'Unknown Chapter')] += len(quotes)
    
    return all_quotes, chapter_quotes_count

# Load the .docx file
doc_path = 'cyberutopias-r.docx'
doc = Document(doc_path)

# Verify document loaded
if not doc:
    print("Failed to load the document.")
else:
    print("Document loaded successfully.")

# Process the document
quotes_json, chapter_quotes_count = process_document(doc)

# Check if any quotes were extracted
if not quotes_json:
    print("No quotes were extracted.")
else:
    # Save the extracted quotes as JSON
    output_json_path = 'extracted_quotes.json'
    with open(output_json_path, 'w') as outfile:
        json.dump(quotes_json, outfile, indent=2)

    # Print the number of excerpts per chapter and the total number
    total_excerpts = sum(chapter_quotes_count.values())
    print(f"Total number of excerpts: {total_excerpts}")
    for chapter, count in chapter_quotes_count.items():
        print(f"{chapter}: {count} excerpts")
