import re
from docx2python import docx2python
from collections import defaultdict

def flatten_list(nested_list):
    flat_list = []
    for item in nested_list:
        if isinstance(item, list):
            flat_list.extend(flatten_list(item))
        else:
            flat_list.append(item)
    return flat_list

def extract_footnotes_and_references(doc_path):
    footnotes = []
    references = []
    in_references = False

    with docx2python(doc_path) as doc:
        print("Debugging Document Structure:")
        print("=============================")
        
        print("\nFootnote Structure:")
        for i, footnote_section in enumerate(doc.footnotes):
            print(f"Footnote Section {i}:")
            for j, footnote in enumerate(footnote_section):
                flat_footnote = flatten_list(footnote)
                text = ' '.join(str(item) for item in flat_footnote).strip()
                print(f"  Footnote {j}: {text[:100]}...")  # Print first 100 characters
                if text:
                    footnotes.append(text)

        print("\nMain Document Structure:")
        for i, section in enumerate(doc.body):
            print(f"Section {i}:")
            for j, paragraph in enumerate(section):
                flat_paragraph = flatten_list(paragraph)
                text = ' '.join(str(item) for item in flat_paragraph).strip()
                print(f"  Paragraph {j}: {text[:100]}...")  # Print first 100 characters
                if text.lower().startswith('references'):
                    print("    Found References Section")
                    in_references = True
                    continue
                if in_references and text:
                    references.append(text)

    print(f"\nExtracted {len(footnotes)} footnotes and {len(references)} references.")
    return footnotes, references

def parse_citation(text):
    # This pattern attempts to match various citation formats
    patterns = [
        r'(\w+),\s+(.*?)\s*\((\d{4}(?:-\d{4})?)\)',  # Author, Title (Year)
        r'(\w+)\s+\((\d{4}(?:-\d{4})?)\)',  # Author (Year)
        r'(\w+)\s+(\w+)\s+(?:&|and)\s+(\w+)\s+(\w+)\s*\((\d{4}(?:-\d{4})?)\)'  # Author1 Lastname1 & Author2 Lastname2 (Year)
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(0)
    return text

def compare_footnotes_and_references(footnotes, references):
    parsed_footnotes = [parse_citation(f) for f in footnotes]
    parsed_references = [parse_citation(r) for r in references]

    footnotes_not_in_references = [f for f in parsed_footnotes if f not in parsed_references]
    references_not_in_footnotes = [r for r in parsed_references if r not in parsed_footnotes]

    return footnotes_not_in_references, references_not_in_footnotes

def main(doc_path):
    print(f"Analyzing document: {doc_path}")
    footnotes, references = extract_footnotes_and_references(doc_path)

    footnotes_not_in_references, references_not_in_footnotes = compare_footnotes_and_references(footnotes, references)

    print("\nFootnotes not found in references:")
    for f in footnotes_not_in_references[:10]:  # Print first 10 for brevity
        print(f"  - {f[:100]}...")

    print("\nReferences not found in footnotes:")
    for r in references_not_in_footnotes[:10]:  # Print first 10 for brevity
        print(f"  - {r[:100]}...")

    print("\nSummary:")
    print(f"Total footnotes: {len(footnotes)}")
    print(f"Total references: {len(references)}")
    print(f"Footnotes not in references: {len(footnotes_not_in_references)}")
    print(f"References not in footnotes: {len(references_not_in_footnotes)}")

if __name__ == "__main__":
    doc_path = 'cyberutopias-r.docx'
    main(doc_path)
