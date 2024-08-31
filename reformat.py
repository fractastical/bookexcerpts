from docx import Document
import re

# Load the .docx file
doc = Document('capitalized_citations.docx')

# Function to capitalize the first letter after a year and period
def capitalize_after_year(paragraph):
    # Regular expression to find the pattern like "2023. ‘m" or "2021. the"
    pattern = re.compile(r'(\d{4}\.\s*[“‘"]?\s*)([a-z])')
    # Capitalize the letter after the year and period
    new_text = pattern.sub(lambda match: match.group(1) + match.group(2).upper(), paragraph.text)
    paragraph.text = new_text

# Initialize counters and regex pattern for chapters
chapter_number = 1
chapter_pattern = re.compile(r'^Chapter \d+:', re.IGNORECASE)

# Iterate over the paragraphs in the document
for paragraph in doc.paragraphs:
    # Check if the paragraph starts with a chapter heading
    if chapter_pattern.match(paragraph.text.strip()):
        # Replace the chapter number with the incremented one
        new_chapter_heading = f'Chapter {chapter_number}: ' + paragraph.text.split(":")[1].strip()
        paragraph.text = new_chapter_heading
        chapter_number += 1
    else:
        # Capitalize the first letter after the year and period
        capitalize_after_year(paragraph)

# Save the modified document
doc.save('capitalized_citations-2.docx')
