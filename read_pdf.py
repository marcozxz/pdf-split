import sys
from pypdf import PdfReader

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\pdf_info.txt"

reader = PdfReader(pdf_path)

with open(output_path, "w", encoding="utf-8") as f:
    f.write(f"Total pages: {len(reader.pages)}\n")
    f.write(f"Metadata: {reader.metadata}\n")
    
    outline = reader.outline
    if outline:
        f.write("Bookmarks found:\n")
        for item in outline:
            f.write(f"  - {item}\n")
    else:
        f.write("No bookmarks found.\n")
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if text:
            first_lines = text[:300].strip().replace('\n', ' | ')
        else:
            first_lines = "(empty page)"
        f.write(f"\nPage {i+1}: {first_lines}\n")

print("Done!")
