import fitz  # PyMuPDF
import os

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_dir = r"C:\Users\zz_20\WorkBuddy\20260430112416\page_images"
os.makedirs(output_dir, exist_ok=True)

doc = fitz.open(pdf_path)
total_pages = len(doc)

results = []
for i in range(total_pages):
    page = doc[i]
    
    # Try text extraction
    text = page.get_text().strip()
    text_preview = text[:200].replace('\n', ' | ') if text else "(no text)"
    
    # Render page to image (200 DPI equivalent)
    mat = fitz.Matrix(2, 2)  # 2x zoom = ~144 DPI
    pix = page.get_pixmap(matrix=mat)
    img_path = os.path.join(output_dir, f"page_{i+1:02d}.png")
    pix.save(img_path)
    
    results.append(f"Page {i+1}: {text_preview}")

# Save results
with open(r"C:\Users\zz_20\WorkBuddy\20260430112416\pymupdf_text.txt", "w", encoding="utf-8") as f:
    f.write(f"Total pages: {total_pages}\n\n")
    for r in results:
        f.write(r + "\n")

doc.close()
