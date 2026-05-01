import sys
import os
from pdf2image import convert_from_path
import pytesseract
from pypdf import PdfReader, PdfWriter

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_dir = r"C:\Users\zz_20\WorkBuddy\20260430112416\split_output"
log_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\ocr_log.txt"

# Create output directory
os.makedirs(output_dir, exist_ok=True)

# Define document categories and their keywords
categories = {
    "1_付款申请单": ["付款申请单", "付款申请"],
    "2_支付申请表": ["支付申请表", "支付申请"],
    "3_发票": ["发票", "增值税", "普通发票", "专用发票"],
    "4_支付情况统计表": ["支付情况统计表", "支付情况统计", "支付统计"],
    "5_合同文件": ["合同", "协议", "合同文件"],
    "6_检测报告": ["检测报告", "检验报告", "试验报告", "检测"],
}

# Page classification result: {category_name: [page_numbers]}
page_classification = {}
unclassified = []

# Step 1: Convert PDF pages to images and OCR
log_lines = []
log_lines.append("Starting OCR processing...")
log_lines.append(f"PDF: {pdf_path}")

reader = PdfReader(pdf_path)
total_pages = len(reader.pages)
log_lines.append(f"Total pages: {total_pages}")

# Convert all pages to images
log_lines.append("Converting PDF to images...")
images = convert_from_path(pdf_path, dpi=200)
log_lines.append(f"Converted {len(images)} pages.")

# OCR each page and classify
for i, img in enumerate(images):
    page_num = i + 1
    log_lines.append(f"\n--- Processing Page {page_num} ---")
    
    # OCR with Chinese + English
    try:
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')
    except Exception as e:
        log_lines.append(f"OCR error on page {page_num}: {e}")
        try:
            text = pytesseract.image_to_string(img, lang='chi_sim')
        except Exception as e2:
            log_lines.append(f"OCR fallback error on page {page_num}: {e2}")
            text = ""
    
    # Show first 300 chars of OCR result
    preview = text.strip()[:300].replace('\n', ' | ')
    log_lines.append(f"OCR text: {preview}")
    
    # Classify the page
    classified = False
    for cat_name, keywords in categories.items():
        for kw in keywords:
            if kw in text:
                if cat_name not in page_classification:
                    page_classification[cat_name] = []
                page_classification[cat_name].append(page_num)
                log_lines.append(f"Classified as: {cat_name} (keyword: '{kw}')")
                classified = True
                break
        if classified:
            break
    
    if not classified:
        unclassified.append(page_num)
        log_lines.append("UNCLASSIFIED")

# Write log
with open(log_path, "w", encoding="utf-8") as f:
    f.write("\n".join(log_lines))

# Write classification summary
summary_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\classification_summary.txt"
with open(summary_path, "w", encoding="utf-8") as f:
    f.write("=== PDF Page Classification Summary ===\n\n")
    for cat_name, pages in page_classification.items():
        f.write(f"{cat_name}: Pages {pages}\n")
    if unclassified:
        f.write(f"\nUnclassified pages: {unclassified}\n")

print("OCR and classification complete!")
