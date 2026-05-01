import fitz
import os
import json
from rapidocr_onnxruntime import RapidOCR

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_dir = r"C:\Users\zz_20\WorkBuddy\20260430112416\page_images"

# Initialize OCR engine
ocr = RapidOCR()

# Categories and keywords
categories = {
    "1_付款申请单": ["付款申请单", "付款申请"],
    "2_支付申请表": ["支付申请表", "支付申请"],
    "3_发票": ["发票", "增值税", "普通发票", "专用发票"],
    "4_支付情况统计表": ["支付情况统计表", "支付情况统计", "支付统计"],
    "5_合同文件": ["合同", "协议", "合同文件"],
    "6_检测报告": ["检测报告", "检验报告", "试验报告", "检测"],
}

def classify_text(text):
    """Classify text into one of the categories."""
    for cat_name, keywords in categories.items():
        for kw in keywords:
            if kw in text:
                return cat_name
    return None

# Open PDF and process each page
doc = fitz.open(pdf_path)
total_pages = len(doc)
results = {}

for page_num in range(total_pages):
    page = doc[page_num]
    # Render page to image at 2x zoom for better OCR
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat)
    
    # Save as PNG temporarily
    img_path = os.path.join(output_dir, f"page_{page_num+1:02d}.png")
    pix.save(img_path)
    
    # OCR the image
    try:
        ocr_result, elapse = ocr(img_path)
        # Extract text from OCR result
        if ocr_result:
            text_lines = [item[1] for item in ocr_result]
            full_text = " ".join(text_lines)
        else:
            full_text = ""
    except Exception as e:
        full_text = f"OCR_ERROR: {e}"
    
    # Classify
    category = classify_text(full_text)
    
    results[str(page_num + 1)] = {
        "page": page_num + 1,
        "category": category if category else "未分类",
        "text_preview": full_text[:200] if full_text else "(empty)"
    }

# Save results
result_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\ocr_results.json"
with open(result_path, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

# Also save a simple summary
summary_lines = []
for page_num, info in results.items():
    summary_lines.append(f"Page {info['page']}: {info['category']} | {info['text_preview'][:80]}")

summary_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\ocr_summary.txt"
with open(summary_path, "w", encoding="utf-8") as f:
    f.write("\n".join(summary_lines))

print("Done!")
