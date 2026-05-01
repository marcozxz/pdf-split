import fitz
import os
import json

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_dir = r"C:\Users\zz_20\WorkBuddy\20260430112416\split_output"

# Manually corrected classification based on OCR results review
# Key corrections:
# - Page 1: 3_发票 -> 1_付款申请单 (it's a payment application with bank info, "按发票管理分类" caused false match)
# - Page 5: 3_发票 -> 5_合同文件 (it's contract content about maintenance assessment)
# - Page 6: 5_合同文件 -> 1_付款申请单 (it's the payment review result, part of payment application)
# - Page 13: 未分类 -> 5_合同文件 (it's a bidding document, part of contract)
# - Pages 21-28: 未分类 -> 6_检测报告 (deformation monitoring reports are inspection reports)
# - Page 39: 5_合同文件 -> 6_检测报告 (fire inspection report, not contract)

page_categories = {
    1: "1_付款申请单",    # Payment application with bank details
    2: "3_发票",          # Electronic invoice
    3: "2_支付申请表",    # Payment application table
    4: "1_付款申请单",    # Contract payment opinion letter (part of payment application)
    5: "5_合同文件",      # Contract terms - maintenance assessment methods
    6: "1_付款申请单",    # Payment review result (part of payment application)
    7: "4_支付情况统计表", # Payment statistics table
    8: "5_合同文件",      # Contract document
    9: "5_合同文件",      # Contract agreement
    10: "5_合同文件",     # Contract pricing
    11: "5_合同文件",     # Contract fund supervision
    12: "5_合同文件",     # Contract signing page
    13: "5_合同文件",     # Bidding document (part of contract)
    14: "6_检测报告",     # Emergency plans (part of inspection requirements)
    15: "6_检测报告",     # Inspection requirements
    16: "6_检测报告",     # Substation test report
    17: "6_检测报告",     # Substation test report
    18: "6_检测报告",     # Safety equipment test report
    19: "6_检测报告",     # Lightning protection inspection
    20: "6_检测报告",     # Lightning protection inspection
    21: "6_检测报告",     # Deformation monitoring report (annual)
    22: "6_检测报告",     # Deformation monitoring report (annual)
    23: "6_检测报告",     # Deformation monitoring report (Q4)
    24: "6_检测报告",     # Deformation monitoring report (Q4)
    25: "6_检测报告",     # Deformation monitoring report (Q3)
    26: "6_检测报告",     # Deformation monitoring report (Q3)
    27: "6_检测报告",     # Deformation monitoring report (Q1)
    28: "6_检测报告",     # Deformation monitoring report (Q1)
    29: "6_检测报告",     # Seepage detection table
    30: "6_检测报告",     # Wastewater PH test table
    31: "6_检测报告",     # Road skid resistance test report
    32: "6_检测报告",     # VI detection table
    33: "6_检测报告",     # CO detection table
    34: "6_检测报告",     # UPS test record
    35: "6_检测报告",     # Elevator inspection report
    36: "6_检测报告",     # UPS test record
    37: "6_检测报告",     # Illumination detection table (July)
    38: "6_检测报告",     # Illumination detection table (January)
    39: "6_检测报告",     # Fire protection inspection report
    40: "6_检测报告",     # Fire protection inspection report
}

# Create output directory
os.makedirs(output_dir, exist_ok=True)

# Open source PDF
doc = fitz.open(pdf_path)

# Group pages by category
category_pages = {}
for page_num, category in page_categories.items():
    if category not in category_pages:
        category_pages[category] = []
    category_pages[category].append(page_num)

# Create a PDF for each category
results = []
for category, pages in sorted(category_pages.items()):
    new_doc = fitz.open()
    for page_num in sorted(pages):
        new_doc.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)
    
    output_path = os.path.join(output_dir, f"{category}.pdf")
    new_doc.save(output_path)
    new_doc.close()
    
    results.append(f"{category}: pages {sorted(pages)} -> {output_path}")

# Save classification results
result_path = os.path.join(output_dir, "classification_results.txt")
with open(result_path, "w", encoding="utf-8") as f:
    f.write("PDF拆分分类结果\n")
    f.write("=" * 60 + "\n\n")
    for r in results:
        f.write(r + "\n")
    f.write("\n\n详细页码分类：\n")
    f.write("-" * 60 + "\n")
    for page_num in range(1, 41):
        f.write(f"第{page_num:2d}页: {page_categories[page_num]}\n")

doc.close()

# Write a simple done marker
with open(os.path.join(output_dir, "done.txt"), "w") as f:
    f.write("Split completed successfully!")
