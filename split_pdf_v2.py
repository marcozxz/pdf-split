import fitz
import os

pdf_path = r"C:\Users\zz_20\WorkBuddy\20260430112416\20260424142335.pdf"
output_dir = r"C:\Users\zz_20\WorkBuddy\20260430112416\split_output"

# Corrected classification (v2):
# - Pages 4-6 are 付款意见书 (Payment Opinion Letter):
#   Page 4: 一、合同支付概况 (payment overview)
#   Page 5: 二、审核情况 (review/audit details)
#   Page 6: 三、审核结果 (review result)
# - Page 1 remains 付款申请单
# - Added 7th category: 付款意见书

page_categories = {
    1: "1_付款申请单",     # 付款申请单 (payment application form)
    2: "3_发票",           # 电子发票
    3: "2_支付申请表",     # 支付申请表
    4: "7_付款意见书",     # 付款意见书 - 一、合同支付概况
    5: "7_付款意见书",     # 付款意见书 - 二、审核情况
    6: "7_付款意见书",     # 付款意见书 - 三、审核结果
    7: "4_支付情况统计表",  # 支付情况统计表
    8: "5_合同文件",       # 合同文件
    9: "5_合同文件",       # 合同协议书
    10: "5_合同文件",      # 合同价款
    11: "5_合同文件",      # 资金监管条款
    12: "5_合同文件",      # 合同签署页
    13: "5_合同文件",      # 招标文件
    14: "6_检测报告",      # 应急预案
    15: "6_检测报告",      # 检测要求
    16: "6_检测报告",      # 浦西变电站预防性试验报告
    17: "6_检测报告",      # 浦东变电站预防性试验报告
    18: "6_检测报告",      # 安全用具试验报告
    19: "6_检测报告",      # 雷电防护装置检验检测报告
    20: "6_检测报告",      # 雷电防护装置检验检测报告（续）
    21: "6_检测报告",      # 变形监测技术报告（年度）
    22: "6_检测报告",      # 变形监测技术报告（年度）
    23: "6_检测报告",      # 变形监测技术报告（第四季度）
    24: "6_检测报告",      # 变形监测技术报告（第四季度）
    25: "6_检测报告",      # 变形监测技术报告（第三季度）
    26: "6_检测报告",      # 变形监测技术报告（第三季度）
    27: "6_检测报告",      # 变形监测技术报告（第一季度）
    28: "6_检测报告",      # 变形监测技术报告（第一季度）
    29: "6_检测报告",      # 渗水量检测表
    30: "6_检测报告",      # 废水PH值检测表
    31: "6_检测报告",      # 路面抗滑能力检测报告
    32: "6_检测报告",      # VI检测表
    33: "6_检测报告",      # CO检测表
    34: "6_检测报告",      # 不间断电源试验记录
    35: "6_检测报告",      # 电梯检验报告
    36: "6_检测报告",      # 不间断电源试验记录
    37: "6_检测报告",      # 照度检测表（7月）
    38: "6_检测报告",      # 照度检测表（1月）
    39: "6_检测报告",      # 消防检测报告
    40: "6_检测报告",      # 消防检测报告（续）
}

# Clean old output files
if os.path.exists(output_dir):
    for f in os.listdir(output_dir):
        if f.endswith('.pdf'):
            os.remove(os.path.join(output_dir, f))
else:
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
    
    page_count = len(pages)
    results.append(f"{category}: {page_count}页 (原始第{sorted(pages)}页)")

# Save classification results
result_path = os.path.join(output_dir, "classification_results.txt")
with open(result_path, "w", encoding="utf-8") as f:
    f.write("PDF拆分分类结果（v2 - 含付款意见书）\n")
    f.write("=" * 60 + "\n\n")
    for r in results:
        f.write(r + "\n")
    f.write(f"\n总计: {sum(len(v) for v in category_pages.values())}页\n")
    f.write("\n详细页码分类：\n")
    f.write("-" * 60 + "\n")
    for page_num in range(1, 41):
        f.write(f"第{page_num:2d}页: {page_categories[page_num]}\n")

doc.close()

with open(os.path.join(output_dir, "done_v2.txt"), "w") as f:
    f.write("Split v2 completed successfully!")
