"""Quick OCR debug - only show classification results for all 40 pages."""
import fitz
from rapidocr_onnxruntime import RapidOCR
import tempfile, os

doc = fitz.open('20260424142335.pdf')
ocr = RapidOCR()
print(f'Total pages: {len(doc)}')

keywords_map = {
    "付款申请单": ["付款申请单", "付款申请"],
    "支付申请表": ["支付申请表", "支付申请"],
    "发票": ["发票", "增值税", "普通发票", "专用发票"],
    "支付情况统计表": ["支付情况统计表", "支付情况统计", "支付统计"],
    "合同文件": ["合同", "协议", "合同文件", "招标文件"],
    "检测报告": ["检测报告", "检验报告", "试验报告", "检测", "试验", "监测", "检验", "预防性试验"],
    "付款意见书": ["付款意见书", "合同付款意见", "支付概况", "审核情况", "审核结果"],
}

for pn in range(len(doc)):
    page = doc[pn]
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat)
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
        tmp_path = tmp.name
    pix.save(tmp_path)

    result, _ = ocr(tmp_path)
    os.unlink(tmp_path)

    if result:
        texts = [item[1] for item in result]
        joined = ' '.join(texts)
    else:
        texts = []
        joined = ''

    # Find ALL matching categories
    all_matched = []
    for cat, kws in keywords_map.items():
        for kw in kws:
            if kw in joined:
                all_matched.append(f"{cat}[{kw}]")
                break

    # Current tool logic: sorted by max keyword length, first match wins
    sorted_cats = sorted(keywords_map.items(), key=lambda c: max(len(kw) for kw in c[1]), reverse=True)
    category = None
    for cat, kws in sorted_cats:
        for kw in kws:
            if kw in joined:
                category = cat
                break
        if category:
            break
    if category is None:
        category = "未分类"

    print(f'Page {pn+1:2d}: {category:10s} | all_matches={all_matched} | text[:80]={joined[:80]}')

doc.close()
