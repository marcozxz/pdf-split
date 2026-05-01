"""Debug OCR results for each page of the PDF."""
import fitz
from rapidocr_onnxruntime import RapidOCR
import tempfile
import os

doc = fitz.open('20260424142335.pdf')
ocr = RapidOCR()
print(f'Total pages: {len(doc)}')

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

    print(f'\n=== Page {pn+1} ===')
    print(f'Segments ({len(texts)}): {texts[:25]}')
    print(f'Full text (first 500): {joined[:500]}')

    # Check which keywords match
    keywords_map = {
        "付款申请单": ["付款申请单", "付款申请"],
        "支付申请表": ["支付申请表", "支付申请"],
        "发票": ["发票", "增值税", "普通发票", "专用发票"],
        "支付情况统计表": ["支付情况统计表", "支付情况统计", "支付统计"],
        "合同文件": ["合同", "协议", "合同文件", "招标文件"],
        "检测报告": ["检测报告", "检验报告", "试验报告", "检测", "试验", "监测", "检验", "预防性试验"],
        "付款意见书": ["付款意见书", "合同付款意见", "支付概况", "审核情况", "审核结果"],
    }

    matched = []
    for cat, kws in keywords_map.items():
        for kw in kws:
            if kw in joined:
                matched.append(f"{cat}(kw={kw})")
                break
    if matched:
        print(f'Matched: {matched}')
    else:
        print('Matched: NONE')

doc.close()
