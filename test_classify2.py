"""Test classification v2 with updated keywords."""
import fitz
from rapidocr_onnxruntime import RapidOCR
import tempfile, os

doc = fitz.open('20260424142335.pdf')
ocr = RapidOCR()

categories = [
    {
        "name": "付款申请单",
        "keywords": ["付款申请单", "付款申请"],
        "combo_keywords": ["付款帐号", "请款金额", "收款单位", "农民工工资", "累计付款额", "按发票管理分类", "请款总", "业务摘要"],
    },
    {
        "name": "支付申请表",
        "keywords": ["支付申请表", "支付申请", "费用支付申请表"],
        "combo_keywords": [],
    },
    {
        "name": "发票",
        "keywords": ["发票", "增值税", "普通发票", "专用发票", "电子发票"],
        "combo_keywords": [],
    },
    {
        "name": "支付情况统计表",
        "keywords": ["支付情况统计表", "合同款支付情况统计表", "支付情况统计"],
        "combo_keywords": [],
    },
    {
        "name": "合同文件",
        "keywords": ["合同文件", "合同协议书", "招标文件", "合同专用章"],
        "combo_keywords": [],
    },
    {
        "name": "检测报告",
        "keywords": ["检测报告", "检验报告", "试验报告", "预防性试验报告", "检验检测报告", "监测技术报告"],
        "combo_keywords": [],
    },
    {
        "name": "付款意见书",
        "keywords": ["付款意见书", "合同付款意见书", "审核结果", "审核情况"],
        "combo_keywords": [],
    },
]


def classify_page(text, cats):
    if not text:
        return None, 0, []
    scores = {}
    details = {}
    for cat in cats:
        score = 0
        cat_details = []
        for kw in cat["keywords"]:
            if kw in text:
                base_score = 10
                length_bonus = max(0, (len(kw) - 2) * 2)
                total = base_score + length_bonus
                score += total
                cat_details.append(f"主[{kw}] +{total}")
        if cat.get("combo_keywords"):
            combo_matches = 0
            matched_combos = []
            for ckw in cat["combo_keywords"]:
                if ckw in text:
                    combo_matches += 1
                    matched_combos.append(ckw)
            if combo_matches >= 2:
                combo_score = combo_matches * 3
                score += combo_score
                cat_details.append(f"组合{matched_combos} +{combo_score}")
        scores[cat["name"]] = score
        details[cat["name"]] = cat_details
    best_cat = None
    best_score = 0
    for cat_name, score in scores.items():
        if score > best_score:
            best_score = score
            best_cat = cat_name
    return best_cat, best_score, details


# Process all pages
page_results = {}
page_scores = {}
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
        text = " ".join([item[1] for item in result])
    else:
        text = ""
    cat, score, details = classify_page(text, categories)
    if cat is None or score == 0:
        cat = "未分类"
    page_results[pn + 1] = cat
    page_scores[pn + 1] = score
    all_cats = {name: (scores, d) for name, (scores, d) in ((n, (s, details.get(n, []))) for n, s in ((n, sum(1 for k in categories if n == k["name"] and any(kw in text for kw in k["keywords"])) or 0) for n in set(c["name"] for c in categories))) if scores > 0} if False else {}
    print(f'Page {pn+1:2d}: {cat:10s} (score={score:3d})')

# Apply context optimization
print('\n--- After context optimization ---')
pages = sorted(page_results.keys())

# Rule 1: Island fix
changed = True
while changed:
    changed = False
    for i in range(len(pages)):
        pn = pages[i]
        current_cat = page_results[pn]
        if i > 0 and i < len(pages) - 1:
            prev_cat = page_results[pages[i - 1]]
            next_cat = page_results[pages[i + 1]]
            if prev_cat == next_cat and prev_cat != current_cat and prev_cat != "未分类":
                page_results[pn] = prev_cat
                print(f'  Island fix: Page {pn} {current_cat} -> {prev_cat}')
                changed = True
                break

# Rule 2: Multi-page gap fix
for candidate in ["付款意见书", "合同文件", "检测报告", "支付申请表", "支付情况统计表", "付款申请单", "发票"]:
    candidate_pages = [pn for pn in pages if page_results[pn] == candidate]
    if not candidate_pages:
        continue
    min_page = min(candidate_pages)
    max_page = max(candidate_pages)
    for pn in range(min_page, max_page + 1):
        if page_results[pn] != candidate:
            prev_same = pn > min_page
            next_same = pn < max_page
            if prev_same and next_same:
                old_cat = page_results[pn]
                page_results[pn] = candidate
                print(f'  Gap fix: Page {pn} {old_cat} -> {candidate}')

print('\n--- Final results ---')
for pn in range(1, len(doc) + 1):
    print(f'Page {pn:2d}: {page_results[pn]}')

# Expected (from original manual split):
# Page 1: 付款申请单
# Page 2: 发票
# Page 3: 支付申请表
# Pages 4-6: 付款意见书
# Page 7: 支付情况统计表
# Pages 8-13: 合同文件
# Pages 14-40: 检测报告

doc.close()
