"""Deep OCR analysis for Page 1 and Page 5."""
import fitz
from rapidocr_onnxruntime import RapidOCR
import tempfile, os

doc = fitz.open('20260424142335.pdf')
ocr = RapidOCR()

for target_pn in [0, 2, 4]:  # Page 1, 3, 5
    page = doc[target_pn]
    mat = fitz.Matrix(3, 3)  # Higher zoom for better OCR
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

    print(f'\n{"="*80}')
    print(f'=== Page {target_pn+1} (zoom=3x) ===')
    print(f'All {len(texts)} segments:')
    for i, t in enumerate(texts):
        print(f'  [{i:3d}] {t}')
    print(f'\nFull joined text:')
    print(joined[:2000])

doc.close()
