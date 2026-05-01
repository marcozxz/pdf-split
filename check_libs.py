import sys
results = []
try:
    import pdf2image
    results.append("pdf2image: OK")
except ImportError as e:
    results.append(f"pdf2image: MISSING - {e}")

try:
    import pytesseract
    results.append("pytesseract: OK")
except ImportError as e:
    results.append(f"pytesseract: MISSING - {e}")

try:
    from PIL import Image
    results.append("PIL: OK")
except ImportError as e:
    results.append(f"PIL: MISSING - {e}")

try:
    from pypdf import PdfReader
    results.append("pypdf: OK")
except ImportError as e:
    results.append(f"pypdf: MISSING - {e}")

with open(r"C:\Users\zz_20\WorkBuddy\20260430112416\lib_check_result.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(results))
