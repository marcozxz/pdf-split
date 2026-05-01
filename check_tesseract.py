import pytesseract
try:
    langs = pytesseract.get_languages()
    with open(r"C:\Users\zz_20\WorkBuddy\20260430112416\tess_langs.txt", "w", encoding="utf-8") as f:
        f.write("Available Tesseract languages:\n")
        for lang in langs:
            f.write(f"  {lang}\n")
except Exception as e:
    with open(r"C:\Users\zz_20\WorkBuddy\20260430112416\tess_langs.txt", "w", encoding="utf-8") as f:
        f.write(f"Error: {e}")
