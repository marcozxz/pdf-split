import subprocess
import sys

result = subprocess.run([sys.executable, "-m", "pip", "list"], capture_output=True, text=True, encoding="utf-8", errors="replace")
with open(r"C:\Users\zz_20\WorkBuddy\20260430112416\pip_list.txt", "w", encoding="utf-8") as f:
    f.write("STDOUT:\n")
    f.write(result.stdout)
    f.write("\nSTDERR:\n")
    f.write(result.stderr)
    f.write(f"\nReturn code: {result.returncode}")
