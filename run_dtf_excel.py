
# Quick script to check pandas and run dtf_excel_processor.py
import subprocess, sys, os

print("=== DTF Excel Processor Setup Check ===")
print()

# Check pandas
try:
    import pandas as pd
    print(f"pandas: OK ({pd.__version__})")
except ImportError:
    print("pandas: MISSING — installing...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"], check=True)
    print("pandas: installed")

# Check image folder
img_folder = r"W:\DTFUnshippedImages_20260411_031645"
if os.path.isdir(img_folder):
    count = len(os.listdir(img_folder))
    print(f"Image folder: OK ({count} files)")
else:
    print(f"Image folder: NOT FOUND at {img_folder}")

# Check Excel file
excel_path = r"W:\VarsaniAutomation\UnshippedDTFOrders_11042026_031657__1_.xlsx"
if os.path.isfile(excel_path):
    print(f"Excel file: OK")
else:
    print(f"Excel file: NOT FOUND")
    print(f"  >>> Please save the Excel file to: {excel_path}")
    print(f"  >>> Then run: python W:\\VarsaniAutomation\\dtf_excel_processor.py")
    sys.exit(1)

print()
print("All checks passed — running processor...")
print()
exec(open(r"W:\VarsaniAutomation\dtf_excel_processor.py").read())
