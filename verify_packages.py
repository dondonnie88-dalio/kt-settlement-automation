pkgs = [
    ("pandas",      "pandas"),
    ("numpy",       "numpy"),
    ("statsmodels", "statsmodels"),
    ("scikit-learn","sklearn"),
    ("openpyxl",    "openpyxl"),
    ("xlsxwriter",  "xlsxwriter"),
]
ok = True
for name, mod in pkgs:
    try:
        m = __import__(mod)
        print(f"  OK  {name:15} {getattr(m, '__version__', '?')}")
    except ImportError:
        print(f"  NG  {name:15} NOT FOUND")
        ok = False
import sys
sys.exit(0 if ok else 1)
