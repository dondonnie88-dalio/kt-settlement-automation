"""
통신/일반 미분류 원인 진단 + 미분류 카테고리 목록 Excel 출력
실행: python diagnose_prod_type.py
결과: 미분류_카테고리목록.xlsx 생성
"""
import sys, json
sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd, openpyxl
from pathlib import Path

BASE_DIR = Path(__file__).parent
cfg_path = BASE_DIR / "forecast_config.json"
pq_path  = BASE_DIR / "processed_data.parquet"

cfg     = json.loads(cfg_path.read_text(encoding="utf-8"))
col_map = cfg.get("col_map", {})
cat_col = col_map.get("category")

print(f"데이터 카테고리 컬럼명: {cat_col!r}")

# ── 매핑 파일 로드 ────────────────────────────────────────────────────────────
cat_keywords = ["서비스", "카테고리", "담당자"]
map_file = next(
    (p for p in BASE_DIR.glob("*.xlsx")
     if not p.name.startswith("~$") and all(k in p.name for k in cat_keywords)), None
)
if not map_file:
    print("매핑 파일 없음"); sys.exit(1)

print(f"매핑 파일: {map_file.name}")

# 매핑 딕셔너리 구성 (demand_forecast.py 와 동일 로직)
TELECOM_KEYWORDS = ["네트워크", "network", "통신", "회선", "NW"]

def dept_to_type(dept):
    s = str(dept).strip()
    if not s or s in ("nan", "None", ""):
        return "미분류"
    return "통신" if any(k in s for k in TELECOM_KEYWORDS) else "일반"

def read_sheet(wb_path, sheet_name):
    df = pd.read_excel(wb_path, sheet_name=sheet_name, engine="openpyxl", header=0)
    df.columns = df.columns.str.strip()
    dept_col = next((c for c in df.columns if "담당부서" in c), None)
    if not dept_col:
        return {}
    KEY_COLS = ["서비스 카테고리", "카테고리 분류", "카테고리명"]
    key_cols = [c for c in KEY_COLS if c in df.columns]
    result = {}
    for kc in reversed(key_cols):
        for _, row in df.iterrows():
            cat = str(row[kc]).strip()
            dept = str(row[dept_col]).strip()
            if cat and cat not in ("nan", "None"):
                result[cat] = dept_to_type(dept)
    return result

wb_tmp = openpyxl.load_workbook(map_file, read_only=True, data_only=True)
sheets = wb_tmp.sheetnames
wb_tmp.close()

new_sheets = [s for s in sheets if "카테고리" in s and "담당자" in s and "개편전" not in s]
old_sheets = [s for s in sheets if "개편전" in s]

mapping = {}
for sh in old_sheets:
    mapping.update(read_sheet(map_file, sh))
for sh in new_sheets:
    mapping.update(read_sheet(map_file, sh))   # 신형식 우선(덮어씀)

print(f"매핑 카테고리 수: {len(mapping)}개 "
      f"(통신 {sum(1 for v in mapping.values() if v=='통신')}개 / "
      f"일반 {sum(1 for v in mapping.values() if v=='일반')}개)")

# ── parquet 로드 및 분류 적용 ─────────────────────────────────────────────────
df = pd.read_parquet(pq_path, engine="pyarrow")
print(f"전체 행수: {len(df):,}")

if not cat_col or cat_col not in df.columns:
    print(f"카테고리 컬럼({cat_col!r})이 데이터에 없음"); sys.exit(1)

df["_prod_type"] = df[cat_col].astype(str).str.strip().map(mapping).fillna("미분류")

print()
print("=" * 60)
print("통신/일반 구분 결과")
print("=" * 60)
vc = df["_prod_type"].value_counts()
for k, v in vc.items():
    print(f"  {k:6s}: {v:>10,}행 ({v/len(df)*100:.1f}%)")

# ── 미분류 카테고리 분석 ──────────────────────────────────────────────────────
miss_df = df[df["_prod_type"] == "미분류"]
miss_cats = (
    miss_df[cat_col]
    .astype(str).str.strip()
    .value_counts()
    .reset_index()
)
miss_cats.columns = ["서비스카테고리", "행수"]
miss_cats = miss_cats[miss_cats["서비스카테고리"].str.len() > 0]
miss_cats["담당부서(안)"] = ""   # 사용자가 채워 넣을 컬럼
miss_cats["통신/일반"]   = ""   # 사용자가 채워 넣을 컬럼

print()
print("=" * 60)
print(f"미분류 카테고리: {len(miss_cats)}종, {len(miss_df):,}행")
print("(상위 20개)")
print("=" * 60)
for _, row in miss_cats.head(20).iterrows():
    print(f"  {row['서비스카테고리'][:50]:50s}  {row['행수']:>8,}행")

# ── Excel 출력 ────────────────────────────────────────────────────────────────
out_path = BASE_DIR / "미분류_카테고리목록.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    # 시트1: 미분류 목록 (담당부서 입력용)
    miss_cats.to_excel(writer, sheet_name="미분류_목록", index=False)
    ws = writer.sheets["미분류_목록"]
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 12

    # 시트2: 현재 매핑 전체 (참고용)
    map_df = pd.DataFrame(
        [(k, v) for k, v in mapping.items()],
        columns=["서비스카테고리", "통신/일반"]
    ).sort_values("서비스카테고리")
    map_df.to_excel(writer, sheet_name="현재_매핑목록", index=False)
    writer.sheets["현재_매핑목록"].column_dimensions["A"].width = 60

print()
print(f"✓ 미분류_카테고리목록.xlsx 생성 완료")
print(f"  → C열(담당부서)에 팀명을 입력하거나")
print(f"     D열(통신/일반)에 직접 '통신' 또는 '일반'을 입력한 후")
print(f"     매핑 파일의 해당 시트에 복사해 넣으시면 됩니다.")
