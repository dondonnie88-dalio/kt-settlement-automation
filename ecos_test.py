"""
ECOS API 진단 스크립트 — demand_forecast.py와 같은 폴더에서 실행
python ecos_test.py
"""
import requests, json, sys
from pathlib import Path

# ── API 키 로드 ──────────────────────────────────────────────────────────
KEY_FILE = Path(__file__).parent / "ecos_api_key.txt"
if KEY_FILE.exists():
    API_KEY = KEY_FILE.read_text(encoding="utf-8").strip()
    print(f"[키 파일] {KEY_FILE} 로드됨: {API_KEY[:6]}...{API_KEY[-4:]}")
else:
    print(f"[오류] {KEY_FILE} 없음 — 파일을 만들고 키를 저장하세요")
    sys.exit(1)

BASE = f"https://ecos.bok.or.kr/api"

# ── 1단계: 통계 목록 검색으로 정확한 코드 확인 ────────────────────────────
print("\n[1단계] 소비자동향조사 통계 검색...")
url = f"{BASE}/StatisticSearch/{API_KEY}/json/kr/1/10/521Y001/MM/202501/202506"
try:
    r = requests.get(url, timeout=15)
    print(f"  HTTP {r.status_code}")
    data = r.json()
    rows = data.get("StatisticSearch", {}).get("row", [])
    if rows:
        print(f"  ✓ 데이터 {len(rows)}건 — 첫 행: {rows[0]}")
    else:
        print(f"  ✗ 데이터 없음 — 응답: {json.dumps(data, ensure_ascii=False)[:200]}")
except Exception as e:
    print(f"  ✗ 연결 오류: {e}")

# ── 2단계: 항목 코드 목록 조회 ───────────────────────────────────────────
print("\n[2단계] 521Y001 통계의 항목 코드 목록...")
url = f"{BASE}/StatisticItemList/{API_KEY}/json/kr/1/50/521Y001"
try:
    r = requests.get(url, timeout=15)
    data = r.json()
    items = data.get("StatisticItemList", {}).get("row", [])
    if items:
        for item in items[:10]:
            print(f"  {item.get('ITEM_CODE','?'):12} | {item.get('ITEM_NAME','?')}")
    else:
        print(f"  응답: {json.dumps(data, ensure_ascii=False)[:300]}")
except Exception as e:
    print(f"  오류: {e}")

# ── 3단계: 후보 코드 전수 테스트 ─────────────────────────────────────────
print("\n[3단계] 후보 시리즈 코드 테스트...")
candidates = [
    ("소비자심리지수",   "521Y001", "I22B"),
    ("소비자심리지수",   "521Y001", "I22A"),
    ("소비자심리지수",   "521Y001", "I22C"),
    ("경기동행지수",     "901Y062", ""),
    ("경기동행지수",     "901Y067", ""),
    ("경기선행지수",     "901Y063", ""),
]
for label, stat, item in candidates:
    url = f"{BASE}/StatisticSearch/{API_KEY}/json/kr/1/3/{stat}/MM/202501/202506"
    if item:
        url += f"/{item}"
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
        rows = data.get("StatisticSearch", {}).get("row", [])
        if rows:
            sample = rows[-1]
            print(f"  ✓ {label}({stat}/{item or '-'}): {sample.get('TIME')}={sample.get('DATA_VALUE')}")
        else:
            print(f"  ✗ {label}({stat}/{item or '-'}): {str(data)[:80]}")
    except Exception as e:
        print(f"  ✗ {label}: 연결 오류 — {e}")

print("\n[완료] 위 결과에서 ✓ 표시된 stat_code / item_code를 알려주시면 코드를 수정하겠습니다.")
