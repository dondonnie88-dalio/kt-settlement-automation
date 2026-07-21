"""
app.py  ─  KT 정산 자동화 웹 포털
실행: python app.py
접속: http://localhost:5000
"""

import os, uuid, json, datetime
from pathlib import Path
from flask import (Flask, render_template, request, redirect, url_for,
                   send_file, jsonify, flash)
from werkzeug.utils import secure_filename

from settlement_web import WebSettlementRunner

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
ALLOWED_EXT = {".xlsx", ".xls", ".xlsm"}

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


# ── 홈 / 업로드 폼 ───────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


# ── 정산 실행 ─────────────────────────────────────────────────────
@app.route("/run", methods=["POST"])
def run_settlement():
    # 파일 검증
    kt_file = request.files.get("kt_file")
    pl_file = request.files.get("pl_file")

    if not kt_file or not kt_file.filename:
        flash("KT 정산 파일을 선택해주세요.", "error")
        return redirect(url_for("index"))
    if not pl_file or not pl_file.filename:
        flash("플랫폼 정산 파일을 선택해주세요.", "error")
        return redirect(url_for("index"))
    if not allowed_file(kt_file.filename):
        flash("Excel 파일(.xlsx / .xls)만 업로드 가능합니다.", "error")
        return redirect(url_for("index"))
    if not allowed_file(pl_file.filename):
        flash("Excel 파일(.xlsx / .xls)만 업로드 가능합니다.", "error")
        return redirect(url_for("index"))

    # 기간 파싱
    period_mode = request.form.get("period_mode", "manual")
    try:
        if period_mode == "auto":
            today = datetime.date.today()
            if today.day == 16:
                period_start = datetime.date(today.year, today.month, 1)
                period_end   = datetime.date(today.year, today.month, 15)
            elif today.day == 1:
                import calendar
                y, m = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)
                period_start = datetime.date(y, m, 16)
                period_end   = datetime.date(y, m, calendar.monthrange(y, m)[1])
            else:
                flash("자동 기간 감지는 매월 1일 또는 16일에만 가능합니다. 직접 입력을 사용하세요.", "error")
                return redirect(url_for("index"))
        else:
            period_start = datetime.date.fromisoformat(request.form["period_start"])
            period_end   = datetime.date.fromisoformat(request.form["period_end"])
    except (ValueError, KeyError):
        flash("정산 기간을 올바르게 입력해주세요 (YYYY-MM-DD).", "error")
        return redirect(url_for("index"))

    if period_start > period_end:
        flash("시작일이 종료일보다 늦습니다.", "error")
        return redirect(url_for("index"))

    # 파일 저장
    run_id  = str(uuid.uuid4())[:8]
    run_dir = UPLOAD_DIR / run_id
    run_dir.mkdir(parents=True)

    kt_filename = "kt_raw" + Path(secure_filename(kt_file.filename)).suffix
    pl_filename = "platform" + Path(secure_filename(pl_file.filename)).suffix
    kt_path = run_dir / kt_filename
    pl_path = run_dir / pl_filename
    kt_file.save(str(kt_path))
    pl_file.save(str(pl_path))

    # 정산 실행
    out_dir = OUTPUT_DIR / run_id
    try:
        runner = WebSettlementRunner(
            kt_path      = str(kt_path),
            pl_path      = str(pl_path),
            period_start = period_start,
            period_end   = period_end,
            output_dir   = str(out_dir),
        )
        result = runner.run()
    except Exception as e:
        flash(f"정산 실행 중 오류가 발생했습니다: {e}", "error")
        return redirect(url_for("index"))

    # 결과 저장 (JSON) — date/datetime 객체를 문자열로 변환
    def _json_default(obj):
        if hasattr(obj, "isoformat"):
            return obj.isoformat()
        raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")

    result_json = out_dir / "result.json"
    with open(result_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=_json_default)

    return redirect(url_for("show_result", run_id=run_id))


# ── 결과 대시보드 ─────────────────────────────────────────────────
@app.route("/result/<run_id>")
def show_result(run_id: str):
    result_json = OUTPUT_DIR / run_id / "result.json"
    if not result_json.exists():
        flash("결과를 찾을 수 없습니다.", "error")
        return redirect(url_for("index"))

    with open(result_json, encoding="utf-8") as f:
        result = json.load(f)

    return render_template("result.html", r=result, run_id=run_id)


# ── Excel 다운로드 ────────────────────────────────────────────────
@app.route("/download/<run_id>/result")
def download_result(run_id: str):
    out_dir = OUTPUT_DIR / run_id
    files   = list(out_dir.glob("정산결과_*.xlsx"))
    if not files:
        return "파일을 찾을 수 없습니다.", 404
    return send_file(str(files[0]),
                     as_attachment=True,
                     download_name=files[0].name)


if __name__ == "__main__":
    print("=" * 55)
    print("  KT 정산 자동화 웹 포털")
    print("  접속 주소: http://localhost:5000")
    print("=" * 55)
    app.run(debug=False, host="0.0.0.0", port=5000)
