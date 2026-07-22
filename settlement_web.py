"""
settlement_web.py  ─  웹 포털용 비대화형 정산 실행기 (v2.0 기반)
run_settlement.py v2.0 의 SettlementRunner 를 상속,
stdin 호출 없이 웹에서 파라미터를 받아 실행 후 결과 딕셔너리 반환.
"""

import datetime, logging
from pathlib import Path
from typing import Optional

from run_settlement import (
    SettlementRunner,
    MappingMaster,
    build_col_map,
    detect_col,
    _KT_CANDIDATES,
    _PL_CANDIDATES,
    _PL_OPT,
    _read_excel_smart,
)


class WebSettlementRunner(SettlementRunner):
    """
    웹 포털 전용 비대화형 정산 실행기.
    run_gui.py 의 _GUIRunner 패턴을 참조하여 작성.
    """

    def __init__(self,
                 kt_path: str,
                 pl_path: str,
                 period_start: datetime.date,
                 period_end: datetime.date,
                 output_dir: str,
                 mm_path: str = "",
                 rm_path: str = ""):
        # base_dir (output_dir) 를 super().__init__ 전에 생성해야 로그 파일 경로가 유효
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        super().__init__(
            base_dir=str(Path(output_dir)),
            start_date=period_start,
            end_date=period_end,
        )
        self._kt_path_str = kt_path
        self._pl_path_str = pl_path
        self._mm_path_str = mm_path
        self._rm_path_str = rm_path
        self._period_start = period_start
        self._period_end   = period_end

        # 로거에 stdout 핸들러가 이미 붙어 있으므로 파일 핸들러만 추가
        log_path = Path(output_dir) / "settlement_log.txt"
        if not any(isinstance(h, logging.FileHandler) and
                   getattr(h, 'baseFilename', '') == str(log_path)
                   for h in self.log.handlers):
            fh = logging.FileHandler(str(log_path), encoding="utf-8")
            fh.setFormatter(logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
            self.log.addHandler(fh)

    # ── 파일 탐지 오버라이드: stdin 없이 파라미터 사용 ────────────
    def detect_files(self):
        self.log.info("=" * 60)
        self.log.info("  KT 정산 자동화 시스템  v2.0 (Web)")
        self.log.info("=" * 60)

        self.kt_path = Path(self._kt_path_str) if self._kt_path_str else None
        self.pl_path = Path(self._pl_path_str) if self._pl_path_str else None
        self.rm_path = Path(self._rm_path_str) if self._rm_path_str else None
        mm_path      = Path(self._mm_path_str) if self._mm_path_str else None

        # 중견기업목록: output_dir 또는 kt 파일 폴더에서 탐색
        midcorp_path = None
        for search_dir in [self.base, Path(self._kt_path_str).parent]:
            _mc = search_dir / "중견기업목록.xlsx"
            if _mc.exists():
                midcorp_path = _mc
                break

        for label, p in [
            ("KT 파일       ", self.kt_path),
            ("플랫폼 파일   ", self.pl_path),
            ("소싱그룹 매핑 ", mm_path),
            ("반품 매핑     ", self.rm_path),
            ("중견기업목록  ", midcorp_path),
        ]:
            self.log.info(f"  {label}: {p or '없음'}")

        self.master = MappingMaster(mm_path, midcorp_path)
        self.master.load(self.log)

    # ── 컬럼 분석 오버라이드: 사용자 확인 단계 생략 ───────────────
    def step0_init(self):
        self.log.info("\n[STEP 0] 파일 컬럼 구조 분석")
        df_kt = _read_excel_smart(self.kt_path,
                                   search_cols=["구매문서번호"], nrows=3, dtype=str)
        df_pl = _read_excel_smart(self.pl_path,
                                   search_cols=["주문번호", "일정산번호"], nrows=3, dtype=str)

        self.kt_cols = build_col_map(df_kt, _KT_CANDIDATES)
        self.pl_cols = build_col_map(df_pl, _PL_CANDIDATES)
        self.pl_cols["요청번호"] = detect_col(df_pl, _PL_OPT["요청번호"], required=False)

        self.log.info("  KT 컬럼:")
        for k, v in self.kt_cols.items():
            self.log.info(f"    [{'O' if v else '-'}] {k:<16}: {v or '(미탐지)'}")
        self.log.info("  플랫폼 컬럼:")
        for k, v in self.pl_cols.items():
            self.log.info(f"    [{'O' if v else '-'}] {k:<16}: {v or '(없음)'}")

    # ── 정산 기간: 파라미터 그대로 사용 ──────────────────────────
    def get_period(self):
        self.log.info(f"  정산 기간: {self._period_start} ~ {self._period_end}")
        return self._period_start, self._period_end

    # ── 웹 실행 엔트리포인트 ─────────────────────────────────────
    def run_web(self) -> dict:
        """
        정산 실행 후 웹 대시보드용 결과 딕셔너리 반환.
        """
        self.run()   # SettlementRunner.run() 호출

        n_miss = len(self.df_missing) if self.df_missing is not None else 0
        n_diff = len(self.df_amtdiff) if self.df_amtdiff is not None else 0
        n_rng  = len(self.df_ret_ng)  if self.df_ret_ng  is not None else 0

        # 미리보기용 행 데이터
        miss_cols = ["주문번호", "입고일", "정산금액", "협력사명", "이동유형"]
        diff_cols = ["주문번호", "정산금액", "플랫폼정산금액", "금액차이", "협력사명"]

        def _to_rows(df, preferred_cols):
            if df is None or df.empty:
                return []
            available = [c for c in preferred_cols if c in df.columns]
            if not available:
                available = list(df.columns[:5])
            return df[available].fillna("").head(20).to_dict("records")

        # 정산금액 합계 (인보이스 기준)
        total_amt = 0
        if self.df_invoice is not None and "정산금액" in self.df_invoice.columns:
            import pandas as pd
            total_amt = int(pd.to_numeric(
                self.df_invoice["정산금액"], errors="coerce").fillna(0).sum())

        return {
            "period_start"  : self._period_start.strftime("%Y-%m-%d"),
            "period_end"    : self._period_end.strftime("%Y-%m-%d"),
            "kt_total"      : len(self.df_kt)      if self.df_kt      is not None else 0,
            "pl_total"      : len(self.df_pl)      if self.df_pl      is not None else 0,
            "matched"       : len(self.df_matched) if self.df_matched  is not None else 0,
            "missing"       : n_miss,
            "pl_only"       : len(self.df_pl_only) if self.df_pl_only is not None else 0,
            "amount_diff"   : n_diff,
            "ret_matched"   : len(self.df_ret_ok)  if self.df_ret_ok  is not None else 0,
            "ret_unmatch"   : n_rng,
            "total_amount"  : total_amt,
            "status"        : "정상" if (n_miss + n_diff + n_rng) == 0 else "확인필요",
            "result_path"   : getattr(self, "result_path", ""),
            "missing_rows"  : _to_rows(self.df_missing, miss_cols),
            "diff_rows"     : _to_rows(self.df_amtdiff, diff_cols),
        }
