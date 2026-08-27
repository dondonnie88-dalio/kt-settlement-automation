"""
정산 자동화 + 과세구분 검증 — PyInstaller 빌드 스크립트
외부망(집 등)에서 실행하세요.

사전 준비:
    pip install pyinstaller

실행:
    python build_exe.py
"""
import subprocess, sys, shutil
from pathlib import Path

BASE = Path(__file__).parent  # C:\FTC_downloads\정산도구

BUILDS = [
    {
        "name"   : "정산자동화",
        "script" : str(BASE / "settlement_gui.py"),
        "icon"   : None,
        "distdir": str(BASE / "dist"),
    },
    {
        "name"   : "과세구분검증",
        "script" : str(BASE.parent / "파이썬 과세구분 검증" / "tax_verify_gui.py"),
        "icon"   : None,
        "distdir": str(BASE / "dist"),
        # tax_verify_gui가 run_tax_verify를 import하므로 같은 폴더에서 빌드
        "workdir": str(BASE.parent / "파이썬 과세구분 검증"),
    },
]

for b in BUILDS:
    print(f"\n{'='*50}")
    print(f"빌드: {b['name']}")
    print(f"{'='*50}")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",                         # 폴더 모드: 압축해제 없어 백신 오탐 없음
        "--clean",                          # 이전 캐시 제거
        "--noconfirm",                      # 덮어쓰기 확인 생략
        f"--name={b['name']}",
        f"--distpath={b['distdir']}",
        # pandas / openpyxl hidden imports
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.styles",
        "--hidden-import=openpyxl.utils",
        "--hidden-import=pandas",
        "--hidden-import=xlrd",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.scrolledtext",
        "--collect-all=openpyxl",
        "--collect-all=pandas",
    ]

    if b.get("icon"):
        cmd.append(f"--icon={b['icon']}")

    cmd.append(b["script"])

    workdir = b.get("workdir", str(BASE))
    result = subprocess.run(cmd, cwd=workdir)

    if result.returncode == 0:
        out_dir = Path(b["distdir"]) / b["name"]   # --onedir 결과 폴더
        exe = out_dir / f"{b['name']}.exe"
        if exe.exists():
            # 실행.bat 생성 (오류가 나도 창이 유지됨)
            bat = out_dir / "실행.bat"
            bat.write_text(
                f"@echo off\nchcp 65001 > nul\necho {b['name']} 시작...\n{b['name']}.exe\npause\n",
                encoding="utf-8"
            )
            # 배포용 zip (폴더째 압축)
            zip_path = Path(b["distdir"]) / f"{b['name']}_배포"
            shutil.make_archive(str(zip_path), "zip", str(out_dir.parent), b["name"])
            print(f"\n✅ 빌드 완료: {exe}")
            print(f"📦 배포용 zip: {zip_path}.zip  (압축 풀고 폴더 안 exe 실행)")
    else:
        print(f"\n❌ 빌드 실패: {b['name']}")

print("\n모든 빌드 완료.")
print(f"배포 파일 위치: {BASE / 'dist'}")
