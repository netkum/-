"""VBA 소스(UTF-8)를 VBA 편집기 가져오기용(CP949 + CRLF)으로 변환한다.

사용법: python tools/build_vba_import.py
  vba/<member|master>/src/*.bas, *.cls  ->  import/  (CP949, CRLF)
  vba/<member|master>/src/*.txt         ->  import/  (UTF-8 BOM, CRLF, 붙여넣기용)
  같은 파일을 Scheduler/vba/<member|master>/ 에도 복사 (F:\\...\\Scheduler 같은 배포 폴더용)
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
for pkg in ["member", "master"]:
    src = ROOT / "vba" / pkg / "src"
    # vba/<pkg>/import  +  Scheduler/vba/<pkg> (배포용 폴더)
    for dst in (ROOT / "vba" / pkg / "import", ROOT / "Scheduler" / "vba" / pkg):
        dst.mkdir(parents=True, exist_ok=True)
        for f in sorted(src.iterdir()):
            text = f.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
            if f.suffix in (".bas", ".cls"):
                (dst / f.name).write_bytes(text.encode("cp949"))
            elif f.suffix == ".txt":
                (dst / f.name).write_bytes(text.encode("utf-8-sig"))
            print("built", dst / f.name)

# 엑셀 안에서 실행하는 빌더 (build.bat을 쓸 수 없을 때): Scheduler/Builder.bas
b = ROOT / "vba" / "builder" / "src" / "modBuilder.bas"
text = b.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
(ROOT / "Scheduler" / "Builder.bas").write_bytes(text.encode("cp949"))
print("built", ROOT / "Scheduler" / "Builder.bas")
