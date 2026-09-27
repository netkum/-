"""VBA 소스(UTF-8)를 VBA 편집기 가져오기용(CP949 + CRLF)으로 변환한다.

사용법: python tools/build_vba_import.py
  vba/<member|master>/src/*.bas, *.cls  ->  import/  (CP949, CRLF)
  vba/<member|master>/src/*.txt         ->  import/  (UTF-8 BOM, CRLF, 붙여넣기용)
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
for pkg in ["member", "master"]:
    src = ROOT / "vba" / pkg / "src"
    dst = ROOT / "vba" / pkg / "import"
    dst.mkdir(parents=True, exist_ok=True)
    for f in sorted(src.iterdir()):
        text = f.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
        if f.suffix in (".bas", ".cls"):
            (dst / f.name).write_bytes(text.encode("cp949"))
        elif f.suffix == ".txt":
            (dst / f.name).write_bytes(text.encode("utf-8-sig"))
        print("built", dst / f.name)
