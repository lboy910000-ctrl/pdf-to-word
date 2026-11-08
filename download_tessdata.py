"""Download the Bengali / Hindi / English OCR data (tessdata_best, ~38 MB) into ./tessdata.
Run once, with internet:   python download_tessdata.py
After that the app works fully offline."""
import os
import sys
import urllib.request
from pathlib import Path

BASE_URL = os.environ.get("TESSDATA_BASE_URL", "https://github.com/tesseract-ocr/tessdata_best/raw/main")
LANGS = ("ben", "hin", "eng")


def main(out_dir=None):
    out = Path(out_dir) if out_dir else Path(__file__).resolve().parent / "tessdata"
    out.mkdir(parents=True, exist_ok=True)
    for lang in LANGS:
        dest = out / f"{lang}.traineddata"
        if dest.exists() and dest.stat().st_size > 1_000_000:
            print(f"{lang}: already present")
            continue
        url = f"{BASE_URL}/{lang}.traineddata"
        print(f"{lang}: downloading {url}")
        tmp = dest.with_suffix(".part")
        with urllib.request.urlopen(url, timeout=120) as r, open(tmp, "wb") as f:
            while chunk := r.read(1 << 20):
                f.write(chunk)
        if tmp.stat().st_size < 1_000_000:
            tmp.unlink()
            sys.exit(f"{lang}: download looks wrong (too small). Check your internet connection.")
        tmp.replace(dest)
    print("Done:", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
