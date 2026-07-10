from pathlib import Path
import csv
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT / "Database" / "products.csv"


def main() -> int:
    errors = []
    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            for field in ("Asset_primary", "Asset_secondary"):
                rel = row.get(field, "").strip()
                if not rel:
                    continue
                path = ROOT / rel
                if not path.exists():
                    errors.append(f"Missing: {rel}")
                    continue
                try:
                    with Image.open(path) as image:
                        image.verify()
                except Exception as exc:
                    errors.append(f"Invalid image {rel}: {exc}")
    if errors:
        print("\n".join(errors))
        return 1
    print("All product assets are present and readable.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
