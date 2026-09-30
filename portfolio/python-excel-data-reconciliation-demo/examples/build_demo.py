from csv import DictReader
from pathlib import Path

from openpyxl import Workbook

from recon.core import reconcile_workbooks, write_report


ROOT = Path(__file__).parents[1]


def csv_to_xlsx(csv_path: Path, xlsx_path: Path) -> None:
    with csv_path.open(encoding="utf-8", newline="") as handle:
        rows = list(DictReader(handle))

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ComponentID", "Status", "Value"])
    for row in rows:
        ws.append(
            [
                row["ComponentID"],
                row["Status"],
                float(row["Value"]),
            ]
        )
    wb.save(xlsx_path)


left = ROOT / "source_a.xlsx"
right = ROOT / "source_b.xlsx"
report = ROOT / "reconciliation_report.xlsx"

csv_to_xlsx(ROOT / "data" / "source_a.csv", left)
csv_to_xlsx(ROOT / "data" / "source_b.csv", right)

result = reconcile_workbooks(left, right)
write_report(result, report)

print(f"matched_rows={result.matched_rows}")
print(f"differences={len(result.differences)}")
print(f"report={report}")
