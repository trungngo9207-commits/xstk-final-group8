"""Export the B1/B2 cells (formula + cached value) of the local workbook to cells.json for the report builders."""
import json, os
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
WB = os.path.join(REPO, "workbooks", "Group_8_local_snapshot_2026-10-11.xlsx")

wf, wv = openpyxl.load_workbook(WB), openpyxl.load_workbook(WB, data_only=True)
want = [("Part B", [f"{c}{r}" for r in range(3, 60) for c in "ABCDEFGH"] + [f"{c}{r}" for r in range(41, 47) for c in "KL"]),
        ("Part A", ["B18", "B19", "B23", "B24"]), ("Data", ["U1", "U2", "V1", "V2"])]
out = {}
for sheet, refs in want:
    for ref in refs:
        f = wf[sheet][ref].value
        if f is None:
            continue
        if isinstance(f, ArrayFormula):
            f = "{" + f.text.replace("TRUE()", "TRUE") + "}"
        out[f"{sheet}!{ref}"] = {"f": f if isinstance(f, (str, int, float)) else str(f), "v": wv[sheet][ref].value}
json.dump(out, open(os.path.join(HERE, "cells.json"), "w"), ensure_ascii=False, indent=0)
print(len(out), "cells written")
