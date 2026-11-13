"""Evaluate the strikeout model against actual results.

Usage:  py evaluate.py            (reads data/history.csv)
"""
import csv
import statistics as st
from pathlib import Path

PATH = Path(__file__).parent / "data" / "history.csv"


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


rows = list(csv.DictReader(open(PATH, encoding="utf-8")))
graded = [r for r in rows if num(r["proj_ks"]) is not None and num(r["actual_ks"]) is not None]
errs = [num(r["proj_ks"]) - num(r["actual_ks"]) for r in graded]

print(f"Graded starts: {len(graded)}  ({min(r['date'] for r in graded)} to {max(r['date'] for r in graded)})")
print(f"Model MAE:   {st.mean(abs(e) for e in errs):.2f} Ks")
print(f"Model RMSE:  {(st.mean(e * e for e in errs)) ** 0.5:.2f} Ks")
print(f"Bias:        {st.mean(errs):+.2f} Ks (positive = over-projects)")
print(f"Within 1 K:  {sum(abs(e) <= 1 for e in errs) / len(errs):.1%}")
print(f"Within 2 Ks: {sum(abs(e) <= 2 for e in errs) / len(errs):.1%}")

# Baseline: always predict the average strikeouts per start
mean_actual = st.mean(num(r["actual_ks"]) for r in graded)
base_mae = st.mean(abs(num(r["actual_ks"]) - mean_actual) for r in graded)
print(f"\nBaseline (always predict {mean_actual:.2f}) MAE: {base_mae:.2f} Ks")

# Compare to the sportsbook line where one was logged
withline = [r for r in graded if num(r["vegas_ou"]) is not None]
if withline:
    m = st.mean(abs(num(r["proj_ks"]) - num(r["actual_ks"])) for r in withline)
    v = st.mean(abs(num(r["vegas_ou"]) - num(r["actual_ks"])) for r in withline)
    print(f"\nStarts with a logged line: {len(withline)}")
    print(f"  Model MAE: {m:.2f}   Line MAE: {v:.2f}")
    n = right = 0
    for r in withline:
        edge = num(r["proj_ks"]) - num(r["vegas_ou"])
        a, line = num(r["actual_ks"]), num(r["vegas_ou"])
        if abs(edge) <= 0.5 or a == line:
            continue  # skip small edges and pushes
        n += 1
        right += (edge > 0 and a > line) or (edge < 0 and a < line)
    print(f"  Right side of the line (edge > 0.5, no pushes): {right}/{n} = {right / n:.1%}")
