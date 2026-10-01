"""
Build Data/sales_manager_map.csv: a fictional sales manager -> salesperson mapping.

Derived from the already generated Data/branch_sales.csv, so run generate_mock_data.py
first. Salespeople are the order creators ("Sales Order Line / Created By") with positive
sales. The layout mirrors the real ERP export: one row per (manager, salesperson) pair, and
a salesperson who covers customers in two managers' regions appears under both.

- Crestline New Jersey Inc: four managers (the existing "sales leader" plus the three
  largest other target holders). Each salesperson reports to one manager; about a quarter
  also cover customers for a second manager. Managers also sell, so they map to themselves.
- Every other company: its existing "sales leader" manages all of its salespeople.
"""
import csv
import random
from collections import defaultdict
from pathlib import Path

import openpyxl

SEED = 2026
ROOT = Path(__file__).resolve().parent.parent / "Data"
NJ = "Crestline New Jersey Inc"
EXCLUDE = {"", "MAINSTREAM"}

rng = random.Random(SEED + 7)

sales = defaultdict(float)
leaders = defaultdict(lambda: defaultdict(int))
with open(ROOT / "branch_sales.csv", encoding="utf-8-sig", newline="") as f:
    for row in csv.DictReader(f):
        try:
            amount = float(row["amount"] or 0)
        except ValueError:
            continue
        creator = row["Sales Order Line / Created By"].strip()
        sales[row["company"], creator] += amount
        if row["sales leader"].strip():
            leaders[row["company"]][row["sales leader"].strip()] += 1

people = defaultdict(list)
for (company, creator), amount in sorted(sales.items(), key=lambda kv: -kv[1]):
    if creator not in EXCLUDE and amount > 0:
        people[company].append(creator)

ws = openpyxl.load_workbook(ROOT / "sales_targets_monthly.xlsx", read_only=True)["KPI by Month"]
rows = list(ws.iter_rows(values_only=True))
col = {name: i for i, name in enumerate(rows[0])}
target = defaultdict(float)
for r in rows[1:]:
    if r[col["Company"]] == NJ:
        target[r[col["Salesperson"]]] += r[col["KPI Target"]] or 0

pairs = set()
for company, sellers in people.items():
    main_leader = max(leaders[company], key=leaders[company].get) if leaders[company] else sellers[0]
    if company != NJ:
        for s in sellers:
            pairs.add((main_leader, s))
        continue
    others = [p for p, _ in sorted(target.items(), key=lambda kv: -kv[1]) if p != main_leader and p in sellers]
    managers = [main_leader] + others[:3]
    for m in managers:
        pairs.add((m, m))
    for i, s in enumerate(p for p in sellers if p not in managers):
        first = managers[i % len(managers)]
        pairs.add((first, s))
        if rng.random() < 0.25:
            pairs.add((rng.choice([m for m in managers if m != first]), s))

out = ROOT / "sales_manager_map.csv"
with open(out, "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Sales Manager", "Salesperson"])
    for m, s in sorted(pairs):
        w.writerow([m, s])
print(f"wrote Data/{out.name} ({len(pairs)} pairs)")
