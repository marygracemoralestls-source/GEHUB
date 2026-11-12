"""
Builds data.json for the GE Team Hub from the GE Tracker Excel file.
Usage: python build_data.py tracker.xlsx data.json
Set the environment variable INCLUDE_GUEST_NAMES=false to leave guest names out.
"""
import datetime as dt
import json
import os
import re
import sys
from openpyxl import load_workbook

SRC, OUT = sys.argv[1], sys.argv[2]
INCLUDE_GUEST_NAMES = os.environ.get("INCLUDE_GUEST_NAMES", "true").lower() != "false"
DAYS_BACK = 14              # include stays that started up to this many days ago
ENDORSEMENT_DAYS_BACK = 30  # include endorsements for arrivals this far back
DUTY_DAYS_AHEAD = 21        # how far ahead of today to export the duty schedule

ZONE_TABS = [("East", "East"), ("Kamala Beach", "Kamala Beach"), ("Samsara", "Samsara & North Phuket"),
             ("Samui", "Samui & Natai"), ("Millionaires", "Millionaires Mile")]
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


def text(v):
    return "" if v is None else re.sub(r"\s+", " ", str(v)).strip()


def to_iso(v):
    """Date cell -> YYYY-MM-DD. Handles real dates and dates typed as text (13/11/2026 = day first)."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, str):
        s = v.strip()
        m = re.search(r"(\d{1,2})[\s\-/]([A-Za-z]{3})[A-Za-z]*[\s\-/,]*(\d{4})", s)
        if m and m.group(2).lower() in MONTHS:
            return "%s-%02d-%02d" % (m.group(3), MONTHS.index(m.group(2).lower()) + 1, int(m.group(1)))
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
        if m:
            return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    return ""


def redact(booking):
    """Keep villa, dates and the booking reference; drop the guest name."""
    parts = booking.split(" - ")
    if len(parts) < 3:
        return booking
    last = parts[-1]
    looks_like_ref = re.match(r"^[#A-Za-z0-9\-]{6,}$", re.sub(r"\s*\(.*\)$", "", last))
    return " - ".join(parts[:2]) + (" - " + last if looks_like_ref and len(parts) > 3 else "")


def link_of(cell_f, cell_v=None):
    """URL from a real hyperlink, or from a =HYPERLINK("url","label") formula."""
    if cell_f.hyperlink is not None and cell_f.hyperlink.target:
        return cell_f.hyperlink.target
    m = re.search(r'HYPERLINK\(\s*"([^"]+)"', str(cell_f.value or ""), re.I)
    return m.group(1) if m else ""


def find_sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    return None


wb = load_workbook(SRC, data_only=True)   # values
wf = load_workbook(SRC)                   # formulas + hyperlinks
today = dt.datetime.now(dt.timezone.utc).date()
iso = lambda d: d.isoformat()
cutoff = iso(today - dt.timedelta(days=DAYS_BACK))
end_cutoff = iso(today - dt.timedelta(days=ENDORSEMENT_DAYS_BACK))
arrivals, endorsements, warnings = [], [], []

for ws in wb.worksheets:
    zone = next((label for start, label in ZONE_TABS if ws.title.startswith(start)), None)
    if not zone:
        continue
    wsf = wf[ws.title]
    for r, row in enumerate(ws.iter_rows(min_row=2, max_col=15, values_only=True), start=2):
        arrival = to_iso(row[4])
        if not arrival or arrival < cutoff or not text(row[2]):
            continue
        departure = to_iso(row[5])
        if not departure or departure < arrival:
            warnings.append("%s row %d: check departure date (%s)" % (zone, r, text(row[2])[:40]))
        booking = text(row[2])
        arrivals.append(dict(
            zone=zone, cross=text(row[0]), airbnb=text(row[1]) != "",
            booking=booking if INCLUDE_GUEST_NAMES else redact(booking),
            iti=link_of(wsf.cell(row=r, column=3)), crm=link_of(wsf.cell(row=r, column=4)),
            arrival=arrival, departure=departure,
            transfer=text(row[7]), meals=text(row[8]), other=text(row[9]), wol=text(row[10]),
            fi=text(row[11]), welcome=text(row[12]), feedback=text(row[13]), notes=text(row[14])))
arrivals.sort(key=lambda a: (a["arrival"], a["zone"]))

ws = find_sheet(wb, "Endorsement")
if ws:
    for row in ws.iter_rows(min_row=2, max_col=7, values_only=True):
        arrival = to_iso(row[1])
        if not text(row[0]) or not arrival or arrival < end_cutoff:
            continue
        endorsements.append(dict(
            ref=text(row[0]) if INCLUDE_GUEST_NAMES else redact(text(row[0])), arrival=arrival,
            task=text(row[3]), status=text(row[4]) or "None",
            notes=" ".join(x for x in [text(row[5]), text(row[6])] if x)))
    endorsements.sort(key=lambda e: e["arrival"])
else:
    warnings.append("No tab named Endorsement was found.")

duty = dict(duties=[], days=[])
ws = find_sheet(wb, "Duty")
if ws:
    rows = list(ws.iter_rows(values_only=True))
    headers = [text(h) for h in rows[0][1:]]
    while headers and headers[-1] == "":
        headers.pop()
    duty["duties"] = headers
    lo, hi = iso(today - dt.timedelta(days=1)), iso(today + dt.timedelta(days=DUTY_DAYS_AHEAD))
    for row in rows[1:]:
        d = to_iso(row[0])
        if d and lo <= d <= hi:
            duty["days"].append(dict(date=d, who=[text(row[i + 1]) if i + 1 < len(row) else "" for i in range(len(headers))]))
    duty["days"].sort(key=lambda x: x["date"])
else:
    warnings.append("No tab named DUTY was found.")

out = dict(generatedAt=dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
           duty=duty, arrivals=arrivals, endorsements=endorsements, warnings=warnings)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False)
print("arrivals", len(arrivals), "| endorsements", len(endorsements), "| duty days", len(duty["days"]), "| warnings", len(warnings))
