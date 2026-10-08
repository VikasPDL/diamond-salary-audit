"""Parse the biometric 'Daily In-Out Report' PDF into one row per employee."""
import re
from collections import Counter

import pypdf

HEADER_RE = re.compile(r"^(?P<name>.*?)Emp Code : Emp Name :(?P<code>\d+)\s+Punch Id :\s*(?P<punch>\d+)")
DEPT_RE = re.compile(r"^(?P<dept>.*?)Department :")
DAY_RE = re.compile(r"(?:^|\s)(?P<status>[A-Z]{1,4})?(?P<day>\d{2})-(?P<mon>[A-Za-z]{3})-(?P<year>\d{4})")
TIME_RE = re.compile(r"\b(\d{1,3}):(\d{2})\b(?!:)")

# Codes that mean the person did NOT work that day
NOT_WORKED = {"A", "WD", "H", "XX", None, ""}
WORK_ON_OFF_DAY = {"WP", "HWP", "PH", "HDH"}


def parse_pdf(file):
    """Return (employees, month_days, (mon, year)) where employees is a list of dicts with daily codes."""
    reader = pypdf.PdfReader(file)
    employees, cur, dept, period = [], None, "", None
    for page in reader.pages:
        for line in (page.extract_text() or "").splitlines():
            line = line.strip()
            m = DEPT_RE.match(line)
            if m:
                dept = m.group("dept").strip()
                continue
            m = HEADER_RE.match(line)
            if m:
                cur = {"code": int(m.group("code")), "name": m.group("name").strip(),
                       "department": dept, "days": {}, "ot_min": 0}
                employees.append(cur)
                continue
            if cur is None:
                continue
            m = DAY_RE.search(line)
            if m:
                day = int(m.group("day"))
                period = period or (m.group("mon"), int(m.group("year")))
                cur["days"][day] = m.group("status") or ""
                times = TIME_RE.findall(line[: m.start()])
                if times:  # last HH:MM before the status is OT hours
                    h, mi = times[-1]
                    cur["ot_min"] += int(h) * 60 + int(mi)
    month_days = max((max(e["days"]) for e in employees if e["days"]), default=31)
    return employees, month_days, period


def summarise(employees, month_days, sundays, holidays):
    """Add ABS / Sunday & Paid Holiday / PAID DAY using the rules from the attendance master."""
    off_days = set(sundays) | set(holidays)
    for e in employees:
        codes = e["days"]
        c = Counter(codes.values())
        e["abs"] = c["A"] + 0.5 * c["HD"]
        e["sun_ph"] = sum(
            1 for d, s in codes.items()
            if (d in off_days and s not in NOT_WORKED) or (d not in off_days and s in WORK_ON_OFF_DAY)
        )
        e["paid"] = month_days - e["abs"] + e["sun_ph"]
        e["present"] = sum(1 for s in codes.values() if s not in NOT_WORKED)
        e["absent"] = c["A"]
        e["ot_hrs"] = round(e["ot_min"] / 30) * 0.5  # total OT rounded to nearest half hour
        e["counts"] = c
    return employees


def detect_holidays(employees, threshold=0.5):
    """A date is a holiday when most employees have H / PH / HDH on it."""
    hc = Counter(d for e in employees for d, s in e["days"].items() if s in ("H", "PH", "HDH"))
    n = max(len(employees), 1)
    return sorted(d for d, k in hc.items() if k / n >= threshold)
