import sys
import re
from urllib.parse import urlparse

from openpyxl import load_workbook


def valid_email(email):
    if not email:
        return False
    if not re.match(r'^[^@]+@[^@]+\.[^@]{2,}$', email):
        return False
    local, domain = email.split("@")
    if domain.lower() in ("example.com", "domain.com", "yourdomain.com", "email.com", "test.com"):
        return False
    if len(local) <= 1:
        return False
    return True


def valid_website_url(url):
    if not url:
        return False
    parsed = urlparse(url)
    return bool(parsed.scheme and parsed.netloc and "." in parsed.netloc)


def valid_phone(phone):
    if not phone:
        return False
    digits = re.sub(r'\D', '', phone)
    return 8 <= len(digits) <= 15


def valid_hall(hall):
    if not hall:
        return True
    return hall.isdigit()


def valid_stand(stand):
    if not stand:
        return True
    return bool(re.match(r'^[A-Za-z0-9]+$', stand))


def valid_detail_url(url):
    if not url:
        return False
    parsed = urlparse(url)
    return bool(parsed.scheme and parsed.netloc)


def normalize_domain(url):
    try:
        domain = urlparse(url).netloc.lower()
        return domain[4:] if domain.startswith("www.") else domain
    except Exception:
        return ""


def collect_format_issues(rows):
    issues = []

    for idx, row in enumerate(rows):
        name = row.get("Company Name", "")
        website = row.get("Website", "")
        phone = row.get("Phone", "")
        email = row.get("Email", "")
        hall = row.get("Hall", "")
        stand = row.get("Stand", "")
        detail_url = row.get("Detail Page URL", "")
        rn = idx + 2

        if email and not valid_email(email):
            issues.append(("email", rn, name, email))

        if website and not valid_website_url(website):
            issues.append(("website_format", rn, name, website))

        if phone and not valid_phone(phone):
            issues.append(("phone", rn, name, phone))

        if hall and not valid_hall(hall):
            issues.append(("hall", rn, name, hall))

        if stand and not valid_stand(stand):
            issues.append(("stand", rn, name, stand))

        if detail_url and not valid_detail_url(detail_url):
            issues.append(("detail_url", rn, name, detail_url))

    return issues


def collect_duplicates(rows):
    dupe_groups = []

    buckets = {
        "company": {},
        "website": {},
        "email": {},
        "detail_url": {},
    }

    for idx, row in enumerate(rows):
        name = row.get("Company Name", "")
        website = row.get("Website", "")
        email = row.get("Email", "")
        hall = row.get("Hall", "")
        stand = row.get("Stand", "")
        detail_url = row.get("Detail Page URL", "")
        rn = idx + 2

        if name:
            key = name.strip().lower()
            buckets["company"].setdefault(key, []).append((rn, name))

        if website:
            key = normalize_domain(website)
            if key:
                buckets["website"].setdefault(key, []).append((rn, website))

        if email:
            key = email.strip().lower()
            buckets["email"].setdefault(key, []).append((rn, email))

        if detail_url:
            key = detail_url.strip()
            buckets["detail_url"].setdefault(key, []).append((rn, detail_url))

    for label, bucket in buckets.items():
        for key, entries in bucket.items():
            if len(entries) > 1:
                dupe_groups.append((label, key, entries))

    return dupe_groups


def print_report(filepath, rows, issues, dupe_groups):
    total = len(rows)
    total_format = len(issues)
    total_dupe_rows = sum(len(e[2]) for e in dupe_groups)

    print("=" * 60)

    fmt_issues = {}
    for cat, rn, name, val in issues:
        fmt_issues.setdefault(cat, []).append((rn, name, val))

    fmt_labels = {
        "email": "Email",
        "website_format": "Website (URL)",
        "phone": "Phone",
        "hall": "Hall",
        "stand": "Stand",
        "detail_url": "Detail Page URL",
    }

    print(" FORMAT ISSUES")
    print("-" * 60)
    any_fmt = False
    for cat in ["email", "website_format", "phone", "hall", "stand", "detail_url"]:
        items = fmt_issues.get(cat, [])
        label = fmt_labels[cat]
        if items:
            any_fmt = True
            print(f"  {len(items):>4}  {label}")
            for rn, name, val in items[:10]:
                print(f"         Row {rn}: {val}")
            if len(items) > 10:
                print(f"         ... and {len(items) - 10} more")
    if not any_fmt:
        print("  (none)")

    print()
    print(" DUPLICATES")
    print("-" * 60)
    dupe_labels = {
        "company": "Company Name",
        "website": "Website domain",
        "email": "Email",
        "detail_url": "Detail Page URL",
    }
    any_dupe = False
    for label, key, entries in dupe_groups:
        any_dupe = True
        rows_str = ", ".join(f"Row {r[0]}" for r in entries)
        print(f"  {len(entries)}x  {dupe_labels[label]}: {key}")
        print(f"         {rows_str}")
    if not any_dupe:
        print("  (none)")

    print()
    print(" SUMMARY")
    print("-" * 60)
    print(f"  Rows checked:     {total}")
    print(f"  Format issues:    {total_format}")
    print(f"  Duplicate groups: {len(dupe_groups)} ({total_dupe_rows} rows)")
    print(f"  Total problems:   {total_format + total_dupe_rows}")
    print("=" * 60)

    return total_format > 0 or len(dupe_groups) > 0


HEADERS_EXPECTED = [
    "Company Name", "Website", "Phone", "Email", "Hall", "Stand", "Detail Page URL",
]


def main():
    if len(sys.argv) < 2:
        print("Usage: python checker.py <enriched.xlsx>")
        sys.exit(1)

    filepath = sys.argv[1]

    try:
        wb = load_workbook(filepath)
    except FileNotFoundError:
        print(f"Error: no such file — {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: cannot open Excel — {e}")
        sys.exit(1)

    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    missing = [h for h in HEADERS_EXPECTED if h not in headers]
    if missing:
        print(f"Error: missing columns — {', '.join(missing)}")
        wb.close()
        sys.exit(1)

    rows = []
    for r in range(2, ws.max_row + 1):
        row = {}
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            row[headers[c - 1]] = str(val).strip() if val is not None else ""
        if any(row.values()):
            rows.append(row)

    wb.close()

    issues = collect_format_issues(rows)
    dupe_groups = collect_duplicates(rows)

    has_problems = print_report(filepath, rows, issues, dupe_groups)

    sys.exit(1 if has_problems else 0)


if __name__ == "__main__":
    main()
