#!/usr/bin/env python3
"""
Generate monthly Stundennachweis (timesheet) Excel files.

Reads worker assignments from current_data.xlsx and fills empty_template.xlsx
for each row, producing one semi-filled template per assignment.

Usage: python3 scripts/generate_templates.py
"""

import os
import sys
import calendar
import zipfile
import io
import re
from datetime import date
from html import escape as html_escape

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
INPUT_DIR = os.path.join(PROJECT_DIR, 'data', 'input')
OUTPUT_DIR = os.path.join(PROJECT_DIR, 'data', 'output')
TEMPLATE_PATH = os.path.join(INPUT_DIR, 'empty_template.xlsx')
DATA_PATH = os.path.join(INPUT_DIR, 'current_data.xlsx')

GERMAN_MONTHS = {
    1: 'Jänner', 2: 'Februar', 3: 'März', 4: 'April',
    5: 'Mai', 6: 'Juni', 7: 'Juli', 8: 'August',
    9: 'September', 10: 'Oktober', 11: 'November', 12: 'Dezember',
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def xml_esc(text):
    """Escape text for safe embedding in XML element content."""
    return html_escape(str(text), quote=True)


def date_to_serial(d):
    """Python date -> Excel serial date number."""
    return (d - date(1899, 12, 30)).days


# ---------------------------------------------------------------------------
# Interactive prompts
# ---------------------------------------------------------------------------


def prompt_month_year():
    while True:
        try:
            year = int(input("Year (e.g. 2026): "))
            if 2020 <= year <= 2100:
                break
            print("  Enter a year between 2020 and 2100.")
        except ValueError:
            print("  Invalid input.")

    while True:
        try:
            month = int(input("Month (1-12): "))
            if 1 <= month <= 12:
                break
            print("  Enter 1-12.")
        except ValueError:
            print("  Invalid input.")

    print(f"\n-> {GERMAN_MONTHS[month]} {year}")
    return year, month


def prompt_holidays(year, month):
    num_days = calendar.monthrange(year, month)[1]
    holidays = []
    print(f"\nEnter national holiday dates for {GERMAN_MONTHS[month]} {year} (1-{num_days}).")
    print("Type the day number + Enter for each holiday.")
    print("Type F + Enter when done.\n")

    while True:
        raw = input("Holiday day (or F to finish): ").strip()
        if raw.upper() == 'F':
            break
        try:
            day = int(raw)
            if not 1 <= day <= num_days:
                print(f"  Must be 1-{num_days}.")
            elif day in holidays:
                print("  Already in list.")
            else:
                holidays.append(day)
                wd_name = calendar.day_name[date(year, month, day).weekday()]
                print(f"  + {day}. {GERMAN_MONTHS[month]} ({wd_name})")
        except ValueError:
            print("  Enter a number or F.")

    holidays.sort()
    print(f"\nHolidays: {holidays if holidays else '(none)'}")

    while True:
        choice = input("Confirm (Y) or re-enter (E)? ").strip().upper()
        if choice == 'Y':
            return holidays
        if choice == 'E':
            return prompt_holidays(year, month)
        print("  Enter Y or E.")


def prompt_contacts(assignments):
    """Prompt for 'Auftragsabwicklung mit' per unique project name."""
    projects = sorted(set(a['project_name'] for a in assignments))
    print(f"\n{len(projects)} unique projects found. Enter contact person for each.\n")
    contacts = {}
    for i, proj in enumerate(projects, 1):
        name = input(f"  [{i}/{len(projects)}] {proj}\n    Auftragsabwicklung mit: ").strip()
        contacts[proj] = name
    print()
    return contacts


def compute_weekdays(year, month, holidays):
    """All weekday dates in *month*, excluding *holidays*."""
    num_days = calendar.monthrange(year, month)[1]
    return [
        date(year, month, d)
        for d in range(1, num_days + 1)
        if date(year, month, d).weekday() < 5 and d not in holidays
    ]


# ---------------------------------------------------------------------------
# Read current_data.xlsx (openpyxl – read only, never saved back)
# ---------------------------------------------------------------------------


def read_current_data(data_path=None):
    import openpyxl

    wb = openpyxl.load_workbook(data_path or DATA_PATH, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        po, proj, rate, resource = r
        if po is None:
            continue
        rows.append({
            'purchase_order': po,
            'project_name': str(proj).strip() if proj else '',
            'rate_name': str(rate).strip() if rate else '',
            'resource': str(resource).strip() if resource else '',
        })
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# XML manipulation (direct string ops – preserves logo, shapes, styles)
# ---------------------------------------------------------------------------


def _build_data_row(n, resource, serial, project, rate):
    return (
        f'<row r="{n}" spans="3:8" x14ac:dyDescent="0.35">'
        f'<c r="C{n}" s="5" t="inlineStr"><is><t>{xml_esc(resource)}</t></is></c>'
        f'<c r="D{n}" s="19"><v>{serial}</v></c>'
        f'<c r="E{n}" s="2"/>'
        f'<c r="F{n}" s="3" t="inlineStr"><is><t>{xml_esc(project)}</t></is></c>'
        f'<c r="G{n}" s="3" t="inlineStr"><is><t>{xml_esc(rate)}</t></is></c>'
        f'<c r="H{n}" s="3"/>'
        f'</row>'
    )


def _build_sum_row(n, last_data):
    return (
        f'<row r="{n}" spans="3:8" x14ac:dyDescent="0.35">'
        f'<c r="C{n}" s="17" t="s"><v>14</v></c>'
        f'<c r="D{n}" s="18"/>'
        f'<c r="E{n}" s="16"><f>SUM(E9:E{last_data})</f><v>0</v></c>'
        f'<c r="F{n}" s="15"/>'
        f'<c r="G{n}" s="15"/>'
        f'<c r="H{n}" s="15"/>'
        f'</row>'
    )


def _build_footer_row(n):
    return (
        f'<row r="{n}" spans="3:8" x14ac:dyDescent="0.35">'
        f'<c r="C{n}" s="7" t="s"><v>9</v></c>'
        f'<c r="D{n}" s="4"/>'
        f'<c r="G{n}" s="7" t="s"><v>15</v></c>'
        f'<c r="H{n}" s="4"/>'
        f'</row>'
    )


def modify_sheet(xml_bytes, assignment, year, month, weekdays, contact):
    xml = xml_bytes.decode('utf-8')

    po = assignment['purchase_order']
    project = assignment['project_name']
    rate = assignment['rate_name']
    resource = assignment['resource']
    month_name = GERMAN_MONTHS[month]

    # --- header cells ---
    xml = xml.replace(
        '<c r="H5" s="6"/>',
        f'<c r="H5" s="6"><v>{po}</v></c>',
    )
    xml = xml.replace(
        '<c r="H6" s="6"/>',
        f'<c r="H6" s="6" t="inlineStr"><is><t>{xml_esc(contact)}</t></is></c>',
    )
    xml = xml.replace(
        '<c r="D7" s="6"/>',
        f'<c r="D7" s="6" t="inlineStr"><is><t>{xml_esc(month_name)}</t></is></c>',
    )
    xml = xml.replace(
        '<c r="E7" s="6"/>',
        f'<c r="E7" s="6"><v>{year}</v></c>',
    )
    xml = xml.replace(
        '<c r="H7" s="6"/>',
        f'<c r="H7" s="6" t="inlineStr"><is><t>{xml_esc(project)}</t></is></c>',
    )

    # --- replace data rows 9-32 + footer rows 33-34 ---
    n = len(weekdays)
    parts = []
    for i, day in enumerate(weekdays):
        parts.append(
            _build_data_row(9 + i, resource, date_to_serial(day), project, rate)
        )
    last_data = 8 + n
    parts.append(_build_sum_row(last_data + 1, last_data))
    parts.append(_build_footer_row(last_data + 2))
    replacement = ''.join(parts)

    # Match the block from <row r="9" ...> through </row> of row 34
    xml = re.sub(
        r'<row r="9".*<row r="34"[^>]*>.*?</row>',
        replacement,
        xml,
        flags=re.DOTALL,
    )

    # --- update dimension ---
    xml = re.sub(
        r'<dimension ref="B1:H\d+"/>',
        f'<dimension ref="B1:H{last_data + 2}"/>',
        xml,
    )

    return xml.encode('utf-8')


def modify_workbook(xml_bytes, sheet_name):
    xml = xml_bytes.decode('utf-8')
    xml = xml.replace('name="moritz.luibrand"', f'name="{xml_esc(sheet_name)}"')
    return xml.encode('utf-8')


def remove_calcchain_ref(content_types_bytes):
    """Remove calcChain override from [Content_Types].xml."""
    xml = content_types_bytes.decode('utf-8')
    xml = re.sub(
        r'<Override[^>]*calcChain[^>]*/>', '', xml
    )
    return xml.encode('utf-8')


# ---------------------------------------------------------------------------
# Generate one output file
# ---------------------------------------------------------------------------


def generate_file(template_bytes, assignment, year, month, weekdays, contact,
                   output_path):
    # Read all entries from the template ZIP
    files = {}
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zf:
        for name in zf.namelist():
            files[name] = zf.read(name)

    # Modify sheet & workbook XML
    files['xl/worksheets/sheet1.xml'] = modify_sheet(
        files['xl/worksheets/sheet1.xml'], assignment, year, month, weekdays,
        contact
    )
    files['xl/workbook.xml'] = modify_workbook(
        files['xl/workbook.xml'], assignment['resource']
    )

    # Remove calcChain (Excel rebuilds on open) and its Content_Types entry
    if 'xl/calcChain.xml' in files:
        del files['xl/calcChain.xml']
        files['[Content_Types].xml'] = remove_calcchain_ref(
            files['[Content_Types].xml']
        )

    # Write output
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    for path, label in [(TEMPLATE_PATH, 'empty_template'), (DATA_PATH, 'current_data')]:
        if not os.path.exists(path):
            print(f"Error: {label} not found at {path}")
            sys.exit(1)

    year, month = prompt_month_year()
    holidays = prompt_holidays(year, month)
    weekdays = compute_weekdays(year, month, holidays)
    print(f"\n{len(weekdays)} working days in {GERMAN_MONTHS[month]} {year}:")
    for wd in weekdays:
        print(f"  {wd.day:2d}. {GERMAN_MONTHS[month]} ({calendar.day_abbr[wd.weekday()]})")
    print()

    assignments = read_current_data()
    print(f"{len(assignments)} rows in current_data.xlsx")

    contacts = prompt_contacts(assignments)

    out_dir = os.path.join(OUTPUT_DIR, f"{year}_{month:02d}")
    os.makedirs(out_dir, exist_ok=True)

    with open(TEMPLATE_PATH, 'rb') as f:
        template_bytes = f.read()

    print(f"Generating in {out_dir}/ ...")
    seen = {}
    for i, a in enumerate(assignments, 1):
        safe_proj = a['project_name'].replace('/', '_').replace('\\', '_')
        proj_dir = os.path.join(out_dir, safe_proj)
        os.makedirs(proj_dir, exist_ok=True)
        safe_name = a['resource'].replace('/', '_').replace('\\', '_')
        base = f"{safe_name}_{a['purchase_order']}_{year}_{month:02d}"
        seen[base] = seen.get(base, 0) + 1
        filename = f"{base}.xlsx" if seen[base] == 1 else f"{base}_{seen[base]}.xlsx"
        contact = contacts.get(a['project_name'], '')
        generate_file(template_bytes, a, year, month, weekdays, contact,
                      os.path.join(proj_dir, filename))
        print(f"  [{i}/{len(assignments)}] {safe_proj}/{filename}")

    print(f"\nDone! {len(assignments)} files in {out_dir}/")


if __name__ == '__main__':
    main()
