"""
GCTS Expiry Letter Generator
-----------------------------
Run in VS Code (or any terminal):
    pip install openpyxl python-docx
    python generate_letters.py GCTS.xlsx

Everything you'd normally edit is right below in this file — no separate
CSV to keep track of. Scroll down to the "EDIT THIS SECTION" block:

- BRAND_CONTACTS: which contact owns each brand (and, optionally, a
  specific product type within a brand — e.g. GENERALCO hobs go to one
  contact, GENERALCO cooking ranges to another).
- CONTACT_EMAILS: each contact's email address(es).

Just edit those two Python structures directly and rerun the script.

Output: one combined .docx, one letter per contact, each starting with a
plain-text "To: email1; email2" line (semicolon-separated — select and
paste straight into Outlook's To field), then the same "Dear [Contact]"
letter format and table as before.

Optional: to pull emails from Certificate_Master.xlsx (Sheet3) instead
of typing them by hand, run:
    python generate_letters.py GCTS.xlsx --suggest-emails Certificate_Master.xlsx
This does NOT edit the script for you — it prints ready-to-paste Python
dict lines for any contact found in Sheet3 that isn't already in
CONTACT_EMAILS below, so you can copy them in yourself.
"""

import argparse
import os
import re
import sys

import openpyxl
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================================
#  EDIT THIS SECTION — your data lives here, nothing else to download/manage
# ============================================================================

# Which contact owns each brand.
# - (brand, contact, "") = brand-wide default: every row of that brand
#   goes to that contact.
# - (brand, contact, "some product") = scoped override: only rows of
#   that exact brand AND that exact Product Name go to that contact
#   (case-insensitive match). Use this when different product types
#   under the same brand are handled by different people.
BRAND_CONTACTS = [
    # (brand,            contact,               product override)
    ("AEG",              "Aru",                 ""),
    ("Electrolux",       "Aru",                 ""),
    ("FRIGIDAIRE",       "Aru",                 ""),
    ("ORYX",             "Gonca",               ""),
    ("Aestörn",          "Mamur",               ""),
    ("Bompani",          "Mamur",               ""),
    ("Fiorenta",         "Mamur",               ""),
    ("THOMSON",          "Mamur",               ""),
    ("GENERALCO",        "Mamur",               "Vitroceramic Hob"),
    ("GENERALCO",        "Sengun",              "Free Standing Cooking Range"),
    ("Ferre",            "Sengun",              ""),
    ("Setech",           "Sengun",              ""),
    ("SUPER GENERAL",    "Sengun",              ""),
    ("ALM",              "Sengun",              ""),
    ("HARMONY",          "Alex",                ""),
    ("Kelvinator",       "Alex",                ""),
    ("Premier",          "Alex",                ""),
    ("Franke",           "Kushal",              ""),
    ("ROBAM",            "Joy",                 ""),
    ("gorenje",          "Richard and Nabil",   ""),
    ("FULGOR MILANO",    "Sanipex Group",       ""),
]

# Each contact's email address(es).
CONTACT_EMAILS = {
    "Aru": ["aru.singhal@electrolux.com", "Nived.Viswanathan@uae.tuv.com"],
    "Mamur": ["eray.gercek@mamurtech.com", "omer.yanik@mamurtech.com",
              "atahan.kizilkaya@mamurtech.com", "mustafa.boztas@ferre.com.tr"],
    "Gonca": ["gonca.izci@ferre.com.tr", "mustafa.boztas@ferre.com.tr"],
    "Sengun": ["aysegul.sengun@ferre.com.tr"],
    "Alex": ["volansky@premierrange.com"],
    "Kushal": ["Kushal.Adhia@franke.com"],
    "Joy": ["Joy.Xu@tuv.com", "KellyC.Feng@tuv.com"],
    "Richard and Nabil": ["certification.uae@hisense.com", "nabil.abdulla@partners.hisense.com"],
    "Sanipex Group": ["vidya.raghavan@sanipexgroup.com", "aswin@sanipexgroup.com"],
}

# ============================================================================
#  End of editable section — everything below is just plumbing
# ============================================================================

SHEET_NAME = "All 35 GCTS Lookup"
HEADER_ROW = 1
FIRST_DATA_ROW = 2

# Columns are found in two passes against the sheet's HEADER_ROW text
# (case-insensitive, whitespace-normalized), not by fixed position:
#   1. Exact match against a field's *_EXACT set first -- this is what
#      resolves sheets with several similarly-named columns correctly
#      (e.g. both "GCTS No" and "Product GCTS" contain "gcts", but only
#      "GCTS No" is an exact match, so it wins outright; likewise
#      "Product Description" beats "Product GCTS"/"Product Category" for
#      the product field).
#   2. Only for any field still unresolved, a loose substring match
#      against *_KEYWORDS, for sheets that don't use any of the exact
#      phrasings below.
# Add more entries to either list if a workbook still isn't recognized.
GCTS_EXACT = {"gcts", "gcts no", "gcts no.", "gcts number", "gcts certificate no"}
PRODUCT_EXACT = {"product name", "product description", "product"}
BRAND_EXACT = {"brand", "brand name", "trademark", "trademark/brand"}
MODEL_EXACT = {"model no.", "model no", "model number", "model"}

GCTS_KEYWORDS = ["gcts"]
PRODUCT_KEYWORDS = ["product"]
BRAND_KEYWORDS = ["brand"]
MODEL_KEYWORDS = ["model"]
EXCLUDED_FILL_COLORS = {"FFFFFF00", "FF00B050"}  # yellow, green — rows to skip

CONTACT_SHEET = "Sheet3"
CONTACT_FIRST_ROW = 2
CONTACT_NAME_COL = 1
CONTACT_EMAIL_START_COL = 3

EXPIRED_STANDARD_LINE1 = "IEC 60335-2-6:2014+AMD1:2018"
EXPIRED_STANDARD_LINE2 = "Expired on 23 Aug 2026"
STANDARD_LINK = "https://www.gso.org.sa/ns/hs/BD-142004-01"
SIGNOFF_NAME = "Abhishek Syamkumar"
SIGNOFF_TITLE = "Senior Project Engineer-P.03/MAS"

GCTS_BLUE = RGBColor(0x1F, 0x4E, 0x79)
STANDARD_OLIVE = RGBColor(0x8C, 0x8C, 0x00)
EXPIRED_RED = RGBColor(0xFF, 0x00, 0x00)
BORDER_COLOR = "8EA9C1"
COL_WIDTHS_CM = [2.3, 3.9, 2.3, 5.6, 3.3]


# ------------------------------------------------------------- extraction --

def _normalize_header(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def find_columns(ws, header_row=HEADER_ROW):
    """Scan the header row and return {'gcts': col, 'product': col,
    'brand': col, 'model': col}. Exact-phrase matches are resolved first
    (so sheets with several similarly-worded columns pick the right one),
    then any field still unresolved falls back to a loose substring match.
    Not by fixed position."""
    exact = {
        "gcts": GCTS_EXACT, "product": PRODUCT_EXACT,
        "brand": BRAND_EXACT, "model": MODEL_EXACT,
    }
    keywords = {
        "gcts": GCTS_KEYWORDS, "product": PRODUCT_KEYWORDS,
        "brand": BRAND_KEYWORDS, "model": MODEL_KEYWORDS,
    }
    found = {}
    all_headers = []
    normalized = []
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(row=header_row, column=col).value
        all_headers.append(raw)
        normalized.append(_normalize_header(raw))

    # Pass 1: exact match, in column order, first hit per field wins.
    for col, text in enumerate(normalized, start=1):
        if not text:
            continue
        for field, phrases in exact.items():
            if field not in found and text in phrases:
                found[field] = col

    # Pass 2: loose substring match, only for fields exact-matching missed.
    for col, text in enumerate(normalized, start=1):
        if not text:
            continue
        for field, kws in keywords.items():
            if field not in found and any(kw in text for kw in kws):
                found[field] = col

    missing = [f for f in exact if f not in found]
    if missing:
        raise KeyError(
            f"Could not find a column for: {', '.join(missing)} in row {header_row} "
            f"of sheet '{ws.title}'. Headers found: {all_headers}. "
            f"If this sheet spells a header differently, add it to the matching "
            f"*_EXACT set or *_KEYWORDS list near the top of this file."
        )
    return found


def load_rows(xlsx_path):
    """Return list of dicts for every non-highlighted row in the GCTS sheet."""
    if not os.path.isfile(xlsx_path):
        sys.exit(f"ERROR: file not found: {xlsx_path}")
    if not xlsx_path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        sys.exit(f"ERROR: '{xlsx_path}' isn't a format openpyxl can read "
                  "(.xlsx/.xlsm/.xltx/.xltm only — convert old .xls first).")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    cols = find_columns(ws)

    rows = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        gcts_cell = ws.cell(row=r, column=cols["gcts"])
        if gcts_cell.value is None:
            continue
        fill = gcts_cell.fill.fgColor.rgb if gcts_cell.fill and gcts_cell.fill.fgColor else None
        if fill in EXCLUDED_FILL_COLORS:
            continue
        rows.append({
            "gcts": str(gcts_cell.value).strip(),
            "model": str(ws.cell(row=r, column=cols["model"]).value or "").strip(),
            "brand": str(ws.cell(row=r, column=cols["brand"]).value or "").strip(),
            "product": str(ws.cell(row=r, column=cols["product"]).value or "").strip(),
        })
    return rows


def build_lookup_dicts(brand_contacts=None):
    """Turn a BRAND_CONTACTS-shaped list into (brand_defaults, product_overrides) dicts.
    Defaults to the module-level BRAND_CONTACTS for plain CLI use; the web app
    passes in the edited-in-browser table instead."""
    if brand_contacts is None:
        brand_contacts = BRAND_CONTACTS
    brand_defaults = {}
    product_overrides = {}  # (brand_lower, product_lower) -> contact
    for brand, contact, product in brand_contacts:
        if product:
            product_overrides[(brand.lower(), product.lower())] = contact
        else:
            brand_defaults[brand] = contact
    return brand_defaults, product_overrides


def group_by_contact(rows, brand_defaults, product_overrides):
    groups, order, unmapped = {}, [], []
    for row in rows:
        key = (row["brand"].lower(), row["product"].lower())
        contact = product_overrides.get(key) or brand_defaults.get(row["brand"])
        if not contact:
            unmapped.append(row)
            continue
        groups.setdefault(contact, [])
        if contact not in order:
            order.append(contact)
        groups[contact].append(row)
    result = [(c, groups[c]) for c in order]
    if unmapped:
        result.append(("(brand/product not in BRAND_CONTACTS)", unmapped))
    return result


def suggest_emails_from_cert_master(cert_master_path):
    """Print copy-pasteable CONTACT_EMAILS lines for contacts in
    Certificate_Master's Sheet3 that aren't already in CONTACT_EMAILS."""
    wb = openpyxl.load_workbook(cert_master_path, data_only=True)
    if CONTACT_SHEET not in wb.sheetnames:
        sys.exit(f"ERROR: '{CONTACT_SHEET}' not found in {cert_master_path}. "
                  f"Sheets present: {wb.sheetnames}")
    ws = wb[CONTACT_SHEET]

    new_lines = []
    for r in range(CONTACT_FIRST_ROW, ws.max_row + 1):
        name = ws.cell(row=r, column=CONTACT_NAME_COL).value
        if not name:
            continue
        name = str(name).strip()
        if name in CONTACT_EMAILS:
            continue  # already have it — don't suggest overwriting a manual edit
        emails = []
        c = CONTACT_EMAIL_START_COL
        while True:
            val = ws.cell(row=r, column=c).value
            if val:
                emails.append(str(val).strip())
            elif c > CONTACT_EMAIL_START_COL + 6:
                break
            c += 1
        if emails:
            email_list = ", ".join(f'"{e}"' for e in emails)
            new_lines.append(f'    "{name}": [{email_list}],')

    if not new_lines:
        print("No new contacts found in Certificate_Master that aren't already "
              "in CONTACT_EMAILS.")
        return

    print(f"\nFound {len(new_lines)} contact(s) in {cert_master_path} not yet in "
          f"CONTACT_EMAILS. Paste these lines into the CONTACT_EMAILS dict in this "
          f"script:\n")
    for line in new_lines:
        print(line)
    print()


# ------------------------------------------------------------- docx build --

def set_cell_background(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_borders(cell, color=BORDER_COLOR, size=4):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def write_cell(cell, runs, size=10):
    cell.text = ""
    items = runs if isinstance(runs, list) else [(runs, False, None)]
    first = True
    for text, bold, color in items:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.name = "Arial"
        run.bold = bold
        if color:
            run.font.color.rgb = color
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_borders(cell)


def add_table(doc, rows):
    table = doc.add_table(rows=1 + len(rows), cols=5)
    table.autofit = False

    headers = ["GCTS", "Model No.", "Brand", "Product Name", "Expired Standard"]
    for c, text in enumerate(headers):
        cell = table.rows[0].cells[c]
        write_cell(cell, [(text, True, RGBColor(0xFF, 0xFF, 0xFF))])
        set_cell_background(cell, "1F4E79")

    for i, row in enumerate(rows, start=1):
        cells = table.rows[i].cells
        write_cell(cells[0], [(row["gcts"], True, GCTS_BLUE)])
        write_cell(cells[1], row["model"])
        write_cell(cells[2], row["brand"])
        write_cell(cells[3], row["product"])
        if i == 1:
            write_cell(cells[4], [
                (EXPIRED_STANDARD_LINE1, False, STANDARD_OLIVE),
                (EXPIRED_STANDARD_LINE2, True, EXPIRED_RED),
            ])
        else:
            write_cell(cells[4], "")

    if len(rows) > 1:
        table.rows[1].cells[4].merge(table.rows[len(rows)].cells[4])

    for col_idx, width_cm in enumerate(COL_WIDTHS_CM):
        for row in table.rows:
            row.cells[col_idx].width = Cm(width_cm)


def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(underline)
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Calibri")
    rPr.append(rFonts)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_body_paragraph(doc, text, size=11, space_after=10, bold=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run.bold = bold
    if color:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    return p


def build_letter(doc, contact, rows, emails, is_first):
    if not is_first:
        pb = doc.add_paragraph()
        pb.paragraph_format.page_break_before = True

    # Outlook's To field expects semicolon-separated addresses, plain text
    # (no hyperlink styling) so the line can be selected and pasted straight in.
    if emails:
        add_body_paragraph(doc, "To: " + "; ".join(emails), size=10, space_after=4,
                            color=RGBColor(0x44, 0x44, 0x44))
    else:
        add_body_paragraph(doc, "To: (no email on file for this contact)", size=10,
                            space_after=4, color=RGBColor(0x99, 0x33, 0x00))

    add_body_paragraph(doc, f"Dear {contact}", size=11, space_after=14)
    add_body_paragraph(
        doc,
        "Please be advised that an additional set of certificates listed below has been "
        "suspended due to an expired standard."
    )
    add_body_paragraph(
        doc,
        "To avoid termination and to maintain compliance, please provide the reports "
        "updated to the latest standards ASAP",
        space_after=14,
    )

    add_table(doc, rows)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_body_paragraph(doc, "Please let us know if you need any clarification or further details.")
    add_body_paragraph(doc, "Note for latest standard, please refer to the link below:", space_after=4)

    link_p = doc.add_paragraph()
    link_p.paragraph_format.space_after = Pt(14)
    add_hyperlink(link_p, STANDARD_LINK, STANDARD_LINK)

    add_body_paragraph(doc, "Best Regards,", space_after=2)
    p = doc.add_paragraph()
    run = p.add_run(SIGNOFF_NAME)
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    p.paragraph_format.space_after = Pt(0)

    p2 = doc.add_paragraph()
    run2 = p2.add_run(SIGNOFF_TITLE)
    run2.bold = True
    run2.font.size = Pt(11)
    run2.font.name = "Calibri"


def build(rows, brand_defaults, product_overrides, output_path, contact_emails=None):
    """contact_emails defaults to the module-level CONTACT_EMAILS for plain
    CLI use; the web app passes in the edited-in-browser table instead."""
    if contact_emails is None:
        contact_emails = CONTACT_EMAILS
    groups = group_by_contact(rows, brand_defaults, product_overrides)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for i, (contact, contact_rows) in enumerate(groups):
        emails = contact_emails.get(contact, [])
        build_letter(doc, contact, contact_rows, emails, is_first=(i == 0))

    doc.save(output_path)
    return groups


# ------------------------------------------------------------------ main --

def main():
    parser = argparse.ArgumentParser(description="Generate GCTS expiry letters, grouped by contact.")
    parser.add_argument("gcts_xlsx", nargs="?", help="Path to the GCTS Lookup Excel file")
    parser.add_argument("-o", "--output", default=None,
                         help="Output .docx path (default: <input filename>_filled.docx)")
    parser.add_argument("--suggest-emails", metavar="Certificate_Master.xlsx",
                         help="Print copy-pasteable CONTACT_EMAILS lines for any contact in "
                              "this file's Sheet3 that isn't already in the script, then exit")
    args = parser.parse_args()

    if args.suggest_emails:
        suggest_emails_from_cert_master(args.suggest_emails)
        return

    gcts_xlsx = args.gcts_xlsx
    if not gcts_xlsx:
        # No command-line args (e.g. running via VS Code's Run/Debug button) —
        # ask interactively instead of exiting.
        gcts_xlsx = input("Path to the GCTS Lookup .xlsx file: ").strip().strip('"')
        if not gcts_xlsx:
            sys.exit("No file given — exiting.")

    rows = load_rows(gcts_xlsx)
    brand_defaults, product_overrides = build_lookup_dicts()

    output_path = args.output
    if not output_path:
        stem = os.path.splitext(os.path.basename(gcts_xlsx))[0]
        output_path = f"{stem}_filled.docx"

    groups = build(rows, brand_defaults, product_overrides, output_path, CONTACT_EMAILS)

    print(f"\nWrote {output_path}: {len(groups)} letters covering {len(rows)} rows.")
    for contact, contact_rows in groups:
        emails = CONTACT_EMAILS.get(contact, [])
        flag = "" if emails else "  <-- NO EMAIL FOUND (add to CONTACT_EMAILS)"
        print(f"  {contact}: {len(contact_rows)} row(s){flag}")


if __name__ == "__main__":
    main()
    