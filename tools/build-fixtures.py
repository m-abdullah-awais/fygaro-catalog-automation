"""
Fygaro Catalog Automation
Regenerates the committed test fixtures.

  tests/fixtures/snapshots.js  the raw DOM of each Fygaro screen the automation
                               drives, copied in so the locator tests can load
                               them without a server. Third party iframes are
                               stripped so the tests never reach the network.
  tests/fixtures/prices.json   every "Precio Total" in the real catalog with its
                               expected result, produced here by a reference
                               implementation that shares no code with
                               src/shared/price.js, so the test stays a real
                               cross check rather than a snapshot of itself.

Run from the project root:  python tools/build-fixtures.py
"""

import io
import json
import os
import re
import sys
import zipfile

SOURCE_DIR = os.path.join("temp", "html")
TARGET = os.path.join("tests", "fixtures", "snapshots.js")
SKIP = {"prompt.txt"}

WORKBOOK = os.path.join("docs", "Cat\u00e1logo de Productos y Servicios Fygaro.xlsx")
PRICES_TARGET = os.path.join("tests", "fixtures", "prices.json")
SHEET_PART = "xl/worksheets/sheet1.xml"
PRICE_COLUMN = "E"

HEADER = """/*
 * Fygaro Catalog Automation
 * Page snapshots captured from the live Fygaro app, used by the locator tests.
 * Generated from temp/html by tools/build-fixtures.py. Third party iframes are
 * stripped so the tests never touch the network.
 */
window.FYG_SNAPSHOTS = """

CELL = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)


def build_snapshots():
    if not os.path.isdir(SOURCE_DIR):
        print("No %s directory. Nothing to do." % SOURCE_DIR)
        return 1

    snapshots = {}
    for name in sorted(os.listdir(SOURCE_DIR)):
        if not name.endswith(".txt") or name in SKIP:
            continue
        with io.open(os.path.join(SOURCE_DIR, name), encoding="utf-8") as handle:
            html = handle.read().strip()
        html = re.sub(r"<iframe\b[^>]*>.*?</iframe>", "", html, flags=re.S | re.I)
        html = re.sub(r"<iframe\b[^>]*/?>", "", html, flags=re.I)
        snapshots[name[:-4]] = html

    if not snapshots:
        print("No snapshots found in %s." % SOURCE_DIR)
        return 1

    body = HEADER + json.dumps(snapshots, ensure_ascii=False, indent=2) + ";\n"
    with io.open(TARGET, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(body)

    print("Wrote %s with %d snapshots: %s" % (TARGET, len(snapshots), ", ".join(sorted(snapshots))))
    return 0


def read_price_column():
    """Pulls one column out of the worksheet without any xlsx library."""
    archive = zipfile.ZipFile(WORKBOOK)
    shared = [
        "".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
        for si in re.findall(
            r"<si>(.*?)</si>", archive.read("xl/sharedStrings.xml").decode("utf-8"), re.S
        )
    ]
    sheet = archive.read(SHEET_PART).decode("utf-8")

    out = []
    for row in re.finditer(r"<row([^>]*)>(.*?)</row>", sheet, re.S):
        number = int(re.search(r'r="(\d+)"', row.group(1)).group(1))
        if number < 2:
            continue
        for attrs, body in CELL.findall(row.group(2)):
            ref = re.search(r'r="([A-Z]+)\d+"', attrs)
            if not ref or ref.group(1) != PRICE_COLUMN:
                continue
            kind = re.search(r't="([^"]+)"', attrs)
            body = body or ""
            value = re.search(r"<v>(.*?)</v>", body, re.S)
            if kind and kind.group(1) == "s" and value:
                text = shared[int(value.group(1))]
            elif kind and kind.group(1) == "inlineStr":
                text = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
            else:
                text = value.group(1) if value else ""
            text = text.strip()
            if text:
                out.append((number, text))
    return out


def reference_parse(raw):
    """
    Reference implementation, deliberately written differently from price.js.

    Every price in this catalog ends in exactly two digits after its final
    separator, so that separator is the decimal point and any earlier ones are
    thousands groupings. A value that breaks the assumption is reported rather
    than guessed at, so a new format cannot slip through as a wrong number.
    """
    digits = re.sub(r"^[^0-9]*", "", raw)
    if not digits or not re.match(r"^[0-9.,]+$", digits):
        return None
    marks = [i for i, ch in enumerate(digits) if ch in ".,"]
    if not marks:
        return None
    last = marks[-1]
    whole = re.sub(r"[.,]", "", digits[:last])
    frac = digits[last + 1:]
    if len(frac) != 2 or not whole.isdigit() or not frac.isdigit():
        return None
    return "%s.%s" % (whole, frac)


def build_prices():
    if not os.path.isfile(WORKBOOK):
        print("No workbook at %s. Skipping prices.json." % WORKBOOK)
        return 0

    cases = []
    rejected = []
    for number, raw in read_price_column():
        text = reference_parse(raw)
        if text is None:
            rejected.append((number, raw))
            continue
        cases.append({"row": number, "raw": raw, "value": float(text), "text": text})

    if rejected:
        print("Refusing to write %s: %d values do not fit the reference rule."
              % (PRICES_TARGET, len(rejected)))
        for number, raw in rejected[:10]:
            print("   row %d: %r" % (number, raw))
        return 1

    lines = [json.dumps(c, ensure_ascii=False) for c in cases]
    body = "[\n" + ",\n".join(lines) + "\n]\n"
    with io.open(PRICES_TARGET, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(body)
    print("Wrote %s with %d prices." % (PRICES_TARGET, len(cases)))
    return 0


def main():
    failed = build_snapshots()
    if failed:
        return failed
    return build_prices()


if __name__ == "__main__":
    sys.exit(main())
