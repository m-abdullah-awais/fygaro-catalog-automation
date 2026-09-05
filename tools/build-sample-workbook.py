"""
Fygaro Catalog Automation
Builds tests/fixtures/catalog-sample.xlsx from the real catalog.

The real workbook is 63.5 MB, and headless Chrome does not finish fetching a
file that size before --dump-dom gives up, so the browser suites cannot use it.
This trims it to 30 data rows and three tiny images while keeping every
structural feature the code has to cope with:

  * the sheet name "Logros " with its trailing space
  * headers ending in "Columna 1", and no Link column anywhere
  * a table with an autoFilter, and a _FilterDatabase defined name
  * no <dimension> element, because this is a Google Sheets export
  * pictures anchored in column H through a drawing part
  * a row carrying empty I through X cells, as row 1017 does in the real file

It also adds three things the real catalog happens not to contain, so the
reader is tested on the format rather than only on what one file looked like on
one day: a row with two pictures, a twoCellAnchor, and an absoluteAnchor.

Run from the project root:  python tools/build-sample-workbook.py
"""

import io
import os
import re
import sys
import zipfile

SOURCE = os.path.join("docs", "Catálogo de Productos y Servicios Fygaro.xlsx")
TARGET = os.path.join("tests", "fixtures", "catalog-sample.xlsx")

SHEET = "xl/worksheets/sheet1.xml"
DRAWING = "xl/drawings/drawing1.xml"
DRAWING_RELS = "xl/drawings/_rels/drawing1.xml.rels"
TABLE = "xl/tables/table1.xml"
WORKBOOK = "xl/workbook.xml"

LAST_ROW = 31
QUIRK_ROW = 17

# A 1x1 PNG and a minimal JPEG, both real files a browser accepts. Written by
# hand so the fixture stays under a kilobyte of image data rather than 65 MB.
PNG_BYTES = bytes([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
    0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
    0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
])

JPEG_BYTES = (
    bytes([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00,
           0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00])
    + bytes([0xFF, 0xDB, 0x00, 0x43, 0x00]) + bytes([0x08] * 64)
    + bytes([0xFF, 0xC9, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01,
             0x11, 0x00])
    + bytes([0xFF, 0xCC, 0x00, 0x06, 0x00, 0x10, 0x10, 0x05])
    + bytes([0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00, 0xD2,
             0xCF, 0x20, 0xFF, 0xD9])
)

# image3 is deliberately JPEG bytes behind a .png name, so the reader has to
# sniff the format rather than trust the extension. The real catalog carries the
# milder version of this: one genuine .jpg among 34 .png files.
MEDIA = {
    "xl/media/image1.png": PNG_BYTES,
    "xl/media/image2.jpg": JPEG_BYTES,
    "xl/media/image3.png": JPEG_BYTES,
}

XDR_NS = (
    'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def picture(rel_id, name):
    return (
        "<xdr:pic><xdr:nvPicPr>"
        '<xdr:cNvPr id="0" name="%s"/><xdr:cNvPicPr preferRelativeResize="0"/>'
        "</xdr:nvPicPr><xdr:blipFill>"
        '<a:blip cstate="print" r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
        "</xdr:blipFill>"
        "<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>"
        "</xdr:pic>" % (name, rel_id)
    )


def one_cell(rel_id, row_zero_based, name):
    """A picture pinned to one cell, which is what the real catalog uses."""
    return (
        "<xdr:oneCellAnchor><xdr:from>"
        "<xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff>"
        "<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff>"
        "</xdr:from><xdr:ext cx=\"1285875\" cy=\"981075\"/>%s"
        '<xdr:clientData fLocksWithSheet="0"/></xdr:oneCellAnchor>'
        % (row_zero_based, picture(rel_id, name))
    )


def two_cell(rel_id, row_zero_based, name):
    """A picture stretched between two cells. Its "from" row is the one it belongs to."""
    return (
        "<xdr:twoCellAnchor><xdr:from>"
        "<xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff>"
        "<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        "<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>"
        "<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>%s"
        '<xdr:clientData fLocksWithSheet="0"/></xdr:twoCellAnchor>'
        % (row_zero_based, row_zero_based + 1, picture(rel_id, name))
    )


def absolute(rel_id, name):
    """Floating, pinned to no row, so it belongs to no product."""
    return (
        '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="100" cy="100"/>%s'
        "<xdr:clientData/></xdr:absoluteAnchor>" % picture(rel_id, name)
    )


def build_drawing():
    anchors = [
        one_cell("rId1", 1, "image1.png"),    # sheet row 2
        one_cell("rId1", 2, "image1.png"),    # sheet row 3, the same picture reused
        one_cell("rId2", 3, "image2.jpg"),    # sheet row 4
        one_cell("rId2", 4, "image2.jpg"),    # sheet row 5, two pictures
        one_cell("rId3", 4, "image3.png"),    # sheet row 5
        two_cell("rId1", 5, "image1.png"),    # sheet row 6
        absolute("rId1", "floating.png"),
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        "<xdr:wsDr %s>%s</xdr:wsDr>" % (XDR_NS, "".join(anchors))
    )


def build_drawing_rels():
    rels = "".join(
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/image" Target="../media/%s"/>'
        % (n + 1, os.path.basename(path))
        for n, path in enumerate(sorted(MEDIA))
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "%s</Relationships>" % rels
    )


def trim_rows(xml, last_row, extra_for_row=None):
    """Drops every <row> past last_row, leaving the rest byte identical."""
    out = []
    cursor = 0
    for match in re.finditer(r"<row\b[^>]*>", xml):
        tag = match.group(0)
        number = re.search(r'\br="(\d+)"', tag)
        if not number:
            continue
        if tag.endswith("/>"):
            end = match.end()
        else:
            close = xml.index("</row>", match.end())
            end = close + len("</row>")
        keep = int(number.group(1)) <= last_row
        out.append(xml[cursor:match.start()])
        if keep:
            body = xml[match.start():end]
            if extra_for_row and int(number.group(1)) == QUIRK_ROW:
                body = extra_for_row(body)
            out.append(body)
        cursor = end
    out.append(xml[cursor:])
    return "".join(out)


def add_trailing_cells(row_markup):
    """Reproduces the real sheet's row 1017, which already carries empty I..X cells."""
    if row_markup.endswith("/>"):
        return row_markup
    cells = "".join(
        '<c r="%s%d" s="65"/>' % (chr(ord("I") + n), QUIRK_ROW) for n in range(16)
    )
    return row_markup[: -len("</row>")] + cells + "</row>"


def main():
    if not os.path.isfile(SOURCE):
        print("No workbook at %s. Nothing to build." % SOURCE)
        return 1

    source = zipfile.ZipFile(SOURCE)
    parts = {}
    order = []

    for name in source.namelist():
        if name.startswith("xl/media/"):
            continue
        order.append(name)
        parts[name] = source.read(name)

    # The target sheet keeps its header and 30 data rows, plus the trailing cell
    # quirk that the real file has on row 1017.
    sheet_xml = parts[SHEET].decode("utf-8")
    parts[SHEET] = trim_rows(sheet_xml, LAST_ROW, add_trailing_cells).encode("utf-8")

    # Every other sheet keeps only its header row, so the fixture stays small
    # while still proving that untargeted sheets survive an export.
    for name in list(parts):
        if name.startswith("xl/worksheets/sheet") and name != SHEET:
            parts[name] = trim_rows(parts[name].decode("utf-8"), 1).encode("utf-8")

    # The table, its filter and the defined name all describe the same extent.
    table = parts[TABLE].decode("utf-8")
    table = table.replace("A1:H2408", "A1:H%d" % LAST_ROW)
    table = table.replace("$A$1:$H$2408", "$A$1:$H$%d" % LAST_ROW)
    parts[TABLE] = table.encode("utf-8")

    workbook = parts[WORKBOOK].decode("utf-8")
    workbook = workbook.replace("$A$1:$H$2408", "$A$1:$H$%d" % LAST_ROW)
    parts[WORKBOOK] = workbook.encode("utf-8")

    parts[DRAWING] = build_drawing().encode("utf-8")
    parts[DRAWING_RELS] = build_drawing_rels().encode("utf-8")

    for path in sorted(MEDIA):
        order.append(path)
        parts[path] = MEDIA[path]

    with zipfile.ZipFile(TARGET, "w", zipfile.ZIP_DEFLATED) as out:
        for name in order:
            out.writestr(name, parts[name])

    size = os.path.getsize(TARGET)
    print("Wrote %s (%d entries, %.0f KB)" % (TARGET, len(order), size / 1024.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
