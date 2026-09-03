"""
Fygaro Catalog Automation
Regenerates tests/fixtures/snapshots.js from the captured page snapshots.

The snapshots in temp/ are the raw DOM of each Fygaro screen the automation
drives. They are copied into a committed fixture file so the locator tests can
load them from file:// without a server, and third party iframes are stripped so
the tests never reach the network.

Run from the project root:  python tools/build-fixtures.py
"""

import io
import json
import os
import re
import sys

SOURCE_DIR = os.path.join("temp", "html")
TARGET = os.path.join("tests", "fixtures", "snapshots.js")
SKIP = {"prompt.txt"}

HEADER = """/*
 * Fygaro Catalog Automation
 * Page snapshots captured from the live Fygaro app, used by the locator tests.
 * Generated from temp/html by tools/build-fixtures.py. Third party iframes are
 * stripped so the tests never touch the network.
 */
window.FYG_SNAPSHOTS = """


def main():
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


if __name__ == "__main__":
    sys.exit(main())
