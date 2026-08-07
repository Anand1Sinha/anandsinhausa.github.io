#!/usr/bin/env python3
"""
build_hero.py — The Signal Voter Guide shared component build step
====================================================================

WHAT THIS DOES
Injects each partial listed in PARTIALS below into both index.html
(New) and classic.html (Classic), byte-for-byte identical, between
that partial's own reserved marker pair:

    hero.html          -> <!-- SIGNAL:HERO:START ... -->  /  <!-- SIGNAL:HERO:END -->
    engagement.html     -> <!-- SIGNAL:ENGAGEMENT:START ... -->  /  <!-- SIGNAL:ENGAGEMENT:END -->

Everything outside those marker pairs in each file is left
untouched — below-the-hero content, JS, screens, and routing are
unaffected by this script.

HOW TO USE
1. Edit hero.html and/or engagement.html — never edit either
   section directly in index.html or classic.html.
2. Run:  python3 build_hero.py
3. Review the diff, then commit index.html and classic.html.

No dependencies beyond the Python 3 standard library. Output is
fully static — no runtime fetch, no client-side injection. The
generated files are exactly what gets deployed.

Run this from the same folder as the partials and target files
(voterguide/).
"""

import re
import sys
from pathlib import Path

# (partial filename, marker name used in <!-- SIGNAL:<name>:START/END -->)
PARTIALS = [
    ("hero.html", "HERO"),
    ("engagement.html", "ENGAGEMENT"),
]

TARGET_FILES = ["index.html", "classic.html"]


def marker_pattern(name):
    # Greedy on purpose: a partial's own content must never be able to
    # prematurely terminate this match, even if it happens to mention
    # the marker text somewhere (this bit us once during development).
    return re.compile(
        r"(<!-- SIGNAL:%s:START.*?-->\n)(.*)(<!-- SIGNAL:%s:END -->)" % (name, name),
        re.DOTALL,
    )


def main():
    base = Path(__file__).parent
    any_errors = False

    partial_contents = {}
    for filename, name in PARTIALS:
        path = base / filename
        if not path.exists():
            sys.exit(f"ERROR: {filename} not found in {base}")
        partial_contents[name] = path.read_text(encoding="utf-8")

    for target_filename in TARGET_FILES:
        target_path = base / target_filename
        if not target_path.exists():
            print(f"SKIP: {target_filename} not found in {base}")
            any_errors = True
            continue

        original = target_path.read_text(encoding="utf-8")
        text = original

        for filename, name in PARTIALS:
            start_marker = f"<!-- SIGNAL:{name}:START"
            end_marker = f"<!-- SIGNAL:{name}:END -->"

            if start_marker not in text or end_marker not in text:
                print(f"ERROR: {target_filename} is missing the "
                      f"SIGNAL:{name} markers — that partial was not "
                      f"injected.")
                any_errors = True
                continue

            pattern = marker_pattern(name)
            new_text, count = pattern.subn(
                lambda m, content=partial_contents[name]: m.group(1) + content + "\n" + m.group(3),
                text,
            )
            if count == 0:
                print(f"ERROR: SIGNAL:{name} markers found but pattern "
                      f"did not match in {target_filename} — that "
                      f"partial was not injected.")
                any_errors = True
                continue
            text = new_text

        if text == original:
            print(f"OK: {target_filename} already up to date, no change written.")
        else:
            target_path.write_text(text, encoding="utf-8")
            print(f"OK: {target_filename} updated.")

    if any_errors:
        sys.exit(1)

    # Self-verify: every partial's injected region must be byte-for-byte
    # identical across all target files.
    verify_failed = False
    for filename, name in PARTIALS:
        pattern = marker_pattern(name)
        regions = {}
        for target_filename in TARGET_FILES:
            text = (base / target_filename).read_text(encoding="utf-8")
            m = pattern.search(text)
            if not m:
                print(f"VERIFY FAILED: could not locate SIGNAL:{name} "
                      f"region in {target_filename} after build.")
                verify_failed = True
                continue
            regions[target_filename] = m.group(2)
        values = list(regions.values())
        if values and not all(v == values[0] for v in values):
            print(f"VERIFY FAILED: SIGNAL:{name} regions are NOT "
                  f"byte-for-byte identical across target files.")
            verify_failed = True

    if verify_failed:
        sys.exit("Nothing should have been committed — investigate before pushing.")

    print(f"\nVerified: every partial is byte-for-byte identical across "
          f"{', '.join(TARGET_FILES)}.")
    print("Done. Review the diffs, then commit.")


if __name__ == "__main__":
    main()
