#!/usr/bin/env python3
"""Inject morph_mix transitions into an assembled PPTX (post-assembly, XML-level).

morph_mix = Morph on related-slide pairs + varied fade/push/wipe elsewhere.
Usage: add_transitions.py <deck.pptx> [--out <out.pptx>]
"""
from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

# slide number (transition applied when ENTERING it) -> transition kind
PLAN = {
    1: ("fade", None),
    2: ("push", "u"),
    3: ("fade", None),
    4: ("morph", None),      # 1947-71 timeline -> constitution arc (related)
    5: ("wipe", "r"),
    6: ("push", "l"),        # history -> framework
    7: ("morph", None),      # four threads -> peoples thread (related)
    8: ("fade", None),
    9: ("morph", None),      # institutions -> census of same photos (related)
    10: ("morph", None),     # numbers -> peoples (related)
    11: ("wipe", "r"),
    12: ("push", "u"),
    13: ("fade", None),
    14: ("wipe", "l"),
    15: ("push", "r"),
    16: ("morph", None),     # politics -> conclusion (related)
    17: ("fade", None),
}

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def transition_xml(kind: str, direction: str | None) -> str:
    if kind == "morph":
        return (
            f'<mc:AlternateContent xmlns:mc="{MC}">'
            f'<mc:Choice xmlns:p159="{P159}" Requires="p159">'
            f'<p:transition xmlns:p14="{P14}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'spd="slow" p14:dur="1000"><p159:morph option="byObject"/></p:transition>'
            f'</mc:Choice>'
            f'<mc:Fallback>'
            f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'spd="slow"><p:fade/></p:transition>'
            f'</mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    if kind == "push":
        inner = f'<p:push dir="{direction or "u"}"/>'
        dur = 700
    elif kind == "wipe":
        inner = f'<p:wipe dir="{direction or "r"}"/>'
        dur = 700
    else:  # fade
        inner = "<p:fade/>"
        dur = 600
    return (
        f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="{P14}" spd="med" p14:dur="{dur}">{inner}</p:transition>'
    )


def inject(slide_xml: str, kind: str, direction: str | None) -> str:
    # drop any existing transition (idempotent)
    slide_xml = re.sub(r"<p:transition\b.*?</p:transition>", "", slide_xml, flags=re.S)
    slide_xml = re.sub(r"<mc:AlternateContent\b.*?</mc:AlternateContent>", "", slide_xml, flags=re.S)
    node = transition_xml(kind, direction)
    # insert after </p:clrMapOvr>, else before </p:sld>
    if "</p:clrMapOvr>" in slide_xml:
        return slide_xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + node, 1)
    return slide_xml.replace("</p:sld>", node + "</p:sld>", 1)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    src = Path(args.pptx)
    out = Path(args.out) if args.out else src

    with zipfile.ZipFile(src) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    touched = 0
    for n, data in list(items.items()):
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
        if not m:
            continue
        idx = int(m.group(1))
        if idx not in PLAN:
            continue
        kind, direction = PLAN[idx]
        xml = data.decode("utf-8")
        if xml.find("<p:transition") != -1 or xml.find("mc:AlternateContent") != -1:
            pass  # will be replaced by inject()
        new = inject(xml, kind, direction)
        items[n] = new.encode("utf-8")
        touched += 1
        print(f"slide {idx:02d} -> {kind}{' ' + direction if direction else ''}")

    tmp = Path(tempfile.mkstemp(suffix=".pptx")[1])
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in items.items():
            zout.writestr(name, blob)
    shutil.move(str(tmp), str(out))
    print(f"injected {touched} transitions -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
