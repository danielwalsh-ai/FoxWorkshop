"""
Transplant the COVER sheet into Mel's master without disturbing anything else.

Mel's workbook carries 180 chart parts and 16 drawings. Re-saving it through
openpyxl destroys them, so the cover sheet is built standalone (lancs_cover.py)
and its parts are copied into the original zip here, renumbered to avoid
collisions and registered as the first sheet.

The one non-obvious bit is styles: a cell's `s=` is an index into the workbook's
own styles.xml, so the cover's indices are meaningless inside Mel's file. Its
fonts, fills, number formats and cellXfs are appended to hers and the indices
remapped by the offset.
"""
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda t: f"{{{NS_MAIN}}}{t}"
QR = lambda t: f"{{{NS_R}}}{t}"


def _next(names, pat):
    ns = [int(m.group(1)) for n in names if (m := re.match(pat, n))]
    return (max(ns) if ns else 0) + 1


def _merge_styles(master_xml, cover_xml):
    """Append the cover's styles to the master's; return (new xml, index offset)."""
    mroot = etree.fromstring(master_xml)
    croot = etree.fromstring(cover_xml)

    def lst(root, tag):
        el = root.find(Q(tag))
        return el, (list(el) if el is not None else [])

    out = {}
    for tag in ("numFmts", "fonts", "fills", "borders"):
        mel, mitems = lst(mroot, tag)
        cel, citems = lst(croot, tag)
        out[tag] = (mel, len(mitems), citems)

    # numFmts are keyed by numFmtId, not position — renumber the cover's above
    # anything the master already uses.
    used = {int(n.get("numFmtId")) for n in out["numFmts"][2]} if out["numFmts"][2] else set()
    mfmt_ids = set()
    if out["numFmts"][0] is not None:
        mfmt_ids = {int(n.get("numFmtId")) for n in list(out["numFmts"][0])}
    base_fmt = max(mfmt_ids | {163}) + 1
    fmt_map = {}
    for i, n in enumerate(out["numFmts"][2]):
        old = int(n.get("numFmtId"))
        fmt_map[old] = base_fmt + i
        n.set("numFmtId", str(fmt_map[old]))

    for tag in ("numFmts", "fonts", "fills", "borders"):
        mel, count, citems = out[tag]
        if not citems:
            continue
        if mel is None:
            mel = etree.SubElement(mroot, Q(tag))
            mel.set("count", "0")
            count = 0
        for it in citems:
            mel.append(it)
        mel.set("count", str(count + len(citems)))

    off = {"fonts": out["fonts"][1], "fills": out["fills"][1], "borders": out["borders"][1]}

    mxfs = mroot.find(Q("cellXfs"))
    cxfs = croot.find(Q("cellXfs"))
    offset = len(list(mxfs)) if mxfs is not None else 0
    if cxfs is not None and mxfs is not None:
        for xf in list(cxfs):
            for attr, key in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                if xf.get(attr) is not None:
                    xf.set(attr, str(int(xf.get(attr)) + off[key]))
            if xf.get("numFmtId") is not None:
                old = int(xf.get("numFmtId"))
                if old in fmt_map:
                    xf.set("numFmtId", str(fmt_map[old]))
            mxfs.append(xf)
        mxfs.set("count", str(offset + len(list(cxfs))))
    return etree.tostring(mroot, xml_declaration=True, encoding="UTF-8",
                          standalone=True), offset


def _shift_style_indices(sheet_xml, offset):
    root = etree.fromstring(sheet_xml)
    for el in root.iter():
        if el.tag in (Q("c"), Q("row"), Q("col")) and el.get("s") is not None:
            el.set("s", str(int(el.get("s")) + offset))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def inject(master_path, cover_path, out_path, sheet_name="COVER", first=True):
    zm = zipfile.ZipFile(master_path)
    zc = zipfile.ZipFile(cover_path)
    mn, cn = zm.namelist(), zc.namelist()

    if any(re.search(rf'name="{re.escape(sheet_name)}"', zm.read("xl/workbook.xml").decode())
           for _ in [0]):
        pass  # replaced below if it already exists

    sheet_no = _next(mn, r"xl/worksheets/sheet(\d+)\.xml")
    draw_no = _next(mn, r"xl/drawings/drawing(\d+)\.xml")
    chart_base = _next(mn, r"xl/charts/chart(\d+)\.xml")

    cover_charts = sorted((n for n in cn if re.match(r"xl/charts/chart\d+\.xml$", n)),
                          key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    chart_map = {n: f"xl/charts/chart{chart_base + i}.xml"
                 for i, n in enumerate(cover_charts)}

    # styles first — the sheet's s= indices depend on the offset
    new_styles, offset = _merge_styles(zm.read("xl/styles.xml"), zc.read("xl/styles.xml"))

    sheet_src = next(n for n in cn if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    sheet_xml = _shift_style_indices(zc.read(sheet_src), offset)
    sheet_dst = f"xl/worksheets/sheet{sheet_no}.xml"

    draw_src = next((n for n in cn if re.match(r"xl/drawings/drawing\d+\.xml$", n)), None)
    draw_dst = f"xl/drawings/drawing{draw_no}.xml"

    # ── new parts ───────────────────────────────────────────────────
    new_parts = {sheet_dst: sheet_xml}
    if draw_src:
        new_parts[draw_dst] = zc.read(draw_src)
        rels_src = f"xl/drawings/_rels/{Path(draw_src).name}.rels"
        if rels_src in cn:
            r = zc.read(rels_src).decode()
            for old, new in chart_map.items():
                r = r.replace(f"../charts/{Path(old).name}", f"../charts/{Path(new).name}")
            new_parts[f"xl/drawings/_rels/drawing{draw_no}.xml.rels"] = r.encode()
        new_parts[f"xl/worksheets/_rels/sheet{sheet_no}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
            f'2006/relationships/drawing" Target="../drawings/drawing{draw_no}.xml"/>'
            '</Relationships>').encode()
    for old, new in chart_map.items():
        new_parts[new] = zc.read(old)
        rsrc = f"xl/charts/_rels/{Path(old).name}.rels"
        if rsrc in cn:
            new_parts[f"xl/charts/_rels/{Path(new).name}.rels"] = zc.read(rsrc)

    # ── workbook.xml: register the sheet, first in the tab order ────
    wb = etree.fromstring(zm.read("xl/workbook.xml"))
    sheets = wb.find(Q("sheets"))
    for s in list(sheets):
        if s.get("name") == sheet_name:
            sheets.remove(s)                       # replacing a previous run
    rels_xml = zm.read("xl/_rels/workbook.xml.rels").decode()
    existing = [int(x) for x in re.findall('Id="rId([0-9]+)"', rels_xml)]
    rid = f"rId{(max(existing) if existing else 0) + 1}"
    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"',
                                         zm.read("xl/workbook.xml").decode())) + 1
    el = etree.Element(Q("sheet"))
    el.set("name", sheet_name)
    el.set("sheetId", str(sid))
    el.set(QR("id"), rid)
    sheets.insert(0 if first else len(sheets), el)
    # open on the cover
    for bv in wb.findall(Q("bookViews")) + [wb]:
        for w in bv.findall(Q("workbookView")):
            w.set("activeTab", "0")
    new_wb = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)

    new_rels = rels_xml.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
        f'2006/relationships/worksheet" Target="worksheets/sheet{sheet_no}.xml"/>'
        "</Relationships>")

    # ── content types ───────────────────────────────────────────────
    ct = zm.read("[Content_Types].xml").decode()
    adds = [f'<Override PartName="/{sheet_dst}" ContentType="application/vnd.openxmlformats-'
            f'officedocument.spreadsheetml.worksheet+xml"/>']
    if draw_src:
        adds.append(f'<Override PartName="/{draw_dst}" ContentType="application/vnd.'
                    f'openxmlformats-officedocument.drawing+xml"/>')
    for new in chart_map.values():
        adds.append(f'<Override PartName="/{new}" ContentType="application/vnd.'
                    f'openxmlformats-officedocument.drawingml.chart+xml"/>')
    ct = ct.replace("</Types>", "".join(adds) + "</Types>")

    replace = {"xl/workbook.xml": new_wb,
               "xl/_rels/workbook.xml.rels": new_rels.encode(),
               "[Content_Types].xml": ct.encode(),
               "xl/styles.xml": new_styles}

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zm.infolist():
            if item.filename in new_parts:          # a previous cover run
                continue
            zo.writestr(item, replace.get(item.filename, zm.read(item.filename)))
        for name, data in new_parts.items():
            zo.writestr(name, data)
    zm.close(); zc.close()
    return {"sheet": sheet_dst, "charts": len(chart_map), "drawing": draw_dst,
            "style_offset": offset, "rid": rid, "sheetId": sid}


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("master"); ap.add_argument("cover"); ap.add_argument("-o", "--out",
                                                                        required=True)
    a = ap.parse_args()
    print(inject(a.master, a.cover, a.out))
