"""Harmonise LS treatment/class codes with the MS coding so the two datasets can be merged.

- trt_to_class_ms.csv is the base: its treatment and class codes are kept unchanged.
- Every treatment/class in trt_to_class_ls.csv is matched to MS by name. Treatments and
  classes that do not exist in MS are appended after the last MS code (in LS order).
- The t[,1]..t[,5] columns of every data block in the "LS SMD bias-adj" sheet are recoded
  to the merged codes, and the sheet's TREATMENT AND CLASS CODES table is replaced by the
  merged table. The "MS SMD bias-adj" sheet is not modified.

Usage: python3 harmonise_codes.py   (requires lxml)
Only the LS worksheet XML and the shared-strings part are rewritten; every other part of
the workbook (MS sheet, Excel tables, styles) is copied unchanged.
"""

import csv
import os
import re
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
IN_DIR = os.path.join(HERE, "input")
OUT_DIR = os.path.join(HERE, "output")

LS_SHEET = "LS SMD bias-adj"
TRT_COLS = range(2, 7)  # columns B..F hold t[,1]..t[,5]
CODE_TABLE_COL = 39  # column AM: trtcode, trt, classcode, class
SST_PATH = "xl/sharedStrings.xml"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def norm(s):
    return " ".join(str(s).split()).casefold()


def read_codes(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return [
            {
                "trtcode": int(r["trtcode"]),
                "trt": r["trt"].strip(),
                "classcode": int(r["classcode"]),
                "class": r["class"].strip(),
            }
            for r in csv.DictReader(f, delimiter=";")
        ]


def write_csv(path, rows, fields):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";", lineterminator="\r\n")
        w.writeheader()
        w.writerows(rows)


def build_merge(ms, ls):
    merged = [dict(r) for r in ms]
    trt_by_name = {norm(r["trt"]): r for r in merged}
    class_by_name = {norm(r["class"]): r["classcode"] for r in merged}
    next_trt = max(r["trtcode"] for r in merged) + 1
    next_class = max(r["classcode"] for r in merged) + 1

    trt_map, class_map, mapping_rows, warnings = {}, {}, [], []
    for r in ls:
        cname = norm(r["class"])
        if cname not in class_by_name:
            class_by_name[cname] = next_class
            next_class += 1
        new_class = class_by_name[cname]
        if class_map.setdefault(r["classcode"], new_class) != new_class:
            warnings.append(f"LS class code {r['classcode']} maps to several merged classes")

        tname = norm(r["trt"])
        if tname in trt_by_name:
            m = trt_by_name[tname]
            status = "matched"
            if m["classcode"] != new_class:
                warnings.append(
                    f"'{r['trt']}': LS class '{r['class']}' differs from MS class '{m['class']}'"
                    " - MS class kept"
                )
        else:
            m = {"trtcode": next_trt, "trt": r["trt"], "classcode": new_class, "class": r["class"]}
            next_trt += 1
            merged.append(m)
            trt_by_name[tname] = m
            status = "added"
        trt_map[r["trtcode"]] = m["trtcode"]
        mapping_rows.append(
            {
                "trt": r["trt"],
                "ls_trtcode": r["trtcode"],
                "merge_trtcode": m["trtcode"],
                "class": r["class"],
                "ls_classcode": r["classcode"],
                "merge_classcode": m["classcode"],
                "status": status,
            }
        )
    return merged, trt_map, mapping_rows, warnings


def col_index(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def recode_ls_sheet(sheet_xml, sst_xml, trt_map, merged):
    """Edit the LS worksheet XML in place so the rest of the workbook stays byte-identical."""
    sheet = etree.fromstring(sheet_xml)
    sst = etree.fromstring(sst_xml)
    strings = ["".join(si.itertext()) for si in sst.findall(f"{{{NS}}}si")]
    string_idx = {}
    for i, t in enumerate(strings):
        string_idx.setdefault(t, i)

    def shared(text):
        if text not in string_idx:
            si = etree.SubElement(sst, f"{{{NS}}}si")
            etree.SubElement(si, f"{{{NS}}}t").text = text
            string_idx[text] = len(strings)
            strings.append(text)
        return string_idx[text]

    def value(c):
        v = c.find(f"{{{NS}}}v")
        if v is None:
            return None
        return strings[int(v.text)] if c.get("t") == "s" else float(v.text)

    sheet_data = sheet.find(f"{{{NS}}}sheetData")
    rows = {int(r.get("r")): r for r in sheet_data.findall(f"{{{NS}}}row")}

    # 1) Recode t[,1]..t[,5] in every data block (a block starts at a "na[]" header row).
    n = 0
    in_block = False
    for rnum in sorted(rows):
        cells = {col_index(c.get("r")): c for c in rows[rnum].findall(f"{{{NS}}}c")}
        first = value(cells[1]) if 1 in cells else None
        if first == "na[]":
            in_block = True
            continue
        if not isinstance(first, float):
            in_block = False
            continue
        if not in_block:
            continue
        for col in TRT_COLS:
            c = cells.get(col)
            if c is None or c.get("t") is not None or c.find(f"{{{NS}}}v") is None:
                continue  # empty or text ("NA")
            v = c.find(f"{{{NS}}}v")
            code = int(float(v.text))
            if code not in trt_map:
                raise KeyError(f"{c.get('r')}: treatment code {code} not in trt_to_class_ls")
            v.text = str(trt_map[code])
            n += 1

    # 2) Replace the TREATMENT AND CLASS CODES table (columns AM:AP, data from row 3).
    table_cols = range(CODE_TABLE_COL, CODE_TABLE_COL + 4)
    style = None
    removed_strings = 0
    for rnum, row in rows.items():
        if rnum < 3:
            continue
        for c in row.findall(f"{{{NS}}}c"):
            if col_index(c.get("r")) in table_cols:
                style = style or c.get("s")
                removed_strings += c.get("t") == "s"
                row.remove(c)

    added_strings = 0
    for i, rec in enumerate(merged):
        rnum = 3 + i
        row = rows.get(rnum)
        if row is None:
            row = etree.Element(f"{{{NS}}}row", r=str(rnum))
            later = [r for k, r in rows.items() if k > rnum]
            if later:
                min(later, key=lambda r: int(r.get("r"))).addprevious(row)
            else:
                sheet_data.append(row)
            rows[rnum] = row
        values = (rec["trtcode"], rec["trt"], rec["classcode"], rec["class"])
        new_cells = []
        for col, val in zip(table_cols, values):
            c = etree.Element(f"{{{NS}}}c", r=f"{col_letter(col)}{rnum}")
            if style:
                c.set("s", style)
            if isinstance(val, str):
                c.set("t", "s")
                val = shared(val)
                added_strings += 1
            etree.SubElement(c, f"{{{NS}}}v").text = str(val)
            new_cells.append(c)
        after = [c for c in row.findall(f"{{{NS}}}c") if col_index(c.get("r")) > table_cols[-1]]
        for c in new_cells:
            if after:
                after[0].addprevious(c)
            else:
                row.append(c)

    dim = sheet.find(f"{{{NS}}}dimension")
    if dim is not None:
        start_ref, _, end_ref = dim.get("ref").partition(":")
        end_ref = end_ref or start_ref
        end_row = max(int(re.search(r"\d+$", end_ref).group(0)), max(rows))
        dim.set("ref", f"{start_ref}:{re.match(r'[A-Z]+', end_ref).group(0)}{end_row}")
    sst.set("count", str(int(sst.get("count")) - removed_strings + added_strings))
    sst.set("uniqueCount", str(len(strings)))

    def dump(root):
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    return dump(sheet), dump(sst), n


def sheet_path(z, name):
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = next(s.get(f"{{{NS_R}}}id") for s in wb.iter(f"{{{NS}}}sheet") if s.get("name") == name)
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return target.lstrip("/") if target.startswith("/") else "xl/" + target


def main():
    ms = read_codes(os.path.join(IN_DIR, "trt_to_class_ms.csv"))
    ls = read_codes(os.path.join(IN_DIR, "trt_to_class_ls.csv"))
    merged, trt_map, mapping_rows, warnings = build_merge(ms, ls)

    os.makedirs(OUT_DIR, exist_ok=True)
    write_csv(os.path.join(OUT_DIR, "trt_to_class_merge.csv"), merged, ["trtcode", "trt", "classcode", "class"])
    write_csv(os.path.join(OUT_DIR, "ls_to_merge_code_map.csv"), mapping_rows, list(mapping_rows[0]))

    src = os.path.join(IN_DIR, "merge_dataset.xlsx")
    dst = os.path.join(OUT_DIR, "merge_dataset.xlsx")
    with zipfile.ZipFile(src) as zin:
        ls_path = sheet_path(zin, LS_SHEET)
        sheet_xml, sst_xml, n = recode_ls_sheet(zin.read(ls_path), zin.read(SST_PATH), trt_map, merged)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = {ls_path: sheet_xml, SST_PATH: sst_xml}.get(item.filename) or zin.read(item.filename)
                zout.writestr(item, data)

    added = sum(r["status"] == "added" for r in mapping_rows)
    print(f"Merged table: {len(merged)} treatments, {max(r['classcode'] for r in merged)} classes")
    print(f"LS treatments matched to MS: {len(mapping_rows) - added}, added: {added}")
    print(f"Recoded {n} treatment cells in '{LS_SHEET}'")
    for w in warnings:
        print("WARNING:", w)


if __name__ == "__main__":
    main()
