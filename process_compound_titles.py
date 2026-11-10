"""
Compound Title Automation Script
=================================
Automates compound title harmonisation for the VMG Repertoire Audit workbook.

Usage:
    python process_compound_titles.py input.xlsm [output.xlsm]

Rules applied:
    1. For each UPC, apply the compound title from the highest-priority platform
       with real data:  Apple > Spotify > Tidal  -- applied wholesale to all rows.
    2. Blank rows (albums absent from a DSP) receive the same compound title as
       the other rows for that UPC.
    3. Tidal-only UPCs: look for a sibling UPC sharing the same album name that
       has Apple/Spotify data, and use that compound title.
    4. Remaining genuine edge cases are flagged in a REVIEW column placed
       immediately after the last column of the IMPORT sheet (AM in the current layout).

Output format:
    .xlsm -- full macro workbook. The output zip is a copy of the source in which
    only three parts change: the IMPORT worksheet xml (Compound title values and the
    REVIEW column), xl/sharedStrings.xml (new strings appended), and xl/workbook.xml
    (fullCalcOnLoad so the helper formula columns recalculate when the file is opened).
    VBA, relationships, table definitions, Power Query parts and formulas are untouched.

Design notes (why it is written this way):
    * Columns are located by header name, not by letter.
    * The IMPORT sheet is located through xl/workbook.xml, not assumed to be sheet1.xml.
    * Rows are matched by position, not by (UPC, Platform), so duplicate pairs are safe.
    * All cell types are read (shared string, inline string, formula string, number).
    * The sheet is processed row by row as text, so a 30,000+ row sheet does not need
      a full XML tree in memory.
    * The run fails with an error if the number of cells written does not match.
"""

import sys, os, re, zipfile, html
from decimal import Decimal
from xml.sax.saxutils import escape
from lxml import etree
import pandas as pd

PLATFORM_PRIORITY = {"Apple": 0, "Spotify": 1, "Tidal": 2}

REVIEW_NOTES = {
    "no_data": "NO DATA: no album data on any DSP for this UPC \u2013 no title can be built",
    "tidal_only": "TIDAL-ONLY: no sibling UPC found – verify artist name order",
    "edge_multi": "MULTI-ARTIST: different credits across DSPs – choose one title",
}

REQUIRED_HEADERS = ["Platform", "Artist name", "Album name",
                    "Number of audio tracks", "Compound title", "UPC"]


def _n(v):
    """Track count as text; never crashes on a blank/odd value."""
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return "unknown"

# ── Processing logic ──────────────────────────────────────────────────────────

def norm(name):
    if pd.isna(name): return ""
    n = re.sub(r'\s*&\s*', ', ', str(name).strip())
    return re.sub(r'\s+', ' ', n).lower().strip(', ')

def artists_equiv(names):
    return len({norm(n) for n in names if pd.notna(n) and str(n).strip()}) <= 1

def _isna(x):
    return x is None or (isinstance(x, float) and x != x)

def compute_titles(df):
    """Return (compound_titles Series, resolution Series) for all rows.

    Same rules as before (Apple > Spotify > Tidal, blanks propagated, Tidal-only UPCs
    resolved from a sibling UPC with the same album name), implemented on plain Python
    lists so 30,000+ rows take seconds rather than minutes.
    """
    plat  = df["Platform"].tolist()
    art   = df["Artist name"].tolist()
    alb   = df["Album name"].tolist()
    ct    = df["Compound title"].astype(object).tolist()
    trk   = df["Number of audio tracks"].tolist()
    upcs  = df["UPC"].tolist()
    n     = len(df)
    res   = ["auto"] * n

    groups = {}
    for i, u in enumerate(upcs):
        groups.setdefault(u, []).append(i)
    order = sorted(groups)                       # same order pandas groupby used

    def real_rows(idx):
        return [i for i in idx if not _isna(art[i])]

    def is_tidal_only(idx):
        r = real_rows(idx)
        return bool(r) and {plat[i] for i in r} == {"Tidal"}

    def build(i):
        return f"{art[i]} - {alb[i]} - {_n(trk[i])} tracks"

    tidal = {u: is_tidal_only(groups[u]) for u in order}

    # Stage 1: every UPC that is not Tidal-only
    for u in order:
        if tidal[u]:
            continue
        idx = groups[u]
        real = real_rows(idx)
        names = {norm(art[i]) for i in real if str(art[i]).strip()}
        multi = len(names) > 1
        auth = None
        for p_ in ("Apple", "Spotify", "Tidal"):
            rows = [i for i in real if plat[i] == p_]
            if rows and not _isna(ct[rows[0]]):
                auth = ct[rows[0]]
                break
        if auth is None and real:
            best = min(real, key=lambda i: PLATFORM_PRIORITY.get(plat[i], 99))   # stable: first of equals
            auth = build(best)
        if auth is None:
            for i in idx:
                res[i] = "no_data"          # nothing to build a title from: make it visible
            continue
        for i in idx:
            was_blank = _isna(ct[i])
            ct[i] = auth
            if multi:
                res[i] = "edge_multi"
            elif was_blank:
                res[i] = "propagated"

    # Stage 2: album name -> Apple/Spotify title (from non-Tidal-only UPCs)
    album_to_title = {}
    for u in order:
        if tidal[u]:
            continue
        real = real_rows(groups[u])
        for p_ in ("Apple", "Spotify"):
            rows = [i for i in real if plat[i] == p_]
            if rows:
                i0 = rows[0]
                if not _isna(alb[i0]) and not _isna(ct[i0]) and alb[i0] not in album_to_title:
                    album_to_title[alb[i0]] = ct[i0]
                break

    # Stage 3: Tidal-only UPCs
    for u in order:
        if not tidal[u]:
            continue
        idx = groups[u]
        t0 = next(i for i in idx if plat[i] == "Tidal")
        album = alb[t0]
        sibling = album_to_title.get(album) if not _isna(album) else None
        if sibling:
            auth, r_ = sibling, "sibling"
        else:
            auth = ct[t0]
            if _isna(auth) and not _isna(art[t0]):
                auth = build(t0)
            r_ = "tidal_only"
        if isinstance(auth, str) and auth or _isna(auth):
            for i in idx:
                ct[i] = auth if not _isna(auth) else None
                res[i] = r_

    return (pd.Series(ct, index=df.index, dtype=object),
            pd.Series(res, index=df.index, dtype=object))


# ── Worksheet / shared-string plumbing ────────────────────────────────────────

NS_XL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

ROW_RE  = re.compile(rb'<row\b[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RE = re.compile(rb'<c\b([^>]*?)(?:/>|>(.*?)</c>)', re.S)
R_ATTR  = re.compile(rb'\br="([A-Z]+)(\d+)"')
T_ATTR  = re.compile(rb'\bt="([^"]*)"')
S_ATTR  = re.compile(rb'\bs="([^"]*)"')
V_RE    = re.compile(rb'<v>(.*?)</v>', re.S)
IS_T_RE = re.compile(rb'<t\b[^>]*>(.*?)</t>', re.S)


def col_to_num(col):
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n

def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _unxml(b):
    return html.unescape(b.decode("utf-8", "replace"))

def read_shared_strings(xml_bytes):
    """Plain-text value of every shared string, in order (rich text flattened)."""
    out = []
    t_tag, r_tag = f"{{{NS_XL}}}t", f"{{{NS_XL}}}r"
    for _, si in etree.iterparse(__import__("io").BytesIO(xml_bytes),
                                 tag=f"{{{NS_XL}}}si", huge_tree=True):
        parts = [t.text or "" for t in si.findall(t_tag)]
        for r in si.findall(r_tag):
            parts += [t.text or "" for t in r.findall(t_tag)]
        out.append("".join(parts))
        si.clear()
    return out

def cell_text(attrs, inner, ss):
    """Value of a cell as text (None if empty). Handles s / inlineStr / str / numbers."""
    if not inner:
        return None
    tm = T_ATTR.search(attrs)
    t = tm.group(1) if tm else b"n"
    if t == b"inlineStr":
        txt = "".join(_unxml(x) for x in IS_T_RE.findall(inner))
        return txt if txt != "" else None
    vm = V_RE.search(inner)
    if not vm:
        return None
    raw = vm.group(1)
    if t == b"s":
        i = int(raw)
        return ss[i] if i < len(ss) else None
    if t in (b"e", b"b"):
        return None
    txt = _unxml(raw)
    if t == b"str":
        return txt if txt != "" else None
    # numeric: render integers without exponent / trailing .0
    try:
        d = Decimal(txt)
        txt = format(d, "f")
        if "." in txt:
            txt = txt.rstrip("0").rstrip(".")
    except Exception:
        pass
    return txt

def find_sheet_part(zin, sheet_name):
    wb = zin.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet\b[^>]*\bname="%s"[^>]*>' % re.escape(sheet_name), wb)
    if not m:
        raise ValueError(f'No sheet named "{sheet_name}" in this workbook.')
    rid = re.search(r'\br:id="([^"]+)"', m.group(0)).group(1)
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        if re.search(r'\bId="%s"' % re.escape(rid), rel):
            tgt = re.search(r'\bTarget="([^"]+)"', rel).group(1)
            return tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
    raise ValueError("Could not resolve the IMPORT sheet file.")

CHUNK = 4 * 1024 * 1024

def stream_sheet(fobj):
    """Walk a worksheet xml stream without loading it whole.
    Yields ("head", bytes) once, then ("row", bytes) / ("gap", bytes), then ("tail", bytes)."""
    buf = b""
    while True:
        chunk = fobj.read(CHUNK)
        buf += chunk
        i = buf.find(b"<sheetData")
        j = buf.find(b">", i) if i != -1 else -1
        if j != -1:
            break
        if not chunk:
            raise ValueError("Could not find <sheetData> in the worksheet.")
    yield "head", buf[:j + 1]
    buf = buf[j + 1:]
    while True:
        end_idx = buf.find(b"</sheetData>")
        region_end = end_idx if end_idx != -1 else len(buf)
        last_end = 0
        for m in ROW_RE.finditer(buf, 0, region_end):
            if m.start() > last_end:
                yield "gap", buf[last_end:m.start()]
            yield "row", m.group(0)
            last_end = m.end()
        if end_idx != -1:
            if region_end > last_end:
                yield "gap", buf[last_end:region_end]
            yield "tail", buf[end_idx:]
            while True:
                chunk = fobj.read(CHUNK)
                if not chunk:
                    return
                yield "tail", chunk
        buf = buf[last_end:]
        chunk = fobj.read(CHUNK)
        if not chunk:
            raise ValueError("Unexpected end of worksheet xml.")
        buf += chunk

def iter_cells(row_bytes):
    for m in CELL_RE.finditer(row_bytes):
        attrs, inner = m.group(1), m.group(2)
        rm = R_ATTR.search(attrs)
        if rm:
            yield rm.group(1).decode(), attrs, inner, m


# ── Main ──────────────────────────────────────────────────────────────────────

def process(input_path, output_path):
    print(f"\nReading: {input_path}")

    with zipfile.ZipFile(input_path, "r") as zin:
        sheet_part = find_sheet_part(zin, "IMPORT")
        ss_bytes = zin.read("xl/sharedStrings.xml")
        wb_bytes = zin.read("xl/workbook.xml")
    ss = read_shared_strings(ss_bytes)
    n_old_ss = len(ss)

    # ---- pass 1: header, data rows -> DataFrame, sheet extent (streamed) ----
    header_cols, last_col_num, last_hdr_style = {}, 0, None
    records, rownums = [], []
    max_row, ct_style = 1, [None]
    wanted = {}
    with zipfile.ZipFile(input_path, "r") as zin, zin.open(sheet_part) as f:
        for kind, data in stream_sheet(f):
            if kind != "row":
                continue
            rnum = int(re.search(rb'\br="(\d+)"', data).group(1))
            max_row = max(max_row, rnum)
            if rnum == 1:
                for col, attrs, inner, _ in iter_cells(data):
                    last_col_num = max(last_col_num, col_to_num(col))
                    sm = S_ATTR.search(attrs)
                    last_hdr_style = sm.group(1) if sm else last_hdr_style
                    name = cell_text(attrs, inner, ss)
                    if name:
                        header_cols[name.strip()] = col
                missing = [h for h in REQUIRED_HEADERS if h not in header_cols]
                if missing:
                    raise ValueError("The IMPORT sheet is missing expected column header(s): "
                                     + ", ".join(missing))
                wanted = {header_cols[h]: h for h in REQUIRED_HEADERS}
                continue
            if not wanted:
                raise ValueError("Header row (row 1) of the IMPORT sheet was not found.")
            rec = {}
            for col, attrs, inner, _ in iter_cells(data):
                if col in wanted:
                    rec[wanted[col]] = cell_text(attrs, inner, ss)
                if col == header_cols["Compound title"] and ct_style[0] is None:
                    sm = S_ATTR.search(attrs)
                    ct_style[0] = sm.group(1) if sm else None
            for k_, v_ in list(rec.items()):          # empty string == missing (as pandas does)
                if isinstance(v_, str) and v_.strip() == "":
                    rec[k_] = None
            plat = (rec.get("Platform") or "").strip()
            if not plat:
                continue
            rec["Platform"] = plat
            upc = (rec.get("UPC") or "").strip()
            rec["UPC"] = upc if upc else f"__no_upc_row_{rnum}"
            try:
                rec["Number of audio tracks"] = float(rec.get("Number of audio tracks"))
            except (TypeError, ValueError):
                rec["Number of audio tracks"] = float("nan")
            records.append(rec); rownums.append(rnum)

    if not records:
        raise ValueError("No data rows found on the IMPORT sheet.")
    review_col = num_to_col(last_col_num + 1)
    ct_col = header_cols["Compound title"]

    df = pd.DataFrame(records, columns=REQUIRED_HEADERS)
    del records
    for c in ("Platform", "Artist name", "Album name", "Compound title", "UPC"):
        df[c] = df[c].astype(object)
    df["Number of audio tracks"] = df["Number of audio tracks"].astype(float)

    compound_titles, resolution = compute_titles(df)

    # Summary (same lines the web app parses)
    c = resolution.value_counts()
    print(f"\n{'='*55}")
    print(f"  SUMMARY")
    print(f"{'='*55}")
    print(f"  Total rows                    : {len(df)}")
    print(f"  Auto-resolved                 : {c.get('auto',0) + c.get('propagated',0)}")
    print(f"    – harmonised                : {c.get('auto',0)}")
    print(f"    – blank rows propagated     : {c.get('propagated',0)}")
    print(f"  Tidal-only via sibling lookup : {c.get('sibling',0)}")
    print(f"  Flagged for review            : {c.get('tidal_only',0) + c.get('edge_multi',0) + c.get('no_data',0)}")
    print(f"    – Tidal-only, no sibling    : {c.get('tidal_only',0)}")
    print(f"    – Genuine multi-artist      : {c.get('edge_multi',0)}")
    print(f"    – No data on any DSP        : {c.get('no_data',0)}")
    print(f"  Null remaining                : {compound_titles.isna().sum()}")
    print(f"{'='*55}")

    # ---- shared strings: append only what is new ----
    ss_index, new_strings = {}, []
    def sidx(value):
        i = ss_index.get(value)
        if i is None:
            i = n_old_ss + len(new_strings)
            new_strings.append(value)
            ss_index[value] = i
        return i

    pos_of_row = {rn: p for p, rn in enumerate(rownums)}
    titles = compound_titles.tolist()
    reskeys = resolution.tolist()
    review_hdr_idx = sidx("REVIEW")
    stats = {"written": 0, "refs_added": 0}
    ct_colb, rv_colb = ct_col.encode(), review_col.encode()

    def edit_row(rowb):
        rnum = int(re.search(rb'\br="(\d+)"', rowb).group(1))
        is_header = (rnum == 1)
        pos = pos_of_row.get(rnum)
        if not is_header and pos is None:
            return rowb
        open_tag = re.match(rb"<row\b[^>]*>", rowb).group(0)
        self_closing = open_tag.endswith(b"/>")
        inner = b"" if self_closing else rowb[len(open_tag):-len(b"</row>")]
        new_open = re.sub(rb'\sspans="[^"]*"', b"", open_tag)
        if self_closing:
            new_open = new_open[:-2] + b">"

        def put(col, make, inner):
            """Replace the cell in `col` or insert it in column order."""
            cells = list(iter_cells(inner))
            for c_ in cells:
                if c_[0] == col:
                    return inner[:c_[3].start()] + make(c_[1]) + inner[c_[3].end():]
            cn = col_to_num(col)
            for c_ in cells:
                if col_to_num(c_[0]) > cn:
                    return inner[:c_[3].start()] + make(b"") + inner[c_[3].start():]
            return inner + make(b"")

        if is_header:
            def hdr_cell(_):
                st = (b' s="%s"' % last_hdr_style) if last_hdr_style else b""
                return b'<c r="%s1"%s t="s"><v>%d</v></c>' % (rv_colb, st, review_hdr_idx)
            inner = put(review_col, hdr_cell, inner)
            stats["refs_added"] += 1
        else:
            title, res = titles[pos], reskeys[pos]
            if isinstance(title, str) and title:
                idx = sidx(title)
                def ct_cell(old_attrs):
                    sm = S_ATTR.search(old_attrs) if old_attrs else None
                    sty = sm.group(1) if sm else ct_style[0]
                    st = (b' s="%s"' % sty) if sty else b""
                    return b'<c r="%s%d"%s t="s"><v>%d</v></c>' % (ct_colb, rnum, st, idx)
                inner = put(ct_col, ct_cell, inner)
                stats["written"] += 1
                stats["refs_added"] += 1
            note = REVIEW_NOTES.get(res)
            if note:
                nidx = sidx(note)
                inner = put(review_col,
                            lambda _: b'<c r="%s%d" t="s"><v>%d</v></c>' % (rv_colb, rnum, nidx),
                            inner)
                stats["refs_added"] += 1
        return new_open + inner + b"</row>"

    # ---- pass 2: stream the sheet through, writing the new output zip ----
    tmp_out = output_path + ".tmp"
    print(f"\nWriting: {output_path}")
    dim_ref = b'A1:%s%d' % (rv_colb, max_row)
    try:
        with zipfile.ZipFile(input_path, "r") as zin, \
             zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == sheet_part:
                    with zin.open(item) as src, zout.open(item, "w", force_zip64=True) as fh:
                        for kind, data in stream_sheet(src):
                            if kind == "head":
                                data = re.sub(rb'<dimension ref="[^"]*"',
                                              b'<dimension ref="' + dim_ref + b'"', data, count=1)
                            elif kind == "row":
                                data = edit_row(data)
                            fh.write(data)
                elif item.filename in ("xl/sharedStrings.xml", "xl/workbook.xml"):
                    continue                      # written after the sheet (needs final counts)
                else:
                    with zin.open(item) as src, zout.open(item, "w", force_zip64=True) as fh:
                        while True:
                            blk = src.read(1 << 20)
                            if not blk:
                                break
                            fh.write(blk)

            expected = int(sum(1 for t in titles if isinstance(t, str) and t))
            if stats["written"] != expected:
                raise RuntimeError(f"Internal check failed: wrote {stats['written']} compound "
                                   f"titles but expected {expected}. Output not created.")

            # shared strings: append new <si>, update counts
            add = []
            for s_ in new_strings:
                sp = ' xml:space="preserve"' if s_ != s_.strip() else ""
                add.append(f"<si><t{sp}>{escape(s_)}</t></si>")
            new_ss = ss_bytes.replace(b"</sst>", "".join(add).encode("utf-8") + b"</sst>")
            sst_tag = re.search(rb"<sst\b[^>]*>", new_ss).group(0)
            cm = re.search(rb'\bcount="(\d+)"', sst_tag)
            old_count = int(cm.group(1)) if cm else n_old_ss
            tag2 = re.sub(rb'\s(count|uniqueCount)="\d+"', b"", sst_tag)
            tag2 = tag2[:-1] + (b' count="%d" uniqueCount="%d">' % (
                old_count + stats["refs_added"], n_old_ss + len(new_strings)))
            new_ss = new_ss.replace(sst_tag, tag2, 1)
            zout.writestr(zin.getinfo("xl/sharedStrings.xml"), new_ss)

            # workbook.xml: full recalculation on open so helper formula columns refresh
            wbx = wb_bytes
            cp = re.search(rb"<calcPr\b[^>]*?/?>", wbx)
            if cp:
                tag = cp.group(0)
                newtag = (re.sub(rb'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="1"', tag)
                          if b"fullCalcOnLoad" in tag
                          else tag.replace(b"<calcPr", b'<calcPr fullCalcOnLoad="1"', 1))
                wbx = wbx.replace(tag, newtag, 1)
            else:
                wbx = wbx.replace(b"</workbook>", b'<calcPr fullCalcOnLoad="1"/></workbook>', 1)
            zout.writestr(zin.getinfo("xl/workbook.xml"), wbx)
        os.replace(tmp_out, output_path)
    except Exception:
        if os.path.exists(tmp_out):
            os.remove(tmp_out)
        raise

    flagged = c.get("tidal_only", 0) + c.get("edge_multi", 0) + c.get("no_data", 0)
    edge_upcs = df[resolution.isin(["tidal_only", "edge_multi", "no_data"])]["UPC"].nunique()
    print(f"Compound titles written: {stats['written']} of {len(df)} rows; "
          f"REVIEW column: {review_col}")
    print(f"Done. {flagged} rows flagged across {edge_upcs} UPCs.\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python process_compound_titles.py input.xlsm [output.xlsm]")
        sys.exit(1)
    inf = sys.argv[1]
    outf = sys.argv[2] if len(sys.argv) > 2 else inf.rsplit(".", 1)[0] + "_processed.xlsm"
    process(inf, outf)
