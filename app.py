"""
Daily Attendance IN/OUT Punch Report -> print-ready PDF (Streamlit app)

Upload the .xls / .xlsx punch report (one sheet per company/contractor) and get a
compact, print-ready PDF:
  * one "Company:" block per sheet, blocks flow one after another
  * period line repeated at the top of every page
  * IN-n / OUT-n columns are added ONLY when someone in that sheet really has a
    punch in that pair (IN-1/OUT-1 is always kept as the minimum)

Run:  streamlit run app.py
"""
import io
import re
from datetime import datetime, time as dtime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

PUNCH_RE = re.compile(r"^(IN|OUT)-(\d+)$", re.I)
MARGIN = 14 * mm


# --------------------------------------------------------------------------- #
# Reading the workbook
# --------------------------------------------------------------------------- #
def fmt_time(x) -> str:
    """Return HH:MM for any kind of time-like cell, '' for empty."""
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    if isinstance(x, (datetime, pd.Timestamp, dtime)):
        return x.strftime("%H:%M")
    if isinstance(x, float) and 0 <= x < 1:          # Excel time fraction
        mins = round(x * 24 * 60)
        return f"{mins // 60:02d}:{mins % 60:02d}"
    s = str(x).strip()
    m = re.search(r"(\d{1,2}:\d{2})(:\d{2})?(\.\d+)?$", s)
    return m.group(1).zfill(5) if m else s


def clean(x) -> str:
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    s = str(x).strip()
    return re.sub(r"\.0$", "", s) if re.fullmatch(r"\d+\.0", s) else s


def parse_sheet(df: pd.DataFrame):
    """Parse one sheet. Returns dict or None when no 'SNo' header row is found."""
    period = company = printed = title = ""
    header = None
    blocks, cur = [], None

    for i in range(len(df)):
        cells = [(j, v) for j, v in enumerate(df.iloc[i]) if clean(v)]
        if not cells:
            continue
        txt = [clean(v) for _, v in cells]
        first = txt[0]

        if header is None:
            if i < 3 and len(txt) == 1 and "report" in first.lower():
                title = re.sub(r"\s*\(.*?\)\s*$", "", first)
                continue
            if any(re.search(r"\bto\b", t, re.I) and re.search(r"\d{4}", t) for t in txt):
                period = re.sub(r"\s+", " ", next(t for t in txt if re.search(r"\bto\b", t, re.I)))
                continue
            if first.lower().startswith("company"):
                company = " ".join(t for t in txt[1:] if not t.lower().startswith("printed"))
                printed = next((t for t in txt if t.lower().startswith("printed")), "")
                continue
            if first.lower() == "date" and len(txt) > 1:
                cur = {"date": txt[1], "rows": []}
                blocks.append(cur)
                continue
            if first.lower() == "sno":
                header = [(j, clean(v)) for j, v in cells]
            continue

        # after header
        if first.lower() == "date" and len(txt) > 1:
            cur = {"date": txt[1], "rows": []}
            blocks.append(cur)
            continue
        if first.lower() == "sno":          # repeated header (multi-date sheets)
            continue
        if re.fullmatch(r"\d+(\.0)?", first):
            if cur is None:
                cur = {"date": "", "rows": []}
                blocks.append(cur)
            cur["rows"].append({j: v for j, v in cells})

    if header is None:
        return None

    punch = {}
    info = []
    for j, name in header:
        m = PUNCH_RE.match(name.replace(" ", ""))
        if m:
            punch[(int(m.group(2)), m.group(1).upper())] = j
        else:
            info.append((j, name))
    return dict(title=title, period=period, company=company, printed=printed,
                info=info, punch=punch, blocks=blocks)


def read_workbook(file) -> list:
    sheets = pd.read_excel(file, header=None, dtype=object, sheet_name=None)
    out = []
    for name, df in sheets.items():
        p = parse_sheet(df)
        if p:
            p["sheet"] = name
            out.append(p)
    return out


def used_pairs(p) -> list:
    """Pair numbers that contain at least one punch time in this sheet."""
    rows = [r for b in p["blocks"] for r in b["rows"]]
    nums = sorted({n for n, _ in p["punch"]})
    used = [n for n in nums
            if any(fmt_time(r.get(p["punch"].get((n, k)))) for r in rows for k in ("IN", "OUT"))]
    return used or nums[:1]            # always keep at least IN-1/OUT-1


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
def build_pdf(sheets, orientation="Auto", same_cols=False, font_size=8.5) -> bytes:
    global_pairs = sorted({n for p in sheets for n in used_pairs(p)}) if same_cols else None
    pairs_for = {id(p): (global_pairs or used_pairs(p)) for p in sheets}
    max_nk = max(2 * len(v) for v in pairs_for.values())

    land = orientation == "Landscape" or (orientation == "Auto" and max_nk > 6)
    page = landscape(A4) if land else portrait(A4)
    avail = page[0] - 2 * MARGIN

    period = next((p["period"] for p in sheets if p["period"]), "")
    buf = io.BytesIO()

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 10)
        canvas.drawCentredString(page[0] / 2, page[1] - 12 * mm, period)
        canvas.setFont("Helvetica", 7)
        canvas.drawRightString(page[0] - MARGIN, 7 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(buf, pagesize=page, leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=19 * mm, bottomMargin=12 * mm,
                            title="Daily Attendance IN/OUT Punch Report")
    co_style = ParagraphStyle("co", fontName="Helvetica", fontSize=11, leading=13)
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=font_size, leading=font_size + 1.5)
    head = ParagraphStyle("head", parent=cell, fontName="Helvetica-Bold")

    story = []
    for p in sheets:
        pairs = pairs_for[id(p)]
        pcols = [(n, k) for n in pairs for k in ("IN", "OUT")]
        info = p["info"]

        # column widths
        fixed = {"sno": 13, "e. code": 17, "department": 27, "shift": 10}
        widths = []
        for _, name in info:
            widths.append(fixed.get(name.lower(), 22 if name.lower() != "name" else 0) * mm)
        pw = 14 * mm
        name_idx = next((i for i, (_, n) in enumerate(info) if n.lower() == "name"), None)
        rest = avail - sum(widths) - pw * len(pcols)
        if name_idx is not None:
            if rest < 45 * mm:                          # shrink punch cols to keep Name readable
                pw = max(9 * mm, (avail - sum(widths) - 45 * mm) / len(pcols))
                rest = avail - sum(widths) - pw * len(pcols)
            widths[name_idx] = rest
        else:
            pw = (avail - sum(widths)) / len(pcols)
        col_w = widths + [pw] * len(pcols)

        for b in p["blocks"]:
            hdr = [Paragraph(n, head) for _, n in info] + [Paragraph(f"{k}-{n}", head) for n, k in pcols]
            data = [hdr]
            for r in b["rows"]:
                row = []
                for j, n in info:
                    v = clean(r.get(j))
                    row.append(Paragraph(v.replace("&", "&amp;").replace("<", "&lt;"), cell)
                               if n.lower() == "name" else v)
                row += [fmt_time(r.get(p["punch"].get((n, k)))) for n, k in pcols]
                data.append(row)

            t = Table(data, colWidths=col_w, repeatRows=1, rowHeights=None)
            style = [
                ("FONT", (0, 0), (-1, -1), "Helvetica", font_size),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BOX", (0, 0), (-1, -1), 1.3, colors.black),
                ("LINEBELOW", (0, 0), (-1, 0), 1.3, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ]
            sno = next((i for i, (_, n) in enumerate(info) if n.lower() == "sno"), None)
            if sno is not None:
                style.append(("ALIGN", (sno, 1), (sno, -1), "RIGHT"))
            t.setStyle(TableStyle(style))

            label = p["company"] + (f"  ({b['date']})" if len(p["blocks"]) > 1 and b["date"] else "")
            co = Table([[Paragraph("<b>Company:</b>", co_style), Paragraph(label, co_style)]],
                       colWidths=[34 * mm, avail - 34 * mm])
            co.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                                    ("TOPPADDING", (0, 0), (-1, -1), 0)]))
            story.append(KeepTogether([co, t]))
            story.append(Spacer(1, 4 * mm))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #
def sheet_frame(p, pairs) -> pd.DataFrame:
    """Preview table for one sheet."""
    cols = [n for _, n in p["info"]] + [f"{k}-{n}" for n in pairs for k in ("IN", "OUT")]
    rows = []
    for b in p["blocks"]:
        for r in b["rows"]:
            rows.append([clean(r.get(j)) for j, _ in p["info"]]
                        + [fmt_time(r.get(p["punch"].get((n, k)))) for n in pairs for k in ("IN", "OUT")])
    return pd.DataFrame(rows, columns=cols)


def main():
    st.set_page_config(page_title="Punch Report to PDF", page_icon="🖨️", layout="wide")
    st.title("🖨️ Daily Attendance IN/OUT Punch Report → Print PDF")
    st.caption("Upload the punch report workbook (.xls / .xlsx). Every sheet becomes a "
               "\"Company\" block. IN-2/OUT-2, IN-3/OUT-3 … columns appear only when punches exist.")

    with st.sidebar:
        st.header("Print options")
        orientation = st.radio("Page orientation", ["Auto", "Portrait", "Landscape"],
                               help="Auto switches to landscape when more than 3 IN/OUT pairs are needed.")
        same_cols = st.checkbox("Same IN/OUT columns for all companies", value=False,
                                help="Off: each company shows only the pairs it has punches in.")
        font_size = st.slider("Font size", 7.0, 11.0, 8.5, 0.5)

    up = st.file_uploader("Punch report file", type=["xls", "xlsx"])
    if not up:
        st.info("Upload a file to begin.")
        return

    try:
        sheets = read_workbook(up)
    except Exception as e:                                   # noqa: BLE001
        st.error(f"Could not read the file: {e}")
        return
    if not sheets:
        st.error("No sheet with an 'SNo' header row was found.")
        return

    global_pairs = sorted({n for p in sheets for n in used_pairs(p)}) if same_cols else None

    summary, frames = [], []
    for p in sheets:
        pairs = global_pairs or used_pairs(p)
        df = sheet_frame(p, pairs)
        punch_cols = [c for c in df.columns if PUNCH_RE.match(c)]
        punched = int((df[punch_cols] != "").any(axis=1).sum()) if len(df) else 0
        summary.append({"Sheet": p["sheet"], "Company": p["company"], "Employees": len(df),
                        "With punches": punched,
                        "IN/OUT pairs shown": ", ".join(str(n) for n in pairs)})
        frames.append((p, df))

    st.subheader(f"{len(sheets)} sheets found · {sheets[0]['period']}")
    st.dataframe(pd.DataFrame(summary), hide_index=True, use_container_width=True)

    pdf = build_pdf(sheets, orientation, same_cols, font_size)
    st.download_button("⬇️ Download print PDF", pdf,
                       file_name=re.sub(r"\.[^.]+$", "", up.name) + "_print.pdf",
                       mime="application/pdf", type="primary")

    with st.expander("Preview data by company"):
        for p, df in frames:
            st.markdown(f"**{p['company']}** — {p['sheet']}")
            st.dataframe(df, hide_index=True, use_container_width=True)


if __name__ == "__main__":
    main()
