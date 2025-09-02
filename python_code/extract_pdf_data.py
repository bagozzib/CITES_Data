# -*- coding: utf-8 -*-
import os
import re
import argparse
from dataclasses import dataclass
from typing import List, Tuple, Dict, Iterable, Optional

import pdfplumber
import pandas as pd

import pytesseract
from pytesseract import Output as TesseractOutput
from pdf2image import convert_from_path
from PIL import Image


title_match_pattern = re.compile(
    r"(Mr\.\s*|H\.R\.H\.\s*|Mx\.\s*|St\.|Miss\ |Mlle\ |Mine\ |H\.H\.\s*|Ind\.\s*|His\ |Ind\ |Ms\ |Mr\ |Sra\ |Sr\ |M\ |On\ |M\ |Fr\ |H\.O\.\s*|Rev\ |Mme\ |Sr\ |Msgr\ |On\.\s*|Fr\.\s*|Rev\.\s*|H\.E(?:\.\s*(?:Ms\.\s*|Mr\.\s*|Ms\ |Mr\ |Sra\ |Sr\ |Sra\.\s*|Mme|Sr\.\s*|Msgr\.\s*))?|Msgr\.\s*|Mrs\.\s*|Sra\.\s*|Sr\.\s*|Ms\.\s*|Dr\.\s*|Prof\.\s*|M\.\s*|Mme|Ms|S\.E(?:\.\s*(?:Ms\.\s*|Mr\.\s*|Mme|Mr|Ms|Dr|Msgr\.\s*|M\.\s*|Ms\ |Mr\ |Sra\ |Sr\ |M\ |Sra\.\s*|Sr\.\s*))?)"
)


@dataclass
class Record:
    """Structured output row for a participant."""
    Delegation: str
    Honorific: str
    Person_Name: str
    Affiliation: str


def split_honorific_and_person(line: str) -> Tuple[str, str]:
    """Split a line into (honorific, person) using the provided pattern."""
    s = (line or "").strip()
    m = title_match_pattern.match(s)
    if m:
        honorific = m.group().strip()
        person = s[len(m.group()):].strip()
        return honorific, person
    return "", s


def is_all_caps_header(text: str) -> bool:
    """
    Heuristic for delegation headers in CITES PDFs:
    allow only uppercase letters, spaces, and slashes.
    e.g., "BAHAMAS", "SWITZERLAND / SUISSE / SUIZA"
    """
    if not text:
        return False
    s = text.strip()
    return bool(s and all(c.isupper() or c in " /" for c in s))


def group_chars_to_lines(chars: List[Dict], y_tol: float = 3.0) -> List[Tuple[str, float, float, bool]]:
    """
    Cluster characters into lines, reconstruct text, and flag bold.
    Returns list of (text, y0, y1, is_bold).
    """
    if not chars:
        return []
    chars = sorted(chars, key=lambda c: (c.get('top', 0.0), c.get('x0', 0.0)))
    lines: List[Tuple[List[Dict], float, float]] = []
    cur = [chars[0]]
    y0 = y1 = chars[0].get('top', 0.0)
    for c in chars[1:]:
        top = c.get('top', 0.0)
        if abs(top - y1) <= y_tol:
            cur.append(c)
            y1 = top
        else:
            lines.append((cur, y0, y1))
            cur, y0, y1 = [c], top, top
    lines.append((cur, y0, y1))

    out: List[Tuple[str, float, float, bool]] = []
    for line_chars, y0, y1 in lines:
        line_chars.sort(key=lambda c: c.get('x0', 0.0))
        text = ''.join(c.get('text', '') for c in line_chars).strip()
        is_bold = any('Bold' in (c.get('fontname') or '') for c in line_chars)
        out.append((text, y0, y1, is_bold))
    return out


def group_words_to_lines_with_y(words: List[Dict], y_tol: float = 3.0) -> List[Tuple[str, float, float]]:
    """
    Given a list of word dicts (with 'text' and 'top'), cluster into lines by y.
    Returns a list of (text, y0, y1) tuples.
    """
    if not words:
        return []

    words_sorted = sorted(words, key=lambda w: (w.get("top", 0.0), w.get("x0", 0.0)))
    lines = []
    cur = [words_sorted[0]]
    y0 = y1 = words_sorted[0].get("top", 0.0)
    for w in words_sorted[1:]:
        top = w.get("top", 0.0)
        if abs(top - y1) <= y_tol:
            cur.append(w)
            y1 = top
        else:
            cur.sort(key=lambda ww: ww.get("x0", 0.0))
            txt = " ".join(ww.get("text", "") for ww in cur)
            lines.append((txt, y0, y1))
            cur = [w]
            y0 = y1 = top

    cur.sort(key=lambda ww: ww.get("x0", 0.0))
    txt = " ".join(ww.get("text", "") for ww in cur)
    lines.append((txt, y0, y1))

    return lines


def collect_paragraphs_with_y(lines_with_y: List[Tuple[str, float, float]], para_factor: float = 1.5) -> List[Tuple[List[str], float, float]]:
    """
    Given [(text,y0,y1),...], compute the median vertical gap between midpoints,
    then split whenever the gap > median*para_factor.
    Returns list of ([line1, line2, ...], block_y0, block_y1).
    """
    if not lines_with_y:
        return []
    mids = [(y0 + y1) / 2.0 for _, y0, y1 in lines_with_y]
    diffs = [mids[i + 1] - mids[i] for i in range(len(mids) - 1)]
    if not diffs:
        return [( [lines_with_y[0][0]], lines_with_y[0][1], lines_with_y[0][2] )]
    median_gap = sorted(diffs)[len(diffs) // 2]
    thresh = median_gap * para_factor

    paras: List[Tuple[List[str], float, float]] = []
    cur_lines: List[str] = []
    block_y0: Optional[float] = None
    block_y1: Optional[float] = None
    prev_mid: Optional[float] = None

    for (txt, y0, y1), mid in zip(lines_with_y, mids):
        if prev_mid is not None and (mid - prev_mid) > thresh and cur_lines:
            paras.append((cur_lines, block_y0 if block_y0 is not None else y0, block_y1 if block_y1 is not None else y1))
            cur_lines = []
            block_y0 = block_y1 = None

        if not cur_lines:
            block_y0 = y0
        cur_lines.append(txt)
        block_y1 = y1
        prev_mid = mid

    if cur_lines:
        paras.append((cur_lines, block_y0 if block_y0 is not None else 0.0, block_y1 if block_y1 is not None else 0.0))

    return paras


def extract_singlecol_textpdf(pdf_path: str) -> List[Record]:
    """
    Single-column text PDFs (font-based). Delegations are bold lines, persons beneath.
    Affiliation lines continue until the next bold line or an empty line.
    """
    records: List[Record] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = group_chars_to_lines(page.chars, y_tol=3.0)
            i = 0
            current_delegation: Optional[str] = None

            while i < len(lines):
                text, _, _, is_bold = lines[i]

                # Bold lines → delegation headers
                if is_bold and text:
                    current_delegation = text.split("/", 1)[0].strip()
                    i += 1
                    continue

                # Non-bold + have a current delegation → person row
                if current_delegation and text:
                    honorific, person = split_honorific_and_person(text)
                    i += 1

                    # Collect affiliation lines (subsequent non-bold, non-empty)
                    affiliation_lines: List[str] = []
                    while i < len(lines) and not lines[i][3] and lines[i][0].strip():
                        affiliation_lines.append(lines[i][0].strip())
                        i += 1

                    affiliation = " ".join(affiliation_lines).strip()
                    records.append(Record(
                        Delegation=current_delegation,
                        Honorific=honorific,
                        Person_Name=person,
                        Affiliation=affiliation
                    ))
                else:
                    i += 1

    return records


def _twocol_records_from_words(words: List[Dict], x0_thresh: float) -> List[Record]:
    """
    Core two-column extraction on a single page's word list.
    - Detect global delegation headers (all-caps single-line paragraphs).
    - Split into left/right columns by x0_thresh.
    - For each column paragraph: first line is name, rest is affiliation.
    """
    records: List[Record] = []

    # --------- Detect global headers on the whole page ----------
    full_lines = group_words_to_lines_with_y(words, y_tol=3.0)
    full_paras = collect_paragraphs_with_y(full_lines, para_factor=1.5)

    # Build sorted list of (midpoint_y, delegation_name)
    headers: List[Tuple[float, str]] = []
    for blk, y0, y1 in full_paras:
        if len(blk) == 1 and is_all_caps_header(blk[0]):
            nm = blk[0].split("/", 1)[0].strip()
            mid = (y0 + y1) / 2.0
            headers.append((mid, nm))
    headers.sort(key=lambda x: x[0])

    def header_for_mid(mid: float) -> Optional[str]:
        ans = None
        for m, name in headers:
            if m <= mid:
                ans = name
            else:
                break
        return ans

    # --------- Split into left/right columns ----------
    left_words = sorted([w for w in words if w.get("x0", 0.0) < x0_thresh], key=lambda w: (w.get("top", 0.0), w.get("x0", 0.0)))
    right_words = sorted([w for w in words if w.get("x0", 0.0) >= x0_thresh], key=lambda w: (w.get("top", 0.0), w.get("x0", 0.0)))

    for col_words in (left_words, right_words):
        lined = group_words_to_lines_with_y(col_words, y_tol=3.0)
        paras = collect_paragraphs_with_y(lined, para_factor=1.5)

        for blk, y0, y1 in paras:
            # Skip pure headers inside the column
            if len(blk) == 1 and is_all_caps_header(blk[0]):
                continue

            if not blk:
                continue

            name_line = blk[0].strip()
            honorific, person = split_honorific_and_person(name_line)
            affiliation = " ".join(ln.strip() for ln in blk[1:]).strip()

            mid = (y0 + y1) / 2.0
            delegation = header_for_mid(mid) or ""

            records.append(Record(
                Delegation=delegation,
                Honorific=honorific,
                Person_Name=person,
                Affiliation=affiliation
            ))

    return records


def extract_twocol_textpdf(pdf_path: str, x0_thresh: float = 260.0) -> List[Record]:
    """Two-column text PDFs using pdfplumber word geometry."""
    out: List[Record] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words() or []
            if not words:
                continue
            out.extend(_twocol_records_from_words(words, x0_thresh=x0_thresh))
    return out


def ocr_page_to_words(img: Image.Image) -> List[Dict]:
    """
    Run Tesseract OCR with TSV output and convert to a pdfplumber-like words list:
    [{'text': str, 'x0': float, 'top': float}, ...]
    """
    data = pytesseract.image_to_data(img, output_type=TesseractOutput.DICT)
    words: List[Dict] = []
    n = len(data.get("text", []))
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        conf = int(data.get("conf", ["-1"] * n)[i])
        if not txt or conf < 0:
            continue
        left = float(data.get("left", [0] * n)[i])
        top = float(data.get("top", [0] * n)[i])
        # width = float(data.get("width", [0]*n)[i])  # not strictly needed here
        # height = float(data.get("height",[0]*n)[i])
        words.append({"text": txt, "x0": left, "top": top})
    return words


def extract_with_ocr(pdf_path: str, x0_thresh: float = 260.0, dpi: int = 300, poppler_path: Optional[str] = None) -> List[Record]:
    """
    OCR path for image PDFs:
    - Convert pages → images
    - Tesseract TSV → per-word coordinates
    - Reuse the two-column paragraph logic on OCR words
    """
    images = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)
    out: List[Record] = []
    for img in images:
        words = ocr_page_to_words(img)
        if not words:
            continue
        out.extend(_twocol_records_from_words(words, x0_thresh=x0_thresh))
    return out


def detect_layout_quick(pdf_path: str, x0_thresh: float = 260.0) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:2]:  # first up to 2 pages
                words = page.extract_words() or []
                if not words:
                    continue
                left = sum(1 for w in words if w.get("x0", 0.0) < x0_thresh)
                right = sum(1 for w in words if w.get("x0", 0.0) >= x0_thresh)
                total = left + right
                if total == 0:
                    continue
                # if both sides carry at least 25% each, assume 2 columns
                if left / total >= 0.25 and right / total >= 0.25:
                    return "two"
            return "one"
    except Exception:
        # If anything fails, default to one
        return "one"


def extract_cites(
    pdf_path: str,
    layout: str = "auto",         # "auto" | "one" | "two"
    x0_thresh: float = 260.0,
    force_ocr: bool = False,
    tesseract_cmd: Optional[str] = None,
    poppler_path: Optional[str] = None,
    ocr_dpi: int = 300
) -> List[Record]:

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    if force_ocr:
        return extract_with_ocr(pdf_path, x0_thresh=x0_thresh, dpi=ocr_dpi, poppler_path=poppler_path)

    # Try using text-based extraction
    mode = detect_layout_quick(pdf_path, x0_thresh=x0_thresh) if layout == "auto" else layout
    if mode == "two":
        return extract_twocol_textpdf(pdf_path, x0_thresh=x0_thresh)
    else:
        return extract_singlecol_textpdf(pdf_path)


def to_dataframe(records: List[Record]) -> pd.DataFrame:
    """Convert records to a pandas DataFrame with stable column order."""
    return pd.DataFrame(
        [r.__dict__ for r in records],
        columns=["Delegation", "Honorific", "Person_Name", "Affiliation"]
    )


def write_output(df: pd.DataFrame, out_path: str) -> None:
    """Write CSV/XLSX depending on file extension."""
    ext = os.path.splitext(out_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df.to_excel(out_path, index=False)
    else:
        df.to_csv(out_path, index=False, encoding="utf-8-sig")


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Extract CITES COP Participant lists (PDF) → CSV/XLSX",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    p.add_argument("pdf", help="Path to input PDF")
    p.add_argument("-o", "--out", default="cites_participants.csv", help="Output file (.csv or .xlsx)")
    p.add_argument("--layout", choices=["auto", "one", "two"], default="auto", help="Force layout or auto-detect")
    p.add_argument("--x-threshold", type=float, default=260.0, help="Column split threshold (x0)")
    p.add_argument("--force-ocr", action="store_true", help="Force OCR pipeline (image PDFs)")
    p.add_argument("--tesseract-cmd", default=None, help="Path to Tesseract executable")
    p.add_argument("--poppler-path", default=None, help="Path to Poppler binaries (Windows)")
    p.add_argument("--ocr-dpi", type=int, default=300, help="DPI for OCR rasterization")
    return p


def main(argv: Optional[List[str]] = None) -> None:
    args = build_arg_parser().parse_args(argv)

    recs = extract_cites(
        pdf_path=args.pdf,
        layout=args.layout,
        x0_thresh=args.x_threshold,
        force_ocr=args.force_ocr,
        tesseract_cmd=args.tesseract_cmd,
        poppler_path=args.poppler_path,
        ocr_dpi=args.ocr_dpi
    )

    df = to_dataframe(recs)
    write_output(df, args.out)
    print(f"Wrote {len(df)} rows to {args.out}")


if __name__ == "__main__":
    main()

# ***************************************
# Uncomment the below to start Execution:
#
# Example 1: Extracting data from a "Single Column" text PDF file.
# records = extract_cites(
#     pdf_path=r"C:\path\to\CITES_COP_Participants.pdf",
#     layout="one",               # or "auto"
#     x0_thresh=260.0,
#     force_ocr=False,
#     tesseract_cmd=r"C:\Program Files\Tesseract-OCR\tesseract.exe",  # if needed
#     poppler_path=None
# )
# df = to_dataframe(records)
# write_output(df, r"C:\path\to\output.csv")
#
# Example 2: Two-column text PDF.
# records = extract_cites(
#     pdf_path=r"C:\path\to\CITES_COP_Participants_2.pdf",
#     layout="two",               # or "auto" if you want detection
#     x0_thresh=260.0,
#     force_ocr=False
# )
# df = to_dataframe(records)
# write_output(df, r"C:\path\to\output.xlsx")
#
# Example 3: Image PDF (OCR). Works for single or two columns via TSV geometry.
# records = extract_cites(
#     pdf_path=r"C:\path\to\CITES_COP_Participants_cop1.pdf",
#     layout="auto",              # layout is ignored for OCR; geometry is from TSV
#     x0_thresh=260.0,
#     force_ocr=True,
#     tesseract_cmd=r"C:\Program Files\Tesseract-OCR\tesseract.exe",
#     poppler_path=r"C:\path\to\poppler\bin",    # needed on Windows for pdf2image
#     ocr_dpi=300
# )
# df = to_dataframe(records)
# write_output(df, r"C:\path\to\output.csv")
