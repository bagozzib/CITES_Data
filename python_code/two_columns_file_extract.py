import pdfplumber
from inputs_file import title_match_pattern
import pandas as pd

def group_words_to_lines_with_y(words, y_tol=3):
    """
    Given a list of word‐dicts (with 'text' and 'top'),
    cluster into lines by y‐coordinate. Returns a list of
    (text, y0, y1) tuples where y0/y1 are the min/max 'top'.
    """
    if not words:
        return []
    lines = []
    cur = [words[0]]
    y0 = y1 = words[0]["top"]
    for w in words[1:]:
        if abs(w["top"] - y1) <= y_tol:
            cur.append(w)
            y1 = w["top"]
        else:
            # flush current
            cur.sort(key=lambda w: w["x0"])
            txt = " ".join(w["text"] for w in cur)
            lines.append((txt, y0, y1))
            cur = [w]
            y0 = y1 = w["top"]

    cur.sort(key=lambda w: w["x0"])
    txt = " ".join(w["text"] for w in cur)
    lines.append((txt, y0, y1))

    return lines

def collect_paragraphs_with_y(lines_with_y, para_factor=1.5):
    """
    Given [(text,y0,y1),...], compute the median vertical gap
    between *midpoints*, then split whenever the gap > median*para_factor.
    Returns list of ( [line1, line2, ...], block_y0, block_y1 ).
    """
    if not lines_with_y:
        return []
    mids = [(y0+y1)/2 for _,y0,y1 in lines_with_y ]
    diffs = [ mids[i+1]-mids[i] for i in range(len(mids)-1) ]
    median_gap = sorted(diffs)[len(diffs)//2] if diffs else 0
    thresh = median_gap * para_factor

    paras = []
    cur_lines = []
    block_y0 = block_y1 = None
    prev_mid = None

    for (txt,y0,y1), mid in zip(lines_with_y, mids):
        if prev_mid is not None and (mid - prev_mid) > thresh and cur_lines:
            paras.append((cur_lines, block_y0, block_y1))
            cur_lines = []
            block_y0 = block_y1 = None

        if not cur_lines:
            block_y0 = y0
        cur_lines.append(txt)
        block_y1 = y1
        prev_mid = mid

    if cur_lines:
        paras.append((cur_lines, block_y0, block_y1))

    return paras

def is_delegation(txt):
    """
    True if txt is all‐caps letters, spaces or slashes,
    e.g. "ARGENTINA/ARGENTINE" or "INTERNATIONAL CAT CONSERVATION COMMITTEE"
    """
    s = txt.strip()

    return bool(s and all(c.isupper() or c in " /" for c in s))

def extract_by_blank_lines(pdf_path, x0_thresh=260):
    """
    1) Scan full‐page words → global paragraphs → pick all Delegation headers.
    2) Re‐split page into left/right columns by x0_thresh.
    3) For each left/right paragraph, assign it the *most recent* header above it.
    """
    records = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()

            # 1) detect global headers anywhere on the page ---
            full_lines = sorted(words, key=lambda w: (w["top"], w["x0"]))
            full_lined = group_words_to_lines_with_y(full_lines)
            full_paras = collect_paragraphs_with_y(full_lined)

            # build a sorted list of (midpoint_y, delegation_name)
            headers = []
            for blk, y0, y1 in full_paras:
                if len(blk)==1 and is_delegation(blk[0]):
                    name = blk[0].split("/",1)[0].strip()
                    mid = (y0+y1)/2
                    headers.append((mid, name))
            # sort by y so we can binary‐search by block midpoint
            headers.sort(key=lambda x: x[0])

            # helper: for any block_mid, find the last header whose mid <= block_mid
            def find_delegation(block_mid):
                ans = None
                for mid,name in headers:
                    if mid <= block_mid:
                        ans = name
                    else:
                        break
                return ans

            # 2) split into left/right columns & paragraphs ---
            left_words  = sorted([w for w in words if w["x0"] <  x0_thresh],
                                 key=lambda w: (w["top"], w["x0"]))
            right_words = sorted([w for w in words if w["x0"] >= x0_thresh],
                                 key=lambda w: (w["top"], w["x0"]))

            left_lined  = group_words_to_lines_with_y(left_words)
            right_lined = group_words_to_lines_with_y(right_words)

            left_paras  = collect_paragraphs_with_y(left_lined)
            right_paras = collect_paragraphs_with_y(right_lined)

            # 3) for each column‐block, emit a record under its delegation
            for paras in (left_paras, right_paras):
                for blk, y0, y1 in paras:
                    # skip pure headers in column
                    if len(blk)==1 and is_delegation(blk[0]):
                        continue

                    # name + optional honorific
                    name_line = blk[0].strip()
                    m = title_match_pattern.match(name_line)
                    if m:
                        honorific = m.group().strip()
                        person    = name_line[len(m.group()):].strip()
                    else:
                        honorific = ""
                        person    = name_line

                    # grab the affiliation
                    affiliation = " ".join(blk[1:]).strip()

                    # assign the delegation by block midpoint
                    block_mid = (y0+y1)/2
                    delegation = find_delegation(block_mid)

                    records.append({
                        "Delegation":   delegation,
                        "Honorific":    honorific,
                        "Person Name":  person,
                        "Affiliation":  affiliation
                    })

    return records

if __name__ == "__main__":
    pdf_path = r"C:\Users\emuru\OneDrive\Desktop\CITES COP Participants\CITES_COP11_Participants.pdf"
    data = extract_by_blank_lines(pdf_path)

    print(data)
    df = pd.DataFrame(data, columns=["Delegation","Honorific","Person Name","Affiliation"])
    out_xlsx = r"C:\Users\emuru\Downloads\CITES_COP2_output.xlsx"
    df.to_excel(out_xlsx, index=False)
    print(f"Wrote {len(df)} rows to {out_xlsx}")
