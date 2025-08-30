import pdfplumber
import pandas as pd

def group_chars_to_lines(chars, y_tol=3):
    """
    Cluster characters into lines by their 'top' coordinate.
    Returns list of (text, y0, y1, is_bold).
    """
    # sort by vertical, then horizontal
    chars = sorted(chars, key=lambda c: (c['top'], c['x0']))
    lines = []
    cur, y0, y1 = [chars[0]], chars[0]['top'], chars[0]['top']
    for c in chars[1:]:
        if abs(c['top'] - y1) <= y_tol:
            cur.append(c)
            y1 = c['top']
        else:
            lines.append((cur, y0, y1))
            cur, y0, y1 = [c], c['top'], c['top']
    lines.append((cur, y0, y1))

    out = []
    for line_chars, y0, y1 in lines:
        # reconstruct full text
        line_chars.sort(key=lambda c: c['x0'])
        text = ''.join(c['text'] for c in line_chars).strip()
        # bold if any char’s fontname contains "Bold"
        is_bold = any('Bold' in c.get('fontname','') for c in line_chars)
        out.append((text, y0, y1, is_bold))
    return out

def extract_delegation_records(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = group_chars_to_lines(page.chars)
            delegation = None
            i = 0
            while i < len(lines):
                text, y0, y1, is_bold = lines[i]

                if is_bold and text:
                    # this line is a delegation header
                    delegation = text.split('/',1)[0].strip()
                    i += 1
                    continue

                # if we have a current delegation and this line is not bold
                if delegation and text:
                    # first non-bold → person name
                    person = text
                    affiliation_lines = []
                    i += 1

                    # collect subsequent non-bold lines as affiliation
                    while i < len(lines) and not lines[i][3] and lines[i][0].strip():
                        affiliation_lines.append(lines[i][0])
                        i += 1

                    affiliation = " ".join(affiliation_lines).strip()
                    records.append({
                        "Delegation": delegation,
                        "Person Name": person,
                        "Affiliation": affiliation
                    })
                else:
                    i += 1

    return records

if __name__ == "__main__":
    df = extract_delegation_records(r"C:\Users\emuru\OneDrive\Desktop\CITES COP Participants\CITES_COP12_Observer_Participants.pdf")
    print(df)
