import pandas as pd

def standardize_comma(name: str) -> str:
    last, rest = [part.strip() for part in name.split(",", 1)]
    return f"{rest} {last}"

def standardize_two_word_caps(name: str) -> str:
    last, first = name.split()
    return f"{first} {last}"

def standardize_multi_word_caps(name: str) -> str:
    parts = name.split()
    # collect all leading ALL‑CAPS tokens as the surname
    surname_tokens = []
    for token in parts:
        if token.isupper():
            surname_tokens.append(token)
        else:
            break
    rest_tokens = parts[len(surname_tokens):]
    # only re‑order if we found at least one surname token and some rest
    if surname_tokens and rest_tokens:
        return " ".join(rest_tokens + surname_tokens)
    return name

def format_person_names(df, col, new_col, start_row, end_row):
    df[new_col] = df.get(new_col, pd.NA)
    for row in range(start_row, end_row + 1):
        orig = df.at[row-1, col]
        if not isinstance(orig, str) or not orig.strip():
            df.at[row-1, new_col] = pd.NA
            continue

        name = orig.strip()
        parts = name.split()

        if "," in name:
            df.at[row-1, new_col] = standardize_comma(name)

        elif len(parts) == 2 and parts[0].isupper() and any(c.islower() for c in parts[1]):
            df.at[row-1, new_col] = standardize_two_word_caps(name)

        elif len(parts) > 2 and parts[0].isupper() and any(any(c.islower() for c in p) for p in parts[1:]):
            df.at[row-1, new_col] = standardize_multi_word_caps(name)

        else:
            df.at[row-1, new_col] = name

df = pd.read_excel(r"C:\Users\emuru\Downloads\CITES Extracted Data V2.xlsx")
format_person_names(df, col="Person Name", new_col="Formatted Person Name",
                    start_row=1, end_row=len(df))
df.to_excel(r"C:\Users\emuru\Downloads\CITES_Data_standardized_3.xlsx", index=False)
