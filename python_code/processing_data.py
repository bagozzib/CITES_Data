import pandas as pd
import re

df = pd.read_csv(r"C:\Users\emuru\Downloads\CITES Extracted Data - CITES Data - Sheet1.csv")
print(df.columns)

def clean_text(value):
    if pd.isna(value):
        return value

    # 1) Remove any sequence of slash-separated "translations", keeping only the first word
    #    e.g. "Belgium/Bélgica/Belgique" → "Belgium"
    value = re.sub(r'(\b[^/\s]+)(?:/[^/\s]+)+', r'\1', value)

    # 2) If there's still a slash (e.g. "Dept A/Dept B"), drop everything after the first slash
    if '/' in value:
        value = value.split('/', 1)[0]

    # 3) If we opened a parenthesis but dropped its closing “)” by trimming, add it back
    if '(' in value and ')' not in value:
        value += ')'

    return value

# Apply to both columns
df['Delegation'] = df['Delegation'].apply(clean_text)
df['Affiliation'] = df['Affiliation'].apply(clean_text)

# Save result to Excel
df.to_excel(r"C:\Users\emuru\Downloads\cleaned_data__1.xlsx", index=False)

print("Saved cleaned_data.xlsx")
