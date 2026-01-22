# -*- coding: utf-8 -*-
"""
Hashing person names by using a private salt file 
"""

import argparse
import hashlib
import os
import secrets
import unicodedata
from typing import Optional, Dict
import pandas as pd


def ensure_parent_dir(path: str) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def read_or_create_salt(salt_file: str) -> str:
    """ Reads salt from file"""
    ensure_parent_dir(salt_file)
    if os.path.exists(salt_file):
        with open(salt_file, "r", encoding="utf-8") as f:
            salt = f.read().strip()
            if salt:
                return salt

    # Create a new salt
    salt = secrets.token_hex(16)
    with open(salt_file, "w", encoding="utf-8") as f:
        f.write(salt)
    return salt


def normalize_text(s: str) -> str:
    # Unicode normalize + trim + collapse spaces
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    s = " ".join(s.split())
    return s


def standardize_person_name(person: Optional[str]) -> Optional[str]:
    """
    Standardize PersonName for hashing (Honorific excluded by design: we only use PersonName).
    """
    if person is None:
        return None
    person = str(person).strip()
    if person == "" or person.upper() == "NA":
        return None

    person = normalize_text(person)

    # Light cleanup (keep letters/diacritics; just remove repeated punctuation spacing)
    person = person.replace(" ,", ",").replace(" .", ".")
    person = person.strip(" ,;")

    return person if person else None


def make_numeric_hash(name_std: str, salt: str) -> str:
    """
    Stable numeric hash (as string).
    We keep it as TEXT (string) so Excel doesn't round it.
    """
    payload = (salt + "|" + name_std).encode("utf-8", errors="ignore")
    digest = hashlib.sha256(payload).digest()

    # Use 12 bytes (~96-bit) => up to 29 digits; stored as string
    num = int.from_bytes(digest[:12], byteorder="big", signed=False)
    return str(num)


def build_person_id_map(names_std, salt: str) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    used: Dict[str, str] = {}  

    for n in sorted(set([x for x in names_std if isinstance(x, str) and x.strip()])):
        pid = make_numeric_hash(n, salt)
        if pid in used and used[pid] != n:
            payload = (salt + "|" + n).encode("utf-8", errors="ignore")
            digest = hashlib.sha256(payload).digest()
            pid2 = str(int.from_bytes(digest[:16], "big", signed=False))
            pid = pid2

        mapping[n] = pid
        used[pid] = n

    return mapping


def read_input(in_csv: Optional[str], in_xlsx: Optional[str]) -> pd.DataFrame:
    if in_xlsx:
        return pd.read_excel(in_xlsx, dtype=str, engine="openpyxl")
    return pd.read_csv(in_csv, dtype=str, encoding="utf-8-sig")


def write_outputs(df: pd.DataFrame, out_csv: Optional[str], out_xlsx: Optional[str]) -> None:
    if out_csv:
        ensure_parent_dir(out_csv)
        df.to_csv(out_csv, index=False, encoding="utf-8-sig", na_rep="NA")
    if out_xlsx:
        ensure_parent_dir(out_xlsx)
        # IMPORTANT: close the Excel file if it's open, or Windows will throw PermissionError.
        df.to_excel(out_xlsx, index=False, na_rep="NA", engine="openpyxl")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in_csv", default=None, help="Input CSV")
    ap.add_argument("--in_xlsx", default=None, help="Input XLSX (use this for your GEOPRIVACY_1.xlsx)")
    ap.add_argument("--out_internal_csv", required=True)
    ap.add_argument("--out_internal_xlsx", required=True)
    ap.add_argument("--out_public_csv", required=True)
    ap.add_argument("--out_public_xlsx", required=True)
    ap.add_argument("--salt_file", default="CITES_Data/master data/PRIVATE_SALT.txt")
    ap.add_argument("--drop_honorific_public", action="store_true", help="Recommended: drop Honorific in public")
    args = ap.parse_args()

    if not args.in_csv and not args.in_xlsx:
        raise SystemExit("Provide either --in_csv or --in_xlsx")

    df = read_input(args.in_csv, args.in_xlsx)

    if "PersonName" not in df.columns:
        raise SystemExit("Input file is missing required column: PersonName")

    # Add stable row_id (based on current row order)
    df = df.copy()
    df.insert(0, "row_id", range(1, len(df) + 1))

    # Standardize names for hashing (Honorific excluded automatically)
    df["PersonName_standardized"] = df["PersonName"].apply(standardize_person_name)

    # Salt (auto-create if missing)
    salt = read_or_create_salt(args.salt_file)

    # Build person_id mapping (stable within this run + future runs if salt stays same)
    name_map = build_person_id_map(df["PersonName_standardized"].dropna().tolist(), salt)

    def lookup_pid(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "NA"
        x = str(x).strip()
        if x == "" or x.upper() == "NA":
            return "NA"
        return name_map.get(x, "NA")

    df["person_id"] = df["PersonName_standardized"].apply(lookup_pid)
    df_internal = df.copy()
    cols = list(df_internal.columns)
    if "PersonName" in cols:
        idx = cols.index("PersonName") + 1
        for c in ["PersonName_standardized", "person_id"]:
            if c in cols:
                cols.remove(c)
        cols[idx:idx] = ["PersonName_standardized", "person_id"]
        df_internal = df_internal[cols]


    df_public = df.copy()
    if "PersonName" in df_public.columns:
        df_public = df_public.drop(columns=["PersonName"])
    if args.drop_honorific_public and "Honorific" in df_public.columns:
        df_public = df_public.drop(columns=["Honorific"])
    if "PersonName_standardized" in df_public.columns:
        df_public = df_public.drop(columns=["PersonName_standardized"])

    pub_cols = list(df_public.columns)
    if "person_id" in pub_cols:
        pub_cols.remove("person_id")

    insert_at = 1  # after row_id
    for anchor in ["URL", "Honorific", "Delegation", "Status"]:
        if anchor in df_public.columns:
            insert_at = df_public.columns.get_loc(anchor) + 1

    pub_cols.insert(insert_at, "person_id")
    df_public = df_public[pub_cols]

    # Write outputs
    write_outputs(df_internal, args.out_internal_csv, args.out_internal_xlsx)
    write_outputs(df_public, args.out_public_csv, args.out_public_xlsx)

    print("DONE")
    print("INTERNAL:", args.out_internal_csv, "and", args.out_internal_xlsx)
    print("PUBLIC  :", args.out_public_csv, "and", args.out_public_xlsx)
    print("SALT (KEEP PRIVATE):", args.salt_file)


if __name__ == "__main__":
    main()
