import pandas as pd
import time
from geopy.geocoders import Nominatim, ArcGIS
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError

XLSX_PATH = r"C:\Users\emuru\Downloads\CITES Extracted Data V2.xlsx"
OUT_PATH  = r"C:\Users\emuru\Downloads\CITES_Data_geocoded.csv"

df = pd.read_excel(XLSX_PATH)

if "latitude" not in df.columns:
    df["latitude"] = None
if "longitude" not in df.columns:
    df["longitude"] = None

nominatim = Nominatim(user_agent="cites_geocoder")
arcgis    = ArcGIS()

def cascade_geocode(addr):
    # Try OSM/Nominatim
    try:
        loc = nominatim.geocode(addr, timeout=10)
    except (GeocoderTimedOut, GeocoderUnavailable):
        loc = None
    if loc:
        return loc.latitude, loc.longitude

    # Fallback ArcGIS (truncate if >300 chars)
    short_addr = addr if len(addr) <= 300 else addr[:300]
    try:
        loc2 = arcgis.geocode(short_addr, timeout=10)
    except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
        loc2 = None
    if loc2:
        return loc2.latitude, loc2.longitude

    return None, None

total = len(df)
for row_num, (idx, raw_addr) in enumerate(df["Affiliation"].fillna("").items(), start=1):
    addr = raw_addr.strip()
    if addr:
        try:
            lat, lng = cascade_geocode(addr)
        except Exception:
            lat, lng = None, None
    else:
        lat, lng = None, None

    df.at[idx, "latitude"]  = lat
    df.at[idx, "longitude"] = lng
    print(f"Row {row_num}/{total}: latitude={lat}, longitude={lng}")

    if row_num % 200 == 0:
        df.to_csv(OUT_PATH, index=False)
        print(f"Checkpoint saved up to row {row_num}")

    time.sleep(0.03)

df.to_csv(OUT_PATH, index=False)
print(f"All done—geocoded data saved to {OUT_PATH}")
