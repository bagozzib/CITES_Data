import pandas as pd
import numpy as np
from math import radians, cos, sin, asin, sqrt

INPUT_PATH = r"C:\\Users\\emuru\\Downloads\\CITES Extracted Data V2 (3).xlsx"

df = pd.read_excel(INPUT_PATH, engine="openpyxl")

def haversine(lat1, lon1, lat2, lon2):
    """
    Calculate the great circle distance in kilometers between two points
    on the Earth specified by latitude and longitude.
    """
    # convert decimal degrees to radians
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    # haversine formula
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371
    return c * r

# Only compute when both Latitude and Longitude are present; else NaN
df["distance"] = df.apply(
    lambda row: haversine(
        row["COPlatitude"],
        row["COPlongitude"],
        row["Latitude"],
        row["Longitude"]
    ) if pd.notnull(row["Latitude"]) and pd.notnull(row["Longitude"]) else np.nan,
    axis=1
)

df.to_excel(INPUT_PATH, index=False, engine="openpyxl")
print(f"Done—added 'distance' column and saved to {INPUT_PATH}")
