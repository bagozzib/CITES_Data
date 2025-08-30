# -*- coding: utf-8 -*-

import os, math, time, json, tempfile
import pandas as pd
from geopy.geocoders import Nominatim, ArcGIS
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError
import pycountry

INPUT_PATH = r"C:\Users\emuru\Downloads\CITES Extracted Data V2 (5).xlsx"
BASE_DIR   = os.path.dirname(INPUT_PATH) or "."
BASE_NAME  = os.path.splitext(os.path.basename(INPUT_PATH))[0]
OUT_PATH   = os.path.join(BASE_DIR, f"{BASE_NAME}_ccflag_coords.csv")

# Cache files (auto-created)
RC_CACHE_PATH       = os.path.join(BASE_DIR, f"{BASE_NAME}__reverse_country_cache.json")
CENTROID_CACHE_PATH = os.path.join(BASE_DIR, f"{BASE_NAME}__country_centroids_cache.json")

EXACT_DECIMALS      = 5
ADAPTIVE_RATIO_BASE = 0.20
NEAR_MIN_KM         = 12.0
NEAR_MAX_KM         = 50.0
MICRO_DIAG_KM       = 100.0   # tighten for microstates
MICRO_MIN_KM        = 5.0
GIANT_DIAG_KM       = 4000.0  # relax a bit for giants
GIANT_RATIO         = 0.25
GIANT_MAX_KM        = 60.0

REQUIRE_DUAL          = True   # if both provider centroids exist, must be near BOTH
DISAGREE_KM           = 100.0  # if providers are >100 km apart, require exact match
OVERSEAS_KM           = 1500.0 # if both distances exist and min > this, don't flag
CHECKPOINT_EVERY      = 200
PRINT_AFF_LEN         = 120

THROTTLE_ARC_REVERSE  = 0.0
THROTTLE_ARC_FORWARD  = 0.2
THROTTLE_NOM_REVERSE  = 0
THROTTLE_NOM_FORWARD  = 0

print("[IO] Reading:", INPUT_PATH)
df = pd.read_excel(INPUT_PATH, engine="openpyxl")
print(f"[IO] Loaded: {len(df)} rows, {len(df.columns)} cols")

# Column resolution (case-insensitive for lat/lon)
cols_lower = {c.lower(): c for c in df.columns}
lat_col = cols_lower.get("latitude") or "Latitude"
lon_col = cols_lower.get("longitude") or "Longitude"
if lat_col not in df.columns or lon_col not in df.columns:
    raise ValueError("Missing Latitude/Longitude columns.")
if "Affiliation" not in df.columns:
    df["Affiliation"] = ""  # optional

# Numeric lat/lon
df[lat_col] = pd.to_numeric(df[lat_col], errors="coerce")
df[lon_col] = pd.to_numeric(df[lon_col], errors="coerce")

# ========= Geocoders =========
nominatim = Nominatim(user_agent="cites_cc_by_coords_robust")
arcgis    = ArcGIS()

def _load_json(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_json(obj, path):
    tmp = tempfile.NamedTemporaryFile(mode="w", delete=False, dir=os.path.dirname(path) or ".",
                                      suffix=".part", encoding="utf-8")
    try:
        json.dump(obj, tmp, ensure_ascii=False)
        tmp_path = tmp.name
    finally:
        tmp.close()
    os.replace(tmp_path, path)

rc_cache = _load_json(RC_CACHE_PATH)
centroid_cache = _load_json(CENTROID_CACHE_PATH)

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0088
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = p2 - p1
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp/2.0)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2.0)**2
    return 2.0 * R * math.asin(math.sqrt(a))

def bbox_diag_km(bbox):
    try:
        latS, latN, lonW, lonE = map(float, bbox)
        return haversine_km(latS, lonW, latN, lonE)
    except Exception:
        return None

def near_radius_km(bbox):
    # Adaptive radius v2
    diag = bbox_diag_km(bbox) if bbox else None
    if diag is None:
        return NEAR_MAX_KM
    if diag < MICRO_DIAG_KM:
        return max(MICRO_MIN_KM, ADAPTIVE_RATIO_BASE * diag)
    if diag > GIANT_DIAG_KM:
        return min(GIANT_MAX_KM, GIANT_RATIO * diag)
    return min(NEAR_MAX_KM, max(NEAR_MIN_KM, ADAPTIVE_RATIO_BASE * diag))

def key_from_coords(lat, lon, ndp=4):
    return f"{round(lat, ndp)},{round(lon, ndp)}"

def country_from_coords(lat, lon):
    key = key_from_coords(lat, lon)
    cached = rc_cache.get(key)
    if cached is not None:
        return cached.get("country"), cached.get("source") or "cache"

    country_name = None
    source = "none"

    # 1) ArcGIS reverse
    try:
        loc = arcgis.reverse((lat, lon), timeout=20)
        if loc and getattr(loc, "raw", None):
            addr = loc.raw.get("address", {}) or {}
            code = addr.get("CountryCode") or addr.get("CountryCode2") or addr.get("Country")
            if code and isinstance(code, str):
                code = code.strip().upper()
                name = None
                if len(code) == 2:
                    rec = pycountry.countries.get(alpha_2=code)
                    if rec: name = rec.name
                if not name and len(code) == 3:
                    rec = pycountry.countries.get(alpha_3=code)
                    if rec: name = rec.name
                if name:
                    country_name = name
                    source = "arcgis"
        time.sleep(THROTTLE_ARC_REVERSE)
    except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
        pass

    # 2) Fallback: Nominatim reverse
    if not country_name:
        try:
            loc = nominatim.reverse((lat, lon), zoom=5, addressdetails=True, timeout=20)
            if loc and getattr(loc, "raw", None):
                addr = loc.raw.get("address", {}) or {}
                c = addr.get("country")
                if c:
                    country_name = c
                    source = "nominatim"
            time.sleep(THROTTLE_NOM_REVERSE)
        except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
            pass

    rc_cache[key] = {"country": country_name, "source": source}
    return country_name, source

def get_country_centroids(country_name):
    if not country_name:
        return None
    cached = centroid_cache.get(country_name)
    if cached:
        return cached

    nom_lat = nom_lon = None
    bbox = None
    arc_lat = arc_lon = None

    try:
        cand = nominatim.geocode(f"country {country_name}", addressdetails=True, timeout=20)
        if not cand:
            cand = nominatim.geocode(country_name, addressdetails=True, timeout=20)
        if cand:
            nom_lat, nom_lon = cand.latitude, cand.longitude
            raw = cand.raw or {}
            bbox = raw.get("boundingbox")
        time.sleep(THROTTLE_NOM_FORWARD)
    except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
        pass

    # ArcGIS forward
    try:
        cand = arcgis.geocode(country_name, timeout=20)
        if cand:
            arc_lat, arc_lon = cand.latitude, cand.longitude
        time.sleep(THROTTLE_ARC_FORWARD)
    except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
        pass

    near_km = near_radius_km(bbox)
    result = {
        "nom_lat": nom_lat, "nom_lon": nom_lon,
        "arc_lat": arc_lat, "arc_lon": arc_lon,
        "bbox": bbox, "near_km": near_km
    }
    centroid_cache[country_name] = result
    print(f"[COUNTRY] {country_name}: NOM=({nom_lat},{nom_lon}) ARC=({arc_lat},{arc_lon}) near_km={near_km:.1f}")
    # Persist immediately so restarts reuse it
    _save_json(centroid_cache, CENTROID_CACHE_PATH)
    return result

def atomic_save_csv(frame, path):
    out = frame.copy()
    if "is_country_centroid" in out.columns:
        out["is_country_centroid"] = out["is_country_centroid"].fillna(0).astype("int64")
    tmp = tempfile.NamedTemporaryFile(mode="w", delete=False, dir=os.path.dirname(path) or ".",
                                      suffix=".part", newline="", encoding="utf-8")
    try:
        out.to_csv(tmp.name, index=False)
        tmp_path = tmp.name
    finally:
        tmp.close()
    os.replace(tmp_path, path)

# Ensure output columns
for c in [
    "country_from_coords","country_from_coords_source","near_km_used",
    "nom_cc_lat","nom_cc_lon","dist_to_nom_cc_km",
    "arc_cc_lat","arc_cc_lon","dist_to_arc_cc_km",
    "centroid_disagreement_km","providers_available",
    "is_country_centroid","cc_reason"
]:
    if c not in df.columns:
        df[c] = pd.NA

total = len(df)
processed = 0

for i, row in df.iterrows():
    lat = row[lat_col]; lon = row[lon_col]
    aff = row.get("Affiliation", "")

    # Missing coords -> not centroid
    if pd.isna(lat) or pd.isna(lon):
        df.at[i, "country_from_coords"] = pd.NA
        df.at[i, "country_from_coords_source"] = pd.NA
        df.at[i, "is_country_centroid"] = 0
        df.at[i, "cc_reason"] = "no_coords"
        print(f"[ROW {i+1}/{total}] Affiliation-> {str(aff)[:PRINT_AFF_LEN]} | coords missing -> flag=0")
        processed += 1
        if processed % CHECKPOINT_EVERY == 0:
            atomic_save_csv(df, OUT_PATH)
            _save_json(rc_cache, RC_CACHE_PATH)
            print(f"[CKPT] saved -> {OUT_PATH}")
        continue

    lat = float(lat); lon = float(lon)
    # Reverse country
    ctry, src = country_from_coords(lat, lon)
    df.at[i, "country_from_coords"] = ctry if ctry else pd.NA
    df.at[i, "country_from_coords_source"] = src

    if not ctry:
        df.at[i, "is_country_centroid"] = 0
        df.at[i, "cc_reason"] = "no_country_from_coords"
        print(f"[ROW {i+1}/{total}] Affiliation-> {str(aff)[:PRINT_AFF_LEN]} | country NA -> flag=0")
        processed += 1
        if processed % CHECKPOINT_EVERY == 0:
            atomic_save_csv(df, OUT_PATH)
            _save_json(rc_cache, RC_CACHE_PATH)
            print(f"[CKPT] saved -> {OUT_PATH}")
        continue

    # Country centroids
    params = get_country_centroids(ctry)
    near_km = params["near_km"]
    df.at[i, "near_km_used"] = round(near_km, 1)

    nom_lat, nom_lon = params["nom_lat"], params["nom_lon"]
    arc_lat, arc_lon = params["arc_lat"], params["arc_lon"]

    df.at[i, "nom_cc_lat"] = nom_lat if nom_lat is not None else pd.NA
    df.at[i, "nom_cc_lon"] = nom_lon if nom_lon is not None else pd.NA
    df.at[i, "arc_cc_lat"] = arc_lat if arc_lat is not None else pd.NA
    df.at[i, "arc_cc_lon"] = arc_lon if arc_lon is not None else pd.NA

    # Providers available
    provs = int(nom_lat is not None and nom_lon is not None) + int(arc_lat is not None and arc_lon is not None)
    df.at[i, "providers_available"] = provs

    # Distances
    def dist(a, b, c, d):
        return None if (c is None or d is None) else haversine_km(a, b, c, d)

    d_nom = dist(lat, lon, nom_lat, nom_lon)
    d_arc = dist(lat, lon, arc_lat, arc_lon)
    df.at[i, "dist_to_nom_cc_km"] = round(d_nom, 3) if isinstance(d_nom, (int, float)) else pd.NA
    df.at[i, "dist_to_arc_cc_km"] = round(d_arc, 3) if isinstance(d_arc, (int, float)) else pd.NA

    disagree_km = None
    if (nom_lat is not None and nom_lon is not None) and (arc_lat is not None and arc_lon is not None):
        disagree_km = haversine_km(nom_lat, nom_lon, arc_lat, arc_lon)
    df.at[i, "centroid_disagreement_km"] = round(disagree_km, 1) if isinstance(disagree_km, (int, float)) else pd.NA

    # Exact & proximity checks
    exact_nom = int(nom_lat is not None and nom_lon is not None and
                    round(lat, EXACT_DECIMALS)==round(nom_lat, EXACT_DECIMALS) and
                    round(lon, EXACT_DECIMALS)==round(nom_lon, EXACT_DECIMALS))
    exact_arc = int(arc_lat is not None and arc_lon is not None and
                    round(lat, EXACT_DECIMALS)==round(arc_lat, EXACT_DECIMALS) and
                    round(lon, EXACT_DECIMALS)==round(arc_lon, EXACT_DECIMALS))

    near_nom = int(d_nom is not None and d_nom <= near_km)
    near_arc = int(d_arc is not None and d_arc <= near_km)

    reasons = []
    flag = 0

    # Overseas territory safeguard
    if d_nom is not None and d_arc is not None and min(d_nom, d_arc) > OVERSEAS_KM:
        flag = 0
        reasons.append("overseas_territory_far_from_centroid")
    else:
        # Disagreement guard
        require_exact_here = False
        if isinstance(disagree_km, (int, float)) and disagree_km > DISAGREE_KM:
            require_exact_here = True
            reasons.append("providers_disagree_gt_threshold")

        # Decide proximity
        if provs >= 2 and REQUIRE_DUAL:
            proximity_ok = bool(near_nom and near_arc)
        else:
            proximity_ok = bool(near_nom or near_arc)

        # Final decision
        if (exact_nom or exact_arc):
            flag = 1
            reasons.append("exact_centroid")
        elif require_exact_here:
            flag = 0
        elif proximity_ok:
            flag = 1
            reasons.append("dual_proximity_ok" if (provs >= 2 and REQUIRE_DUAL) else "provider_proximity_ok")
        else:
            flag = 0

    df.at[i, "is_country_centroid"] = flag
    df.at[i, "cc_reason"] = "+".join(reasons) if reasons else "not_centroid"

    print(f"[ROW {i+1}/{total}]")

    processed += 1
    if processed % CHECKPOINT_EVERY == 0:
        atomic_save_csv(df, OUT_PATH)
        _save_json(rc_cache, RC_CACHE_PATH)
        print(f"[CKPT] saved -> {OUT_PATH}")

# Final save + cache persist
atomic_save_csv(df, OUT_PATH)
_save_json(rc_cache, RC_CACHE_PATH)
_save_json(centroid_cache, CENTROID_CACHE_PATH)
print("[DONE] saved ->", OUT_PATH)
