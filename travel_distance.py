"""Smart driving distance calculator for Taiwan business travel routes (OSRM + Nominatim + Tainan landmark cache)."""

from __future__ import annotations

import json
import math
import re
import urllib.parse
import urllib.request
from typing import Any

from ortools.sat.python import cp_model

# Pre-cached coordinates (lat, lon) for common Tainan schools & government offices
KNOWN_PLACES: dict[str, tuple[float, float]] = {
    "臺南市立土城高級中學": (23.0665, 120.1287),
    "土城高中": (23.0665, 120.1287),
    "臺南市政府教育局": (22.9922, 120.1848),
    "教育局": (22.9922, 120.1848),
    "永華市政中心": (22.9922, 120.1848),
    "臺南市政府永華市政中心": (22.9922, 120.1848),
    "民治市政中心": (23.3106, 120.3168),
    "臺南市政府民治市政中心": (23.3106, 120.3168),
    "新營區": (23.3102, 120.3166),
    "高鐵臺南站": (22.9248, 120.2856),
    "高鐵台南站": (22.9248, 120.2856),
    "臺南火車站": (22.9971, 120.2126),
    "台南火車站": (22.9971, 120.2126),
    "新營火車站": (23.3067, 120.3230),
    "南寧高中": (22.9369, 120.1828),
    "臺南市立南寧高級中學": (22.9369, 120.1828),
    "臺南一中": (22.9935, 120.2166),
    "國立臺南第一高級中學": (22.9935, 120.2166),
    "臺南女中": (22.9845, 120.2044),
    "國立臺南女子高級中學": (22.9845, 120.2044),
    "臺南高商": (22.9788, 120.2063),
    "國立臺南高級商業職業學校": (22.9788, 120.2063),
    "新豐高中": (22.9672, 120.2852),
    "國立新豐高級中學": (22.9672, 120.2852),
    "新營高工": (23.3019, 120.3095),
    "國立新營高級工業職業學校": (23.3019, 120.3095),
    "白河商工": (23.3474, 120.4158),
    "國立白河高級商工職業學校": (23.3474, 120.4158),
    "大灣高中": (23.0033, 120.2647),
    "永仁高中": (23.0056, 120.2336),
    "安南國中": (23.0472, 120.1753),
    "海佃國中": (23.0258, 120.1883),
    "和順國中": (23.0588, 120.2194),
    "成功大學": (22.9969, 120.2195),
    "臺南大學": (22.9849, 120.2082),
    "臺南市政府": (22.9922, 120.1848),
    "臺南市立圖書館": (23.0218, 120.2367),
    "臺南文化中心": (22.9744, 120.2215),
    "臺灣歷史博物館": (23.0589, 120.2347),
    "國立臺灣歷史博物館": (23.0589, 120.2347),
}

_GEOCODE_CACHE: dict[str, tuple[float, float]] = dict(KNOWN_PLACES)


def _normalize_place_name(name: str) -> str:
    return name.strip().replace("台", "臺")


def geocode_place(name: str, *, timeout_sec: float = 6.0) -> tuple[float, float]:
    """Resolve a Taiwan place name to (lat, lon) via local cache or OpenStreetMap Nominatim."""
    raw = name.strip()
    if not raw:
        raise ValueError("地點名稱不可為空")

    norm = _normalize_place_name(raw)
    if raw in _GEOCODE_CACHE:
        return _GEOCODE_CACHE[raw]
    if norm in _GEOCODE_CACHE:
        return _GEOCODE_CACHE[norm]

    # Try fuzzy substring match against known landmarks first
    for k, coord in KNOWN_PLACES.items():
        if norm == _normalize_place_name(k):
            _GEOCODE_CACHE[raw] = coord
            return coord

    coord = _geocode_with_nominatim(raw, timeout_sec=timeout_sec)
    if coord is None:
        coord = _geocode_with_photon(raw, timeout_sec=timeout_sec)
    if coord is None:
        raise ValueError(
            f"找不到地點「{raw}」的座標（或地圖服務暫時繁忙），"
            "請嘗試輸入更完整地名（如加上縣市名稱），或直接於表格手動輸入「起點->終點:公里數」"
        )

    _GEOCODE_CACHE[raw] = coord
    return coord


def _geocode_with_nominatim(raw: str, *, timeout_sec: float = 6.0) -> tuple[float, float] | None:
    query = urllib.parse.quote(raw)
    url = f"https://nominatim.openstreetmap.org/search?q={query}&countrycodes=tw&format=json&limit=1"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "TainanSchoolTravelHelper/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data:
            return (float(data[0]["lat"]), float(data[0]["lon"]))
    except Exception:
        return None
    return None


def _geocode_with_photon(raw: str, *, timeout_sec: float = 6.0) -> tuple[float, float] | None:
    """Fallback OSM geocoder (Photon) with Taiwan bounding-box bias when Nominatim 429s."""
    query = urllib.parse.quote(raw)
    # Bias toward Taiwan center (23.7, 120.9)
    url = f"https://photon.komoot.io/api/?q={query}&lat=23.7&lon=120.9&limit=5"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "TainanSchoolTravelHelper/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        features = data.get("features") or []
        for feat in features:
            props = feat.get("properties") or {}
            coords = (feat.get("geometry") or {}).get("coordinates") or []
            if len(coords) >= 2 and props.get("countrycode") == "TW":
                return (float(coords[1]), float(coords[0]))
        if features:
            coords = (features[0].get("geometry") or {}).get("coordinates") or []
            if len(coords) >= 2:
                return (float(coords[1]), float(coords[0]))
    except Exception:
        return None
    return None


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return r * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def _osrm_driving_legs(coords: list[tuple[float, float]], *, timeout_sec: float = 6.0) -> list[float] | None:
    """Query OSRM public routing server for actual driving distance (km) per leg."""
    if len(coords) < 2:
        return []
    coord_str = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
    url = f"https://router.project-osrm.org/route/v1/driving/{coord_str}?overview=false"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "TainanSchoolTravelHelper/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("code") == "Ok" and data.get("routes"):
            legs = data["routes"][0].get("legs", [])
            if len(legs) == len(coords) - 1:
                return [round(float(leg["distance"]) / 1000.0, 1) for leg in legs]
    except Exception:
        return None
    return None


def parse_stops_text(raw: str) -> list[str]:
    """Split 'A -> B -> C' or 'A->B:62.2, B->A:62.2' into a clean list of stops."""
    parts: list[str] = []
    for chunk in re.split(r"(?:->|→|,|，|、|;|；)+", raw):
        cleaned = re.sub(r"\s*[:：]\s*\d+(?:\.\d+)?\s*(?:km|公里)?\s*$", "", chunk, flags=re.IGNORECASE).strip()
        if not cleaned:
            continue
        if not parts or parts[-1] != cleaned:
            parts.append(cleaned)
    return parts


def calculate_route_mileage(
    stops: list[str],
    *,
    round_trip: bool = False,
) -> dict[str, Any]:
    """Calculate leg-by-leg driving distances for a sequence of stops."""
    clean_stops = [s.strip() for s in stops if s.strip()]
    if len(clean_stops) < 2:
        raise ValueError("請至少輸入「起點」與「終點」兩個地點（例如：土城高中 -> 教育局）")

    if round_trip and clean_stops[0] != clean_stops[-1]:
        clean_stops.append(clean_stops[0])

    coords = [geocode_place(s) for s in clean_stops]
    legs_km = _osrm_driving_legs(coords)
    source = "osrm"

    if legs_km is None:
        source = "haversine_fallback"
        legs_km = [
            round(_haversine_km(coords[i][0], coords[i][1], coords[i + 1][0], coords[i + 1][1]) * 1.3, 1)
            for i in range(len(coords) - 1)
        ]

    segments = [
        {
            "origin": clean_stops[i],
            "destination": clean_stops[i + 1],
            "distance_km": legs_km[i],
        }
        for i in range(len(legs_km))
    ]
    total_km = round(sum(legs_km), 1)
    is_loop = len(clean_stops) >= 3 and clean_stops[0] == clean_stops[-1]
    one_way_km = round(total_km / 2.0, 1) if is_loop else total_km
    segments_text = ", ".join(
        f"{seg['origin']}->{seg['destination']}:{seg['distance_km']}" for seg in segments
    )

    return {
        "stops": clean_stops,
        "segments": segments,
        "segments_text": segments_text,
        "total_km": total_km,
        "one_way_km": one_way_km,
        "source": source,
        "cpsat_optimized": False,
    }


def _osrm_distance_matrix(coords: list[tuple[float, float]], *, timeout_sec: float = 6.0) -> list[list[float]] | None:
    """Query OSRM table service for an NxN driving distance matrix in km."""
    if len(coords) < 2:
        return [[0.0]]
    coord_str = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
    url = f"https://router.project-osrm.org/table/v1/driving/{coord_str}?annotations=distance"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "TainanSchoolTravelHelper/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("code") == "Ok" and data.get("distances"):
            raw_mat = data["distances"]
            return [
                [round(float(val or 0.0) / 1000.0, 1) for val in row]
                for row in raw_mat
            ]
    except Exception:
        return None
    return None


def optimize_route_with_cpsat(
    stops: list[str],
    *,
    round_trip: bool = True,
    time_limit_sec: float = 2.0,
) -> dict[str, Any]:
    """Use OR-Tools CP-SAT AddCircuit to find the shortest direct route order (公告 78184 號順路原則)."""
    raw_clean = [s.strip() for s in stops if s.strip()]
    if len(raw_clean) < 2:
        raise ValueError("請至少輸入「起點」與「終點」兩個地點")

    origin = raw_clean[0]
    unique_stops: list[str] = [origin]
    for s in raw_clean[1:]:
        if s not in unique_stops:
            unique_stops.append(s)

    if len(unique_stops) <= 2:
        res = calculate_route_mileage(unique_stops, round_trip=round_trip)
        res["cpsat_optimized"] = True
        return res

    coords = [geocode_place(s) for s in unique_stops]
    matrix_km = _osrm_distance_matrix(coords)
    source = "osrm"
    if matrix_km is None:
        source = "haversine_fallback"
        n_pts = len(coords)
        matrix_km = [
            [
                0.0
                if i == j
                else round(_haversine_km(coords[i][0], coords[i][1], coords[j][0], coords[j][1]) * 1.3, 1)
                for j in range(n_pts)
            ]
            for i in range(n_pts)
        ]

    n = len(unique_stops)
    # Scale to 0.1 km integers for exact CP-SAT minimization
    scaled = [[int(round(matrix_km[i][j] * 10)) for j in range(n)] for i in range(n)]

    model = cp_model.CpModel()
    arcs: list[tuple[int, int, cp_model.IntVar]] = []
    obj_terms: list[cp_model.LinearExpr] = []

    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            lit = model.new_bool_var(f"arc_{i}_{j}")
            arcs.append((i, j, lit))
            cost = 0 if (not round_trip and j == 0) else scaled[i][j]
            obj_terms.append(cost * lit)

    model.add_circuit(arcs)
    model.minimize(sum(obj_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_sec
    status = solver.solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(f"CP-SAT 路線最佳化失敗 (status={solver.status_name(status)})")

    next_node: dict[int, int] = {}
    for i, j, lit in arcs:
        if solver.value(lit) == 1:
            next_node[i] = j

    ordered_indices = [0]
    curr = next_node[0]
    while curr != 0 and len(ordered_indices) < n:
        ordered_indices.append(curr)
        curr = next_node[curr]

    ordered_stops = [unique_stops[idx] for idx in ordered_indices]
    if round_trip:
        ordered_stops.append(origin)
        ordered_indices.append(0)

    segments = [
        {
            "origin": ordered_stops[k],
            "destination": ordered_stops[k + 1],
            "distance_km": matrix_km[ordered_indices[k]][ordered_indices[k + 1]],
        }
        for k in range(len(ordered_stops) - 1)
    ]
    total_km = round(sum(seg["distance_km"] for seg in segments), 1)
    is_loop = len(ordered_stops) >= 3 and ordered_stops[0] == ordered_stops[-1]
    one_way_km = round(total_km / 2.0, 1) if is_loop else total_km
    segments_text = ", ".join(
        f"{seg['origin']}->{seg['destination']}:{seg['distance_km']}" for seg in segments
    )

    return {
        "stops": ordered_stops,
        "segments": segments,
        "segments_text": segments_text,
        "total_km": total_km,
        "one_way_km": one_way_km,
        "source": source,
        "cpsat_optimized": True,
    }
