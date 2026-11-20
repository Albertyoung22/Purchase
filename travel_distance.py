"""Smart driving distance calculator for Taiwan business travel routes (OSRM + Nominatim + Tainan landmark cache)."""

from __future__ import annotations

import json
import math
import re
import urllib.parse
import urllib.request
from typing import Any

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

    query = urllib.parse.quote(raw)
    url = f"https://nominatim.openstreetmap.org/search?q={query}&countrycodes=tw&format=json&limit=1"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "TainanSchoolTravelHelper/1.0"},
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    if not data:
        raise ValueError(f"找不到地點「{raw}」的座標，請嘗試輸入更完整地名（如加上「臺南市」）")

    coord = (float(data[0]["lat"]), float(data[0]["lon"]))
    _GEOCODE_CACHE[raw] = coord
    return coord


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
    """Split 'A -> B -> C' or 'A, B, C' into a clean list of stops."""
    parts = [p.strip() for p in re.split(r"(?:->|→|,|，|、|;|；)+", raw) if p.strip()]
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
    }
