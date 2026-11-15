#!/usr/bin/env python3
"""Kuro kainų robotas „Mano garažui“.

Kasdien parsisiunčia oficialų Lietuvos energetikos agentūros (LEA) degalų kainų
Excel failą, paima naujausias kiekvienos degalinės kainas, suranda degalinių
koordinates (OpenStreetMap, rezultatai įsimenami) ir išveda `fuel.json` tokiu
formatu, kokį supranta programėlė (žr. app/fuel-snapshot.json).

Paleidimas:
    python robotas.py --out out/fuel.json --cache out/geocache.json
    python robotas.py --xlsx DK-2026.xlsx --no-geo --out fuel.json   # be interneto
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import math
import re
import sys
import time
import unicodedata
from pathlib import Path
from typing import Callable, Iterable

LEA_PAGE = "https://www.ena.lt/dk-pr-pr-duomenys/"
NOMINATIM = "https://nominatim.openstreetmap.org/search"
OVERPASS = "https://overpass-api.de/api/interpreter"
USER_AGENT = "ManoGarazas-kuro-robotas/1.0 (+https://github.com/andrius3213/kuro-kainos)"

FUELS = {"95 benzinas": 0, "dyzelinas": 1, "snd": 2}  # indeksai: 95, dyzelinas, dujos (LPG)
KEEP_DAYS = 45          # degalinė įtraukiama, jei kainas teikė per paskutines N dienų
RETRY_NONE_DAYS = 30    # nerastus adresus bandyti iš naujo po N dienų
MIN_STATIONS = 300      # saugiklis: mažiau – tikriausiai sugedo failas

BRANDS = {  # teisinis pavadinimas (be UAB/AB) -> trumpas, žinomas pavadinimas
    "orlen baltics retail": "Orlen",
    "orlen lietuva": "Orlen",
    "circle k lietuva": "Circle K",
    "viada lt": "Viada",
    "neste lietuva": "Neste",
    "baltic petroleum": "Baltic Petroleum",
    "virsi lietuva": "Virši",
}


# ---------------------------------------------------------------- pagalbinės
def fold(s: str) -> str:
    """Mažosios raidės be diakritikų – palyginimui."""
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


def clean(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def brand_name(legal: str) -> str:
    """„AB Orlen Baltics Retail“ -> „Orlen“, „UAB Vimijula“ -> „Vimijula“."""
    name = clean(legal).replace("„", "").replace("“", "").replace('"', "")
    name = re.sub(r"^(UAB|AB|ŽŪB|ŽŪK|KB|MB|VšĮ)\s+", "", name)
    name = re.sub(r"\s+(UAB|AB)$", "", name)
    return BRANDS.get(fold(name), name) or clean(legal)


def split_address(addr: str) -> tuple[str, str]:
    """„Ručiūnų k., Kauno g. 2A, 55355“ -> („Ručiūnų k., Kauno g. 2A“, „55355“)."""
    a = clean(addr).strip(", ")
    m = re.search(r",?\s*(?:LT-?)?(\d{5})$", a)
    if m:
        return a[: m.start()].strip(", "), m.group(1)
    return a, ""


def parse_price(v) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        p = float(v)
    else:
        t = clean(v).replace(",", ".")
        if not re.fullmatch(r"\d+(\.\d+)?", t):
            return None  # pvz., „nepateikė“
        p = float(t)
    return round(p, 3) if 0.3 < p < 5 else None


def parse_date(v) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:  # Excel serijinis numeris
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", clean(v))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def haversine(a: tuple[float, float], b: tuple[float, float]) -> float:
    r = math.radians
    dl, dg = r(b[0] - a[0]), r(b[1] - a[1])
    h = math.sin(dl / 2) ** 2 + math.cos(r(a[0])) * math.cos(r(b[0])) * math.sin(dg / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


# ---------------------------------------------------------------- LEA failas
def find_lea_link(html: str) -> str:
    m = re.search(r'https://[a-z0-9.-]+\.sharepoint\.com/:x:/[^"\'<>\s]+', html)
    if not m:
        raise RuntimeError("LEA puslapyje nerasta Excel failo nuoroda")
    return m.group(0).replace("&amp;", "&")


def download_xlsx(session, share_url: str) -> bytes:
    url = share_url + ("&" if "?" in share_url else "?") + "download=1"
    r = session.get(url, timeout=120, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):
        raise RuntimeError(f"Gautas ne Excel failas ({r.headers.get('content-type')})")
    return r.content


HEADERS = {"imone": "company", "savivaldybe": "mun", "adresas": "addr",
           "degalu tipas": "fuel", "kaina": "price", "pateikimo data": "date"}


def parse_lea(data: bytes | str | Path) -> tuple[dict, dt.date]:
    """Grąžina ({raktas: degalinė}, naujausia data). Degalinė: company, mun, addr,
    prices={kuro_nr: (kaina, data)}, last=data."""
    import openpyxl

    src = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    col = None
    for _ in range(30):  # antraštės eilutė
        row = next(rows, None)
        if row is None:
            break
        names = [fold(c) for c in row]
        found = {}
        for i, n in enumerate(names):
            for key, field in HEADERS.items():
                if n.startswith(key):
                    found[field] = i
        if {"company", "addr", "fuel", "price", "date"} <= found.keys():
            col = found
            break
    if not col:
        raise RuntimeError("LEA faile nerasta antraštė (Įmonė, Adresas, Degalų tipas, Kaina, data)")

    stations: dict[str, dict] = {}
    max_date = None
    for row in rows:
        try:
            fuel = FUELS.get(fold(row[col["fuel"]]))
            d = parse_date(row[col["date"]])
            company, addr = clean(row[col["company"]]), clean(row[col["addr"]])
            mun = clean(row[col["mun"]]) if "mun" in col else ""
            raw_price = row[col["price"]]
        except IndexError:
            continue
        if not company or not addr or d is None:
            continue
        key = f"{company}|{mun}|{addr}"
        st = stations.setdefault(key, {"company": company, "mun": mun, "addr": addr, "prices": {}, "last": d})
        if d > st["last"]:
            st["last"] = d
        if fuel is None:
            continue
        p = parse_price(raw_price)
        if p is None:
            continue
        old = st["prices"].get(fuel)
        if old is None or d >= old[1]:
            st["prices"][fuel] = (p, d)
        if max_date is None or d > max_date:
            max_date = d
    wb.close()
    if max_date is None:
        raise RuntimeError("LEA faile nerasta nė vienos kainos")
    cutoff = max_date - dt.timedelta(days=KEEP_DAYS)
    stations = {k: s for k, s in stations.items() if s["prices"] and s["last"] >= cutoff}
    return stations, max_date


# ---------------------------------------------------------------- koordinatės
class Geocoder:
    """Randa degalinės koordinates: adresas -> Nominatim, tada „prisegama“ prie
    artimiausios OpenStreetMap degalinės (amenity=fuel). Viskas įsimenama cache."""

    def __init__(self, cache: dict, fetch_json: Callable, sleep: Callable = time.sleep,
                 today: dt.date | None = None, max_lookups: int = 900):
        self.cache, self.fetch_json, self.sleep = cache, fetch_json, sleep
        self.today = today or dt.date.today()
        self.max_lookups, self.lookups, self.errors = max_lookups, 0, 0
        self._fuel_points = None

    def fuel_points(self) -> list[dict]:
        if self._fuel_points is None:
            q = ('[out:json][timeout:180];area["ISO3166-1"="LT"][admin_level=2]->.a;'
                 '(node["amenity"="fuel"](area.a);way["amenity"="fuel"](area.a););out center tags;')
            try:
                data = self.fetch_json(OVERPASS, {"data": q}, post=True)
                pts = []
                for e in data.get("elements", []):
                    lat = e.get("lat", (e.get("center") or {}).get("lat"))
                    lng = e.get("lon", (e.get("center") or {}).get("lon"))
                    if lat is None or lng is None:
                        continue
                    t = e.get("tags", {})
                    pts.append({"lat": lat, "lng": lng,
                                "names": fold(" ".join(t.get(k, "") for k in ("brand", "name", "operator")))})
                self._fuel_points = pts
            except Exception as ex:  # be jų irgi veikiame
                print(f"Įspėjimas: nepavyko gauti OSM degalinių ({ex})", file=sys.stderr)
                self._fuel_points = []
        return self._fuel_points

    def snap(self, lat: float, lng: float, brand: str, radius_km: float) -> tuple[float, float] | None:
        words = [w for w in re.split(r"[^a-z0-9]+", fold(brand)) if len(w) >= 3]
        best, best_d, near, near_d = None, radius_km, None, 0.15
        for p in self.fuel_points():
            d = haversine((lat, lng), (p["lat"], p["lng"]))
            if d <= best_d and words and any(w in p["names"] for w in words):
                best, best_d = p, d
            if d <= near_d:
                near, near_d = p, d
        p = best or (near if radius_km <= 0.5 else None)
        return (round(p["lat"], 6), round(p["lng"], 6)) if p else None

    def nominatim(self, q: str):
        if self.lookups >= self.max_lookups:
            raise LookupError("viršytas paieškų skaičius šiam paleidimui")
        self.lookups += 1
        res = self.fetch_json(NOMINATIM, {"q": q, "format": "jsonv2", "countrycodes": "lt", "limit": 1})
        self.sleep(1.1)  # Nominatim taisyklė: ne daugiau kaip 1 užklausa per sekundę
        return (float(res[0]["lat"]), float(res[0]["lon"])) if res else None

    def locate(self, st: dict) -> dict:
        key = f"{st['company']}|{st['mun']}|{st['addr']}"
        hit = self.cache.get(key)
        if hit and (hit.get("q") != "none" or
                    (self.today - dt.date.fromisoformat(hit.get("t", "2000-01-01"))).days < RETRY_NONE_DAYS):
            return hit
        street, post = split_address(st["addr"])
        parts = [p.strip() for p in street.split(",") if p.strip()]
        settl = [p for p in parts if not re.search(r"\d", p) or re.search(r"\b(k|km|vs)\.?$", p)]
        tries = [
            (f"{street}, LT-{post}" if post else None, "adresas", 0.4),
            (street, "adresas", 0.4),
            (", ".join(settl + [st["mun"]]) if settl else None, "gyvenviete", 2.5),
            (st["mun"] or None, "savivaldybe", 0),
        ]
        out = {"q": "none", "t": self.today.isoformat()}
        try:
            for q, quality, radius in tries:
                if not q:
                    continue
                pt = self.nominatim(q)
                if not pt:
                    continue
                snapped = self.snap(pt[0], pt[1], brand_name(st["company"]), radius) if radius else None
                if snapped:
                    out = {"lat": snapped[0], "lng": snapped[1], "q": "osm-degaline", "t": out["t"]}
                else:
                    out = {"lat": round(pt[0], 6), "lng": round(pt[1], 6), "q": quality, "t": out["t"]}
                break
        except LookupError:
            return hit or {"q": "pending"}
        except Exception as ex:  # tinklo klaida – bandysime kitą kartą
            self.errors += 1
            print(f"Įspėjimas: {st['addr']}: {ex}", file=sys.stderr)
            if self.errors >= 5:
                self.max_lookups = self.lookups  # sustojame, kad neapkrautume serverio
            return hit or {"q": "pending"}
        self.cache[key] = out
        return out


def requests_fetch_json(session):
    def fetch(url, params, post=False):
        r = (session.post(url, data=params, timeout=200) if post
             else session.get(url, params=params, timeout=60))
        r.raise_for_status()
        return r.json()
    return fetch


# ---------------------------------------------------------------- išvestis
def build_output(stations: dict, coords: dict, max_date: dt.date,
                 generated: str | None = None) -> dict:
    items = sorted(stations.values(), key=lambda s: (brand_name(s["company"]).lower(), s["mun"], s["addr"]))
    N, L, C, D = [], [], [], []
    idx = lambda arr, v: arr.index(v) if v in arr else (arr.append(v) or len(arr) - 1)
    rows = []
    for s in items:
        key = f"{s['company']}|{s['mun']}|{s['addr']}"
        g = coords.get(key) or {}
        lat, lng = g.get("lat"), g.get("lng")
        approx = 1 if g.get("q") in ("gyvenviete", "savivaldybe") else 0
        if g.get("q") == "savivaldybe":
            lat = lng = None  # per netikslu atstumui skaičiuoti
        newest = max(d for _, d in s["prices"].values())
        price = lambda i: s["prices"][i][0] if i in s["prices"] else None
        rows.append([idx(N, brand_name(s["company"])), split_address(s["addr"])[0], idx(C, s["mun"]),
                     lat, lng, price(0), price(1), price(2), idx(D, newest.isoformat()),
                     idx(L, s["company"]), approx])
    return {
        "v": 1,
        "source": "Lietuvos energetikos agentūra (LEA) ir degalines valdančios įmonės",
        "geo": "© OpenStreetMap bendraautoriai (ODbL)",
        "prices_date": max_date.isoformat(),
        "generated": generated or dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "N": N, "L": L, "C": C, "D": D, "s": rows,
    }


def validate(out: dict, today: dt.date, prev: dict | None = None, max_age_days: int = 10) -> list[str]:
    """Saugikliai: tuščias ar pasenęs failas neturi pakeisti gero."""
    errs = []
    n = len(out["s"])
    if n < MIN_STATIONS:
        errs.append(f"per mažai degalinių: {n} (< {MIN_STATIONS})")
    age = (today - dt.date.fromisoformat(out["prices_date"])).days
    if age > max_age_days:
        errs.append(f"kainos per senos: {out['prices_date']} ({age} d.)")
    if prev and prev.get("s") and n < 0.6 * len(prev["s"]):
        errs.append(f"degalinių sumažėjo per daug: {len(prev['s'])} -> {n}")
    priced = sum(1 for r in out["s"] if any(r[i] is not None for i in (5, 6, 7)))
    if priced < 0.8 * max(n, 1):
        errs.append(f"per mažai degalinių su kainomis: {priced}/{n}")
    return errs


# ---------------------------------------------------------------- main
def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", default="out/fuel.json")
    ap.add_argument("--cache", default="out/geocache.json")
    ap.add_argument("--xlsx", help="vietinis LEA failas (be atsisiuntimo)")
    ap.add_argument("--no-geo", action="store_true", help="neieškoti naujų koordinačių")
    ap.add_argument("--max-lookups", type=int, default=900)
    ap.add_argument("--max-age", type=int, default=10, help="kiek dienų senumo kainos dar priimamos")
    a = ap.parse_args(list(argv) if argv is not None else None)

    import requests
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT
    today = dt.date.today()

    if a.xlsx:
        data = Path(a.xlsx).read_bytes()
    else:
        html = session.get(LEA_PAGE, timeout=60).text
        link = find_lea_link(html)
        print(f"LEA failas: {link}")
        data = download_xlsx(session, link)
    print(f"Atsisiųsta {len(data) / 1e6:.1f} MB")
    stations, max_date = parse_lea(data)
    print(f"Degalinių: {len(stations)}, naujausia data: {max_date}")

    cache_p, out_p = Path(a.cache), Path(a.out)
    cache = json.loads(cache_p.read_text("utf-8")) if cache_p.exists() else {}
    geo = Geocoder(cache, requests_fetch_json(session), today=today,
                   max_lookups=0 if a.no_geo else a.max_lookups)
    coords = {}
    for key, st in stations.items():
        coords[key] = geo.locate(st)
    found = sum(1 for g in coords.values() if g.get("lat") is not None)
    print(f"Koordinatės: {found}/{len(stations)} (naujų paieškų: {geo.lookups})")

    out = build_output(stations, coords, max_date)
    prev = json.loads(out_p.read_text("utf-8")) if out_p.exists() else None
    errs = validate(out, today, prev, a.max_age)
    cache_p.parent.mkdir(parents=True, exist_ok=True)
    cache_p.write_text(json.dumps(cache, ensure_ascii=False, indent=0, sort_keys=True), "utf-8")
    if errs:
        print("NEIŠSAUGOTA: " + "; ".join(errs), file=sys.stderr)
        return 2
    out_p.parent.mkdir(parents=True, exist_ok=True)
    out_p.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), "utf-8")
    print(f"Išsaugota: {out_p} ({out_p.stat().st_size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
