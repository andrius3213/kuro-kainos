import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import robotas as R  # noqa: E402

D12, D13 = dt.datetime(2026, 11, 12), dt.datetime(2026, 11, 13)


def make_xlsx(path, extra_rows=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DK-2026"
    ws["A2"] = "Degalų kainos Lietuvos degalinėse (nuo 2026-04-08)"
    ws["A5"] = "Duomenys: LEA. Pastaba: naudojant šią informaciją, būtina nurodyti šaltinį"
    ws.append([])
    ws.cell(row=7, column=1, value="Įmonė")
    for i, h in enumerate(["Įmonė", "Savivaldybė", "Adresas", "Degalų tipas", "Kaina (EUR/l)", "Pateikimo data"], 1):
        ws.cell(row=7, column=i, value=h)
    ck = ("UAB Circle K Lietuva", "Vilniaus m. sav.", "Vilnius, Savanorių pr. 118, 03153")
    orl = ("AB Orlen Baltics Retail", "Kauno m. sav.", "Kaunas, Pramonės pr. 37, 51270")
    rows = [
        (*ck, "95 benzinas", 1.829, D12), (*ck, "Dyzelinas", 2.229, D12), (*ck, "SND", 1.099, D12),
        (*ck, "95 benzinas", 1.809, D13), (*ck, "Dyzelinas", 2.209, D13), (*ck, "SND", "nepateikė", D13),
        (*ck, "98 benzinas", 1.999, D13),
        (*orl, "95 benzinas", "1,729", "2026-11-13"), (*orl, "Dyzelinas", "2,149", "2026-11-13"),
        ("UAB Senoji", "Alytaus m. sav.", "Alytus, Senoji g. 1, 62100", "95 benzinas", 1.5, dt.datetime(2026, 8, 1)),
        ("A. Praškevičiaus IĮ", "Kaišiadorių r. sav.", "Migonių k., Kalvių g. 4, 56123", "SND", 0.98, D13),
        ("UAB Vimijula", "Jonavos r. sav.", "Ručiūnų k., Kauno g. 2A, 55355", "95 benzinas", 1.96, D13),
        *extra_rows,
    ]
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return path


@pytest.fixture
def xlsx(tmp_path):
    return make_xlsx(tmp_path / "DK-2026.xlsx")


def test_small_helpers():
    assert R.brand_name("AB Orlen Baltics Retail") == "Orlen"
    assert R.brand_name("UAB Circle K Lietuva") == "Circle K"
    assert R.brand_name("UAB Vimijula") == "Vimijula"
    assert R.brand_name("A. Praškevičiaus IĮ") == "A. Praškevičiaus IĮ"
    assert R.split_address("Ručiūnų k., Kauno g. 2A, 55355") == ("Ručiūnų k., Kauno g. 2A", "55355")
    assert R.split_address("Bučių k.") == ("Bučių k.", "")
    assert R.parse_price("nepateikė") is None
    assert R.parse_price("1,030") == 1.03
    assert R.parse_price(2.2799999) == 2.28
    assert R.parse_date("2026-11-13") == dt.date(2026, 11, 13)
    assert R.parse_date(46339) == dt.date(2026, 11, 13)


def test_find_lea_link():
    html = '<a href="https://ltenergagen.sharepoint.com/:x:/s/intra/doc/IQBf?e=7t&amp;x=1"><img></a>'
    assert R.find_lea_link(html) == "https://ltenergagen.sharepoint.com/:x:/s/intra/doc/IQBf?e=7t&x=1"
    with pytest.raises(RuntimeError):
        R.find_lea_link("<html></html>")


def test_parse_lea_takes_latest_prices(xlsx):
    st, max_date = R.parse_lea(Path(xlsx).read_bytes())
    assert max_date == dt.date(2026, 11, 13)
    assert len(st) == 4  # senoji (2026-08-01) atmesta
    ck = st["UAB Circle K Lietuva|Vilniaus m. sav.|Vilnius, Savanorių pr. 118, 03153"]
    assert ck["prices"][0] == (1.809, dt.date(2026, 11, 13))
    assert ck["prices"][1] == (2.209, dt.date(2026, 11, 13))
    assert ck["prices"][2] == (1.099, dt.date(2026, 11, 12))  # 13 d. „nepateikė“ – lieka ankstesnė
    orl = st["AB Orlen Baltics Retail|Kauno m. sav.|Kaunas, Pramonės pr. 37, 51270"]
    assert orl["prices"] == {0: (1.729, dt.date(2026, 11, 13)), 1: (2.149, dt.date(2026, 11, 13))}


def test_parse_lea_rejects_wrong_file(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.append(["kažkas", "ne", "tas"])
    p = tmp_path / "x.xlsx"
    wb.save(p)
    with pytest.raises(RuntimeError):
        R.parse_lea(p.read_bytes())


class FakeNet:
    def __init__(self, nominatim=None, fuel=None, fail=False):
        self.nominatim, self.fuel, self.fail, self.calls = nominatim or {}, fuel or [], fail, []

    def __call__(self, url, params, post=False):
        self.calls.append(url)
        if self.fail:
            raise ConnectionError("tinklas neveikia")
        if url == R.OVERPASS:
            return {"elements": self.fuel}
        hit = self.nominatim.get(params["q"])
        return [{"lat": str(hit[0]), "lon": str(hit[1])}] if hit else []


ST = {"company": "UAB Circle K Lietuva", "mun": "Vilniaus m. sav.", "addr": "Vilnius, Savanorių pr. 118, 03153"}
KEY = "UAB Circle K Lietuva|Vilniaus m. sav.|Vilnius, Savanorių pr. 118, 03153"


def test_geocoder_snaps_to_matching_osm_station_and_caches():
    net = FakeNet(nominatim={"Vilnius, Savanorių pr. 118, LT-03153": (54.6615, 25.2350)},
                  fuel=[{"type": "node", "lat": 54.6617, "lon": 25.2353, "tags": {"brand": "Circle K"}},
                        {"type": "node", "lat": 54.6616, "lon": 25.2351, "tags": {"brand": "Neste"}}])
    cache = {}
    g = R.Geocoder(cache, net, sleep=lambda s: None, today=dt.date(2026, 11, 15))
    res = g.locate(ST)
    assert res["q"] == "osm-degaline" and (res["lat"], res["lng"]) == (54.6617, 25.2353)
    n = len(net.calls)
    assert g.locate(ST) == res and len(net.calls) == n  # antrą kartą – iš atminties
    assert cache[KEY]["q"] == "osm-degaline"


def test_geocoder_falls_back_and_retries_unknown_later():
    net = FakeNet(nominatim={"Vilniaus m. sav.": (54.68, 25.28)})
    cache = {}
    g = R.Geocoder(cache, net, sleep=lambda s: None, today=dt.date(2026, 11, 15))
    assert g.locate(ST)["q"] == "savivaldybe"
    st2 = dict(ST, addr="Niekur, Nėra g. 1, 00000", mun="")
    assert g.locate(st2)["q"] == "none"
    calls = len(net.calls)
    g.locate(st2)
    assert len(net.calls) == calls  # nerastas – nebandoma iš naujo tą pačią dieną
    g2 = R.Geocoder(cache, net, sleep=lambda s: None, today=dt.date(2026, 12, 31))
    g2.locate(st2)
    assert len(net.calls) > calls  # po 30 d. bandoma vėl


def test_geocoder_limits_and_network_errors():
    g = R.Geocoder({}, FakeNet(), sleep=lambda s: None, max_lookups=0)
    assert g.locate(ST) == {"q": "pending"}
    cache = {}
    g = R.Geocoder(cache, FakeNet(fail=True), sleep=lambda s: None)
    for _ in range(6):
        assert g.locate(ST) == {"q": "pending"}
    assert cache == {} and g.max_lookups == g.lookups  # po 5 klaidų sustojo


def test_build_output_matches_app_format(xlsx):
    st, max_date = R.parse_lea(Path(xlsx).read_bytes())
    coords = {k: {"lat": 54.0, "lng": 25.0, "q": "adresas"} for k in st}
    approx_key = next(k for k in st if "Vimijula" in k)
    coords[approx_key] = {"lat": 55.1, "lng": 24.2, "q": "gyvenviete"}
    mun_key = next(k for k in st if "Praškevičiaus" in k)
    coords[mun_key] = {"lat": 54.8, "lng": 24.4, "q": "savivaldybe"}
    out = R.build_output(st, coords, max_date, generated="2026-11-15T10:00:00+00:00")
    assert out["prices_date"] == "2026-11-13" and len(out["s"]) == 4
    json.dumps(out)
    for r in out["s"]:
        assert len(r) == 11
        assert out["N"][r[0]] and out["C"][r[2]] and out["D"][r[8]] and out["L"][r[9]]
    by = {out["N"][r[0]]: r for r in out["s"]}
    assert by["Circle K"][1] == "Vilnius, Savanorių pr. 118" and by["Circle K"][5:8] == [1.809, 2.209, 1.099]
    assert out["L"][by["Orlen"][9]] == "AB Orlen Baltics Retail"
    assert by["Vimijula"][10] == 1 and by["Vimijula"][3] == 55.1
    assert by["A. Praškevičiaus IĮ"][3] is None  # tik savivaldybės tikslumas – be koordinačių


def test_validate_guards():
    row = [0, "a", 0, 1, 1, 1.7, 2.1, None, 0, 0, 0]
    good = {"prices_date": "2026-11-13", "s": [row] * 400}
    assert R.validate(good, dt.date(2026, 11, 15)) == []
    assert any("per mažai" in e for e in R.validate({"prices_date": "2026-11-13", "s": [row] * 10}, dt.date(2026, 11, 15)))
    assert any("per senos" in e for e in R.validate(good, dt.date(2026, 12, 15)))
    assert any("sumažėjo" in e for e in R.validate({"prices_date": "2026-11-13", "s": [row] * 300},
                                                    dt.date(2026, 11, 15), prev={"s": [row] * 700}))


def test_main_offline_end_to_end(xlsx, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(R, "MIN_STATIONS", 1)
    out, cache = tmp_path / "out" / "fuel.json", tmp_path / "out" / "geocache.json"
    rc = R.main(["--xlsx", str(xlsx), "--no-geo", "--out", str(out), "--cache", str(cache), "--max-age", "100000"])
    assert rc == 0, capsys.readouterr().err
    data = json.loads(out.read_text("utf-8"))
    assert len(data["s"]) == 4 and data["geo"].startswith("©")
    monkeypatch.setattr(R, "MIN_STATIONS", 300)
    rc = R.main(["--xlsx", str(xlsx), "--no-geo", "--out", str(out), "--cache", str(cache), "--max-age", "100000"])
    assert rc == 2 and json.loads(out.read_text("utf-8")) == data  # blogas rezultatas neperrašo gero
