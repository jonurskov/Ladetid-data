"""
Ladetid data, Jon Pedersen

Kører to gange om dagen på GitHub og gør seks ting:
1. Gemmer de prognoser, appen bruger (elpriser.org og Strømligning), som et "øjebliksbillede".
   Morgenkørslen fanger prognosen for i morgen, eftermiddagskørslen den nye prognose, der kommer ved 13 tiden.
2. Gemmer de faktorer, der påvirker prisen, som de så ud på det tidspunkt: vind, sol, temperatur og helligdage.
3. Henter de faktiske spotpriser fra Energi Data Service (Energinet).
4. Sammenligner gamle prognoser med de faktiske priser og måler fejlen.
5. Lærer af fejlene og beregner en rettelse. Rettelsen må aldrig flytte prognosen mere end 10 % og 15 øre.
6. Laver en Excel fil for hver uge med alle fejlene.

Alle priser er spotpris i kr pr. kWh uden moms, tariffer og afgifter.
"""

import json
import os
import shutil
import statistics
import sys
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

CPH = ZoneInfo("Europe/Copenhagen")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
EXCEL = os.path.join(ROOT, "excel")
AREAS = ["DK1", "DK2"]
HORIZONS = range(1, 10)          # dage fra prognosen blev lavet til den gælder
SHORT = range(1, 5)              # kortsigtet: 1 til 4 dage frem
LONG = range(5, 10)              # langsigtet: 5 til 9 dage frem
WINDOW_DAYS = 28                 # rapporten bygger på de sidste 4 uger
TRAIN_DAYS = 90                  # rettelsen lærer af op til 90 dage
CHARGE_HOURS = 4                 # ladetid brugt til "ramte ladetiden"
MAX_FIX_PCT = 0.10               # rettelsen må højst flytte prognosen 10 %
MAX_FIX_KR = 0.15                # og højst 15 øre
MIN_DAYS_FACTORS = 21            # faktorerne bruges først, når der er 21 dage at lære af
FACTORS = ["temp", "vind", "sol", "helligdag"]
UA = "Ladetid-data/1.1 (github.com/jonurskov/Ladetid-data)"


# ---------------------------------------------------------------- hjælpere

def http_json(url, tries=4):
    """Henter JSON. Prøver igen, hvis siden siger, at den har for travlt."""
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=60) as r:
                j = json.loads(r.read().decode("utf-8"))
            if isinstance(j, dict) and j.get("statusCode") == 429:
                raise IOError("for mange forespørgsler")
            return j
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(30)


def read_json(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def hourly(values_by_key):
    """{(dato, time): [værdier]} til {dato: [24 værdier eller None]}"""
    out = {}
    for (d, h), vals in values_by_key.items():
        out.setdefault(d, [None] * 24)[h] = round(sum(vals) / len(vals), 5)
    return out


def best_window_start(prices, n=CHARGE_HOURS):
    best, start = None, None
    for s in range(0, 24 - n + 1):
        part = prices[s:s + n]
        if any(p is None for p in part):
            continue
        avg = sum(part) / n
        if best is None or avg < best:
            best, start = avg, s
    return start


def tertiles(prices):
    """Deler døgnets timer i billig (0), middel (1) og dyr (2)."""
    order = sorted(range(24), key=lambda h: prices[h])
    cls = [0] * 24
    for rank, h in enumerate(order):
        cls[h] = 0 if rank < 8 else (1 if rank < 16 else 2)
    return cls


def clamp_fix(fix, forecast):
    """Rettelsen må ikke flytte prognosen langt væk fra kilderne."""
    limit = min(MAX_FIX_KR, MAX_FIX_PCT * abs(forecast))
    return max(-limit, min(limit, fix))


# ---------------------------------------------------------------- helligdage

def easter(y):
    a, b, c = y % 19, y // 100, y % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(y, month, day)


def holidays_dk(y):
    e = easter(y)
    days = [date(y, 1, 1), e - timedelta(3), e - timedelta(2), e, e + timedelta(1),
            e + timedelta(39), e + timedelta(49), e + timedelta(50), date(y, 6, 5),
            date(y, 12, 24), date(y, 12, 25), date(y, 12, 26), date(y, 12, 31)]
    return set(days)


def holidays_de(y):
    e = easter(y)
    return {date(y, 1, 1), e - timedelta(2), e + timedelta(1), date(y, 5, 1), e + timedelta(39),
            e + timedelta(50), date(y, 10, 3), date(y, 12, 25), date(y, 12, 26)}


def holiday_score(d):
    """1 for dansk helligdag, 0,5 for tysk helligdag eller dagene mellem jul og nytår."""
    s = 0.0
    if d in holidays_dk(d.year):
        s = 1.0
    elif d in holidays_de(d.year):
        s = 0.5
    if d.month == 12 and 27 <= d.day <= 30:
        s = max(s, 0.5)
    return s


# ---------------------------------------------------------------- 1. prognoser

def fetch_elpriser(area):
    j = http_json(f"https://elpriser.org/api/forecast?area={area}&mode=spot_ex")
    out = {}
    for day in j.get("days", []):
        if day.get("type") != "forecast":
            continue
        arr = [None] * 24
        for p in day.get("prices", []):
            h = int(p["hour"])
            if 0 <= h < 24:
                arr[h] = round(float(p["price"]), 5)
        out[day["date"]] = arr
    return out


def fetch_stromligning(area):
    j = http_json(f"https://stromligning.dk/api/prices?priceArea={area}&forecast=true")
    vals = {}
    for p in j.get("prices", []):
        if not p.get("forecast"):
            continue
        ld = p["localDate"]
        d, h = ld[:10], int(ld[11:13])
        el = (p.get("details") or {}).get("electricity") or {}
        v = el.get("value", (p.get("price") or {}).get("value"))
        if v is None:
            continue
        vals.setdefault((d, h), []).append(float(v))
    return hourly(vals)


# ---------------------------------------------------------------- 2. faktorer

# Steder hvor meget af vind og solstrømmen produceres, og hvor mange bor (temperatur)
POINTS = [
    ("vind", 56.0, 8.1),    # Vestjylland
    ("vind", 54.2, 8.8),    # Nordtyskland ved Nordsøen
    ("vind", 57.0, 13.5),   # Sydsverige
    ("sol", 50.0, 10.0),    # Midt i Tyskland
    ("sol", 55.8, 10.5),    # Midt i Danmark
    ("temp", 55.7, 12.4),   # København
    ("temp", 56.2, 10.1),   # Aarhus
    ("temp", 59.3, 18.0),   # Stockholm
    ("temp", 59.9, 10.7),   # Oslo
]


def fetch_factors():
    """Vejrudsigt fra Open-Meteo (bygger bl.a. på DMI og ECMWF): {dato: {temp, vind, sol, helligdag}}"""
    lat = ",".join(str(p[1]) for p in POINTS)
    lon = ",".join(str(p[2]) for p in POINTS)
    url = (f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}"
           "&hourly=temperature_2m,wind_speed_100m,shortwave_radiation&wind_speed_unit=ms"
           "&timezone=Europe%2FCopenhagen&forecast_days=10")
    j = http_json(url)
    if isinstance(j, dict):
        j = [j]
    per = {}  # (dato, slags) -> liste af dagsværdier pr. sted
    for (kind, _, _), loc in zip(POINTS, j):
        hr = loc["hourly"]
        by_day = {}
        for t, temp, wind, rad in zip(hr["time"], hr["temperature_2m"], hr["wind_speed_100m"],
                                      hr["shortwave_radiation"]):
            d = t[:10]
            v = {"temp": temp, "vind": wind, "sol": rad}[kind]
            if v is not None:
                by_day.setdefault(d, []).append(v)
        for d, vals in by_day.items():
            if len(vals) < 20:
                continue
            # sol: samlet indstråling i MJ pr. m², ellers døgnets gennemsnit
            val = sum(vals) * 0.0036 if kind == "sol" else sum(vals) / len(vals)
            per.setdefault((d, kind), []).append(val)
    out = {}
    for (d, kind), vals in per.items():
        out.setdefault(d, {})[kind] = round(sum(vals) / len(vals), 2)
    for d in out:
        out[d]["helligdag"] = holiday_score(date.fromisoformat(d))
    return out


def save_snapshot(today, now_utc):
    """Gemmer prognoser og faktorer. Én fil om morgenen og én om eftermiddagen."""
    part = "am" if now_utc.hour < 12 else "pm"
    try:
        factors = fetch_factors()
    except Exception as e:
        print(f"Advarsel: vejr: {e}", file=sys.stderr)
        factors = {}
    for area in AREAS:
        path = os.path.join(DATA, "prognoser", area, f"{today.isoformat()}-{part}.json")
        if os.path.exists(path):
            continue
        snap = {"lavet": now_utc.isoformat(), "dato": today.isoformat(), "kilder": {},
                "faktorer": {d: f for d, f in factors.items() if d > today.isoformat()}}
        for name, fn in (("elpriser", fetch_elpriser), ("stromligning", fetch_stromligning)):
            try:
                snap["kilder"][name] = {d: v for d, v in fn(area).items() if d > today.isoformat()}
            except Exception as e:
                print(f"Advarsel: {name} {area}: {e}", file=sys.stderr)
        if snap["kilder"]:
            write_json(path, snap)
            print(f"Gemte prognose {area} {today} {part}")


# ---------------------------------------------------------------- 3. faktiske priser

def fetch_actual(areas, start, end):
    """Faktiske spotpriser for alle områder i én forespørgsel: {område: {dato: [24]}}"""
    flt = urllib.parse.quote(json.dumps({"PriceArea": list(areas)}))
    url = (f"https://api.energidataservice.dk/dataset/DayAheadPrices?start={start.isoformat()}"
           f"&end={end.isoformat()}&filter={flt}&sort=TimeDK&limit=0")
    j = http_json(url)
    vals = {}
    for r in j.get("records", []):
        t, v, a = r.get("TimeDK"), r.get("DayAheadPriceDKK"), r.get("PriceArea")
        if t is None or v is None or a not in areas:
            continue
        vals.setdefault(a, {}).setdefault((t[:10], int(t[11:13])), []).append(float(v) / 1000)
    return {a: hourly(v) for a, v in vals.items()}


def update_actuals(today):
    path = os.path.join(DATA, "faktiske_priser.json")
    actual = read_json(path, {})
    try:
        all_got = fetch_actual(AREAS, today - timedelta(days=45), today + timedelta(days=2))
    except Exception as e:
        print(f"Advarsel: faktiske priser: {e}", file=sys.stderr)
        all_got = {}
    for area in AREAS:
        got = all_got.get(area, {})
        a = actual.setdefault(area, {})
        for d, arr in got.items():
            if sum(v is not None for v in arr) >= 23:  # kun hele døgn
                a[d] = arr
    write_json(path, actual)
    return actual


# ---------------------------------------------------------------- 4. fejl

def load_snapshots(area):
    """{dato: snapshot}. Eftermiddagens prognose bruges frem for morgenens, hvor den findes,
    fordi det er den, appen viser det meste af dagen."""
    folder = os.path.join(DATA, "prognoser", area)
    if not os.path.isdir(folder):
        return {}
    parts = {}
    for fn in sorted(os.listdir(folder)):
        if not fn.endswith(".json"):
            continue
        d = fn[:10]
        part = "am" if "-am" in fn else "pm"   # gamle filer uden endelse blev lavet om eftermiddagen
        s = read_json(os.path.join(folder, fn))
        if s:
            parts.setdefault(d, {})[part] = s
    out = {}
    for d, p in parts.items():
        am, pm = p.get("am"), p.get("pm")
        base = {"dato": d, "kilder": {}, "faktorer": {}}
        for snap in (am, pm):  # pm overskriver am
            if not snap:
                continue
            for src, days in snap.get("kilder", {}).items():
                base["kilder"].setdefault(src, {}).update(days)
            base["faktorer"].update(snap.get("faktorer", {}))
        out[d] = base
    return out


def combine(snap, target):
    """Appens prognose er gennemsnittet af de kilder, der har et tal for timen."""
    srcs = snap["kilder"]
    e = (srcs.get("elpriser") or {}).get(target)
    s = (srcs.get("stromligning") or {}).get(target)
    if not e and not s:
        return None, None, None
    app = []
    for h in range(24):
        vals = [x[h] for x in (e, s) if x and x[h] is not None]
        app.append(sum(vals) / len(vals) if vals else None)
    return e, s, app


def build_rows(actual):
    """En række pr. område, dato, time og horisont, hvor både prognose og facit findes."""
    rows = []
    for area in AREAS:
        snaps = load_snapshots(area)
        act = actual.get(area, {})
        for made, snap in snaps.items():
            made_d = date.fromisoformat(made)
            for h in HORIZONS:
                target = (made_d + timedelta(days=h)).isoformat()
                if target not in act:
                    continue
                e, s, app = combine(snap, target)
                if app is None:
                    continue
                fac = snap.get("faktorer", {}).get(target, {})
                for hour in range(24):
                    a = act[target][hour]
                    if app[hour] is None or a is None:
                        continue
                    rows.append({
                        "område": area, "lavet": made, "dato": target, "time": hour, "horisont": h,
                        "elpriser": e[hour] if e else None, "stromligning": s[hour] if s else None,
                        "prognose": round(app[hour], 5), "faktisk": a,
                        **{k: fac.get(k) for k in FACTORS},
                    })
    return rows


# ---------------------------------------------------------------- 5. læring

def bias_table(rows):
    """Typisk fejl (prognose minus facit) pr. horisont og time på døgnet.
    Timerne udglattes over naboer, og tallet trækkes mod 0, når der er få data."""
    buckets = {}
    for r in rows:
        buckets.setdefault((r["horisont"], r["time"]), []).append(r["prognose"] - r["faktisk"])
    table = {}
    for h in HORIZONS:
        arr = []
        for hour in range(24):
            vals = []
            for dh in (-1, 0, 1):
                vals += buckets.get((h, (hour + dh) % 24), [])
            if not vals:
                arr.append(0.0)
                continue
            n_days = len(vals) / 3
            arr.append(round(statistics.median(vals) * n_days / (n_days + 7), 5))
        table[str(h)] = arr
    return table


def solve(a, b):
    """Løser a x = b (lille ligningssystem) med Gauss elimination."""
    n = len(b)
    m = [row[:] + [b[i]] for i, row in enumerate(a)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(m[r][c]))
        m[c], m[p] = m[p], m[c]
        if abs(m[c][c]) < 1e-12:
            return [0.0] * n
        for r in range(n):
            if r != c:
                f = m[r][c] / m[c][c]
                m[r] = [x - f * y for x, y in zip(m[r], m[c])]
    return [m[i][n] / m[i][i] for i in range(n)]


def group_of(h):
    return "kort" if h in SHORT else "lang"


def peak_of(hour):
    return "spids" if 7 <= hour <= 20 else "ro"


def factor_model(rows, bias):
    """Lærer hvordan fejlen hænger sammen med faktorerne, når den typiske fejl er trukket fra.
    Bruger dagsgennemsnit, så hver dag tæller én gang, og en forsigtig ridge regression,
    som trækker vægtene mod 0. Returnerer {gruppe: {spids/ro: {mean, sd, w}}} eller {}."""
    days = {}
    for r in rows:
        if any(r.get(k) is None for k in FACTORS):
            continue
        res = r["prognose"] - r["faktisk"] - bias[str(r["horisont"])][r["time"]]
        k = (group_of(r["horisont"]), peak_of(r["time"]), r["lavet"], r["dato"])
        d = days.setdefault(k, {"res": [], "x": [r[f] for f in FACTORS]})
        d["res"].append(res)
    model = {}
    for g in ("kort", "lang"):
        for p in ("spids", "ro"):
            pts = [(v["x"], sum(v["res"]) / len(v["res"])) for (gg, pp, _, _), v in days.items()
                   if gg == g and pp == p]
            if len({k[3] for k in days if k[0] == g and k[1] == p}) < MIN_DAYS_FACTORS:
                continue
            n = len(FACTORS)
            mean = [sum(x[i] for x, _ in pts) / len(pts) for i in range(n)]
            sd = [max(1e-6, (sum((x[i] - mean[i]) ** 2 for x, _ in pts) / len(pts)) ** 0.5) for i in range(n)]
            z = [([(x[i] - mean[i]) / sd[i] for i in range(n)], y) for x, y in pts]
            lam = 0.5 * len(z)  # kraftig regularisering, så vægtene holdes små
            a = [[sum(zz[i] * zz[j] for zz, _ in z) + (lam if i == j else 0) for j in range(n)] for i in range(n)]
            b = [sum(zz[i] * y for zz, y in z) for i in range(n)]
            w = solve(a, b)
            model.setdefault(g, {})[p] = {"mean": mean, "sd": sd, "w": [round(x, 5) for x in w]}
    return model


def predict_fix(bias, model, horizon, hour, fac, forecast):
    fix = bias[str(horizon)][hour]
    m = model.get(group_of(horizon), {}).get(peak_of(hour))
    if m and fac and all(fac.get(k) is not None for k in FACTORS):
        fix += sum(w * (fac[k] - mu) / sd for w, k, mu, sd in zip(m["w"], FACTORS, m["mean"], m["sd"]))
    return clamp_fix(fix, forecast)


def train(rows, before=None):
    """Træner på data, der var kendt før datoen 'before' (de sidste TRAIN_DAYS dage)."""
    if before:
        start = (date.fromisoformat(before) - timedelta(days=TRAIN_DAYS)).isoformat()
        rows = [r for r in rows if start <= r["dato"] < before]
    bias = bias_table(rows)
    return bias, factor_model(rows, bias)


def corrected_rows(rows, history):
    """Rettede prognoser for 'rows', hvor rettelsen for hver prognose kun bygger på det,
    der var kendt, da prognosen blev lavet. Så er sammenligningen ærlig."""
    cache = {}
    by_area = {}
    for x in history:
        by_area.setdefault(x["område"], []).append(x)
    out = []
    for r in rows:
        key = (r["område"], r["lavet"])
        if key not in cache:
            cache[key] = train(by_area.get(r["område"], []), before=r["lavet"])
        bias, model = cache[key]
        fac = {k: r.get(k) for k in FACTORS}
        out.append(r["prognose"] - predict_fix(bias, model, r["horisont"], r["time"], fac, r["prognose"]))
    return out


# ---------------------------------------------------------------- rapport

def metrics(rows, values=None):
    """values: alternative prognosetal (fx de rettede), samme rækkefølge som rows."""
    if not rows:
        return None
    vals = values if values is not None else [r["prognose"] for r in rows]
    abs_err = [abs(v - r["faktisk"]) for v, r in zip(vals, rows)]
    sum_act = sum(abs(r["faktisk"]) for r in rows)
    bias = sum(v - r["faktisk"] for v, r in zip(vals, rows)) / len(rows)

    days = {}
    for v, r in zip(vals, rows):
        k = (r["område"], r["dato"], r["horisont"])
        d = days.setdefault(k, ([None] * 24, [None] * 24))
        d[0][r["time"]] = v
        d[1][r["time"]] = r["faktisk"]
    hit_win, n_win, hit_cls, n_cls = 0, 0, 0, 0
    for fc, ac in days.values():
        if any(x is None for x in fc + ac):
            continue
        sf, sa = best_window_start(fc), best_window_start(ac)
        if sf is not None and sa is not None:
            n_win += 1
            hit_win += abs(sf - sa) <= 1
        cf, ca = tertiles(fc), tertiles(ac)
        n_cls += 24
        hit_cls += sum(1 for h in range(24) if cf[h] == ca[h])

    per_hour = []
    for hour in range(24):
        e = [abs(v - r["faktisk"]) for v, r in zip(vals, rows) if r["time"] == hour]
        per_hour.append(round(100 * sum(e) / len(e), 1) if e else None)

    return {
        "dage": len({r["dato"] for r in rows}),
        "fejl_ore": round(100 * sum(abs_err) / len(abs_err), 1),
        "fejl_pct": round(100 * sum(abs_err) / sum_act, 1) if sum_act > 0 else None,
        "skaevhed_ore": round(100 * bias, 1),
        "ramte_ladetid_pct": round(100 * hit_win / n_win) if n_win else None,
        "ramte_doegn_pct": round(100 * hit_cls / n_cls) if n_cls else None,
        "fejl_pr_time_ore": per_hour,
    }


def describe_model(model):
    """Vægtene oversat til øre pr. enhed, så de kan læses af et menneske."""
    unit = {"temp": "pr. grad varmere", "vind": "pr. m/s mere vind", "sol": "pr. MJ/m² mere sol",
            "helligdag": "på en helligdag"}
    out = {}
    for g, parts in model.items():
        for p, m in parts.items():
            out.setdefault(g, {})[p] = {
                k: f"{round(-100 * w / sd, 2)} øre {unit[k]}" for k, w, sd in zip(FACTORS, m["w"], m["sd"])
            }
    return out


def make_report(rows, today, snaps_today):
    start = (today - timedelta(days=WINDOW_DAYS)).isoformat()
    recent = [r for r in rows if r["dato"] >= start]
    corr = corrected_rows(recent, rows) if recent else []
    now = datetime.now(timezone.utc).isoformat()
    report = {"lavet": now, "periode_fra": start, "periode_til": today.isoformat(),
              "graense": f"Rettelsen flytter højst prognosen {int(MAX_FIX_PCT * 100)} % og {int(MAX_FIX_KR * 100)} øre",
              "omraader": {}}
    rettelse = {"lavet": now, "max_pct": MAX_FIX_PCT, "max_kr": MAX_FIX_KR, "omraader": {}}
    for area in AREAS:
        idx = [i for i, r in enumerate(recent) if r["område"] == area]
        a_rows = [recent[i] for i in idx]
        a_corr = [corr[i] for i in idx]
        res = {"horisonter": []}
        use = {}
        for name, rng in (("kort", SHORT), ("lang", LONG)):
            sel = [i for i, r in enumerate(a_rows) if r["horisont"] in rng]
            raw = metrics([a_rows[i] for i in sel])
            fixed = metrics([a_rows[i] for i in sel], [a_corr[i] for i in sel])
            res[name] = raw
            res[name + "_rettet"] = fixed
            # rettelsen bruges kun, hvis den har ramt mindst 3 % bedre over mindst 14 dage
            use[name] = bool(raw and fixed and raw["dage"] >= 14
                             and fixed["fejl_ore"] <= raw["fejl_ore"] * 0.97)
        for h in HORIZONS:
            m = metrics([r for r in a_rows if r["horisont"] == h])
            if m:
                m.pop("fejl_pr_time_ore")
                res["horisonter"].append({"horisont": h, **m})
        res["bruger_rettelse"] = use

        # Rettelse til de kommende dage ud fra dagens nyeste prognose og faktorer
        all_area = [r for r in rows if r["område"] == area]
        bias, model = train(all_area, before=(today + timedelta(days=1)).isoformat())
        res["faktorer"] = describe_model(model)
        report["omraader"][area] = res
        dage = {}
        snap = snaps_today.get(area)
        if snap:
            for h in HORIZONS:
                target = (today + timedelta(days=h)).isoformat()
                _, _, app = combine(snap, target)
                if not app:
                    continue
                fac = snap.get("faktorer", {}).get(target, {})
                dage[target] = [round(predict_fix(bias, model, h, hour, fac, app[hour]), 5)
                                if app[hour] is not None else 0.0 for hour in range(24)]
        skaev = {k: [round(clamp_fix(v, 1.5), 5) for v in arr] for k, arr in bias.items()}
        rettelse["omraader"][area] = {"brug": use, "skaevhed": skaev, "dage": dage}
    write_json(os.path.join(DATA, "rapport.json"), report)
    write_json(os.path.join(DATA, "rettelse.json"), rettelse)
    return report


# ---------------------------------------------------------------- 6. Excel

def write_excel(rows, week_start, path, title):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    week_end = week_start + timedelta(days=6)
    sel = [r for r in rows if week_start.isoformat() <= r["dato"] <= week_end.isoformat()]
    corr_map = {id(r): c for r, c in zip(sel, corrected_rows(sel, rows))} if sel else {}

    wb = Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F7A4D")

    def header(ws, cols, row=1):
        for i, c in enumerate(cols, 1):
            cell = ws.cell(row=row, column=i, value=c)
            cell.font, cell.fill = head, fill
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    ws = wb.active
    ws.title = "Oversigt"
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Spotpris i kr pr. kWh uden moms. Fejl i øre er den gennemsnitlige afvigelse pr. time. "
                "Fejl i procent er den samlede afvigelse delt med den samlede faktiske pris. "
                "Ramte ladetiden: den billigste 4 timers periode ramt inden for 1 time. "
                "Ramte døgnets forløb: andel af timer, hvor prognosen havde ret i om timen var billig, middel eller dyr. "
                f"Rettelsen flytter højst prognosen {int(MAX_FIX_PCT * 100)} % og {int(MAX_FIX_KR * 100)} øre.")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 75
    cols = ["Område", "Prognose", "Dage", "Fejl (øre)", "Fejl (%)", "Skævhed (øre)",
            "Ramte ladetiden (%)", "Ramte døgnets forløb (%)", "Fejl med rettelse (øre)"]
    header(ws, cols, 4)
    rr = 5
    for area in AREAS:
        for name, rng in (("Kortsigtet (1 til 4 dage)", SHORT), ("Langsigtet (5 til 9 dage)", LONG)):
            part = [r for r in sel if r["område"] == area and r["horisont"] in rng]
            m = metrics(part)
            mc = metrics(part, [corr_map[id(r)] for r in part]) if part else None
            ws.append([area, name] + ([m["dage"], m["fejl_ore"], m["fejl_pct"], m["skaevhed_ore"],
                                       m["ramte_ladetid_pct"], m["ramte_doegn_pct"], mc["fejl_ore"]]
                                      if m else ["Ingen data endnu"]))
            rr += 1
    rr += 1
    ws.cell(row=rr, column=1, value="Pr. horisont").font = Font(bold=True)
    header(ws, ["Område", "Dage frem", "Dage", "Fejl (øre)", "Fejl (%)", "Skævhed (øre)",
                "Ramte ladetiden (%)", "Ramte døgnets forløb (%)"], rr + 1)
    for area in AREAS:
        for h in HORIZONS:
            m = metrics([r for r in sel if r["område"] == area and r["horisont"] == h])
            if m:
                ws.append([area, h, m["dage"], m["fejl_ore"], m["fejl_pct"], m["skaevhed_ore"],
                           m["ramte_ladetid_pct"], m["ramte_doegn_pct"]])
    for i, w in enumerate([10, 28, 8, 12, 10, 14, 18, 22, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("Pr time på døgnet")
    header(ws2, ["Time", "DK1 kort (øre)", "DK1 lang (øre)", "DK2 kort (øre)", "DK2 lang (øre)"])
    per = {}
    for area in AREAS:
        for name, rng in (("kort", SHORT), ("lang", LONG)):
            m = metrics([r for r in sel if r["område"] == area and r["horisont"] in rng])
            per[(area, name)] = m["fejl_pr_time_ore"] if m else [None] * 24
    for hour in range(24):
        ws2.append([f"{hour:02d}:00"] + [per[(a, n)][hour] for a in AREAS for n in ("kort", "lang")])
    for i in range(1, 6):
        ws2.column_dimensions[get_column_letter(i)].width = 16

    ws3 = wb.create_sheet("Alle timer")
    header(ws3, ["Område", "Dato", "Time", "Prognose lavet", "Dage frem", "elpriser.org (kr)",
                 "Strømligning (kr)", "Appens prognose (kr)", "Rettet prognose (kr)", "Faktisk (kr)",
                 "Fejl (øre)", "Fejl (%)", "Temperatur (°C)", "Vind (m/s)", "Sol (MJ/m²)", "Helligdag"])
    for r in sorted(sel, key=lambda x: (x["område"], x["dato"], x["time"], x["horisont"])):
        err = r["prognose"] - r["faktisk"]
        pct = round(100 * err / abs(r["faktisk"]), 1) if abs(r["faktisk"]) > 0.05 else None
        ws3.append([r["område"], r["dato"], r["time"], r["lavet"], r["horisont"], r["elpriser"],
                    r["stromligning"], round(r["prognose"], 4), round(corr_map[id(r)], 4),
                    round(r["faktisk"], 4), round(100 * err, 1), pct,
                    r.get("temp"), r.get("vind"), r.get("sol"), r.get("helligdag")])
    for i, w in enumerate([8, 12, 6, 14, 10, 16, 16, 18, 18, 12, 10, 10, 14, 11, 12, 10], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = ws3.dimensions

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def update_excel(rows, today):
    """En fil pr. afsluttet uge, plus 'seneste.xlsx' og 'denne-uge.xlsx'."""
    this_monday = today - timedelta(days=today.weekday())
    dates = sorted({r["dato"] for r in rows})
    latest = None
    if dates:
        first = date.fromisoformat(dates[0])
        monday = first - timedelta(days=first.weekday())
        while monday < this_monday:
            y, w, _ = monday.isocalendar()
            name = f"Ladetid-uge-{y}-{w:02d}.xlsx"
            path = os.path.join(EXCEL, name)
            if not os.path.exists(path):
                write_excel(rows, monday, path, f"Ladetid prognosekvalitet uge {w} {y}")
                print(f"Lavede {name}")
            latest = path
            monday += timedelta(days=7)
    y, w, _ = this_monday.isocalendar()
    now_path = os.path.join(EXCEL, "denne-uge.xlsx")
    write_excel(rows, this_monday, now_path, f"Ladetid prognosekvalitet uge {w} {y} (indtil nu)")
    shutil.copyfile(latest or now_path, os.path.join(EXCEL, "seneste.xlsx"))


# ---------------------------------------------------------------- start

def main():
    now = datetime.now(timezone.utc)
    today = now.astimezone(CPH).date()
    save_snapshot(today, now)
    actual = update_actuals(today)
    rows = build_rows(actual)
    print(f"{len(rows)} timer med både prognose og facit")
    snaps_today = {a: load_snapshots(a).get(today.isoformat()) for a in AREAS}
    make_report(rows, today, snaps_today)
    update_excel(rows, today)


if __name__ == "__main__":
    main()
