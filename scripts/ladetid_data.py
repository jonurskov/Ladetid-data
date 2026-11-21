"""
Ladetid data, Jon Pedersen

Kører hver dag på GitHub og gør fem ting:
1. Gemmer de prognoser, appen bruger (elpriser.org og Strømligning), som et "øjebliksbillede".
2. Henter de faktiske spotpriser fra Energi Data Service (Energinet).
3. Sammenligner gamle prognoser med de faktiske priser og måler fejlen.
4. Beregner en rettelse af prognosen og tjekker, om den rammer bedre end den rå prognose.
5. Laver en Excel fil for hver uge med alle fejlene.

Alle priser er spotpris i kr pr. kWh uden moms, tariffer og afgifter.
"""

import csv
import json
import os
import statistics
import sys
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
WINDOW_DAYS = 28                 # rapporten og rettelsen bygger på de sidste 4 uger
CHARGE_HOURS = 4                 # ladetid brugt til "ramte ladetiden"
UA = "Ladetid-data/1.0 (github.com/jonurskov/Ladetid-data)"


# ---------------------------------------------------------------- hjælpere

def http_json(url, tries=4):
    """Henter JSON. Prøver igen, hvis siden siger, at den har for travlt (fx Energi Data Service)."""
    import time
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


def save_snapshot(today, now_utc):
    """Gemmer dagens prognoser. Kun timer der stadig er prognoser, gemmes."""
    for area in AREAS:
        path = os.path.join(DATA, "prognoser", area, f"{today.isoformat()}.json")
        if os.path.exists(path):
            continue  # kun ét øjebliksbillede pr. dag
        snap = {"lavet": now_utc.isoformat(), "dato": today.isoformat(), "kilder": {}}
        for name, fn in (("elpriser", fetch_elpriser), ("stromligning", fetch_stromligning)):
            try:
                snap["kilder"][name] = {d: v for d, v in fn(area).items() if d > today.isoformat()}
            except Exception as e:  # en kilde må gerne fejle
                print(f"Advarsel: {name} {area}: {e}", file=sys.stderr)
        if snap["kilder"]:
            write_json(path, snap)
            print(f"Gemte prognose {area} {today}")


# ---------------------------------------------------------------- 2. faktiske priser

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


# ---------------------------------------------------------------- 3. fejl

def load_snapshots(area):
    folder = os.path.join(DATA, "prognoser", area)
    out = {}
    if not os.path.isdir(folder):
        return out
    for fn in sorted(os.listdir(folder)):
        if fn.endswith(".json"):
            s = read_json(os.path.join(folder, fn))
            if s:
                out[s["dato"]] = s
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
                for hour in range(24):
                    a = act[target][hour]
                    if app[hour] is None or a is None:
                        continue
                    rows.append({
                        "område": area, "lavet": made, "dato": target, "time": hour, "horisont": h,
                        "elpriser": e[hour] if e else None, "stromligning": s[hour] if s else None,
                        "prognose": round(app[hour], 5), "faktisk": a,
                    })
    return rows


# ---------------------------------------------------------------- 4. rettelse

def bias_table(rows, before=None):
    """Typisk fejl (prognose minus facit) pr. horisont og time på døgnet.
    Timerne udglattes over naboer, og tallet trækkes mod 0, når der er få data."""
    buckets = {}
    for r in rows:
        if before and r["dato"] >= before:
            continue
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
            shrink = n_days / (n_days + 7)
            arr.append(round(statistics.median(vals) * shrink, 5))
        table[str(h)] = arr
    return table


def corrected_rows(rows):
    """Rettelse beregnet uden at kigge ind i fremtiden: for hver dag bruges kun dage før."""
    cache = {}
    out = []
    for r in rows:
        key = r["lavet"]  # rettelsen kendes kun med data, der var kendt, da prognosen blev lavet
        if key not in cache:
            cache[key] = bias_table(rows, before=key)
        b = cache[key][str(r["horisont"])][r["time"]]
        out.append(r["prognose"] - b)
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

    # pr. dag og horisont: ramte ladetiden og døgnets forløb
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


def make_report(rows, today):
    start = (today - timedelta(days=WINDOW_DAYS)).isoformat()
    recent = [r for r in rows if r["dato"] >= start]
    corr = corrected_rows(recent) if recent else []
    report = {"lavet": datetime.now(timezone.utc).isoformat(), "periode_fra": start,
              "periode_til": today.isoformat(), "omraader": {}}
    rettelse = {"lavet": report["lavet"], "omraader": {}}
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
        report["omraader"][area] = res
        rettelse["omraader"][area] = {"brug": use, "skaevhed": bias_table(a_rows)}
    write_json(os.path.join(DATA, "rapport.json"), report)
    write_json(os.path.join(DATA, "rettelse.json"), rettelse)
    return report


# ---------------------------------------------------------------- 5. Excel

def write_excel(rows, week_start, path, title):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    week_end = week_start + timedelta(days=6)
    sel = [r for r in rows if week_start.isoformat() <= r["dato"] <= week_end.isoformat()]
    base_from = (week_start - timedelta(days=WINDOW_DAYS)).isoformat()
    base = [r for r in rows if base_from <= r["dato"] <= week_end.isoformat()]
    corr_map = {id(r): c for r, c in zip(base, corrected_rows(base))} if base else {}

    wb = Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F7A4D")

    def header(ws, cols, row=1):
        for i, c in enumerate(cols, 1):
            cell = ws.cell(row=row, column=i, value=c)
            cell.font, cell.fill = head, fill
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Oversigt
    ws = wb.active
    ws.title = "Oversigt"
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Spotpris i kr pr. kWh uden moms. Fejl i øre er den gennemsnitlige afvigelse pr. time. "
                "Fejl i procent er den samlede afvigelse delt med den samlede faktiske pris. "
                "Ramte ladetiden: den billigste 4 timers periode ramt inden for 1 time. "
                "Ramte døgnets forløb: andel af timer, hvor prognosen havde ret i om timen var billig, middel eller dyr.")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 60
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

    # Pr. time på døgnet
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

    # Alle timer
    ws3 = wb.create_sheet("Alle timer")
    header(ws3, ["Område", "Dato", "Time", "Prognose lavet", "Dage frem", "elpriser.org (kr)",
                 "Strømligning (kr)", "Appens prognose (kr)", "Rettet prognose (kr)", "Faktisk (kr)",
                 "Fejl (øre)", "Fejl (%)"])
    for r in sorted(sel, key=lambda x: (x["område"], x["dato"], x["time"], x["horisont"])):
        err = r["prognose"] - r["faktisk"]
        pct = round(100 * err / abs(r["faktisk"]), 1) if abs(r["faktisk"]) > 0.05 else None
        ws3.append([r["område"], r["dato"], r["time"], r["lavet"], r["horisont"], r["elpriser"],
                    r["stromligning"], round(r["prognose"], 4), round(corr_map[id(r)], 4),
                    round(r["faktisk"], 4), round(100 * err, 1), pct])
    for i, w in enumerate([8, 12, 6, 14, 10, 16, 16, 18, 18, 12, 10, 10], 1):
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
            latest = (path, w, y)
            monday += timedelta(days=7)
    y, w, _ = this_monday.isocalendar()
    write_excel(rows, this_monday, os.path.join(EXCEL, "denne-uge.xlsx"),
                f"Ladetid prognosekvalitet uge {w} {y} (indtil nu)")
    if latest:
        import shutil
        shutil.copyfile(latest[0], os.path.join(EXCEL, "seneste.xlsx"))
    else:
        import shutil
        shutil.copyfile(os.path.join(EXCEL, "denne-uge.xlsx"), os.path.join(EXCEL, "seneste.xlsx"))


# ---------------------------------------------------------------- start

def main():
    now = datetime.now(timezone.utc)
    today = now.astimezone(CPH).date()
    save_snapshot(today, now)
    actual = update_actuals(today)
    rows = build_rows(actual)
    print(f"{len(rows)} timer med både prognose og facit")
    make_report(rows, today)
    update_excel(rows, today)


if __name__ == "__main__":
    main()
