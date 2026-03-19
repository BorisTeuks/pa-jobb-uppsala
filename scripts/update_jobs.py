"""
PA-jobb Uppsala – Daglig uppdatering  (v3)
─────────────────────────────────────────
Changes from v2:
  • Distance filter  – excludes jobs >50 km from Gunsta
  • Email shows excluded-by-distance towns in a separate section
"""

import re, json, smtplib, os, math
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Configuration ─────────────────────────────────────────────────────────────
EXCEL_FILE   = Path("PA_Jobb_Uppsala_Gunsta.xlsx")
SEEN_FILE    = Path("scripts/seen_jobs.json")
TODAY        = date.today().isoformat()
SHEET_NAME   = "📋 Lediga Jobb"

# Home base
GUNSTA_LAT   = 59.9878
GUNSTA_LON   = 17.7542
MAX_DISTANCE = 50          # km – jobs beyond this are skipped

# ── Location → coordinates lookup ────────────────────────────────────────────
# Covers all towns/areas that regularly appear in Uppsala-region PA ads.
# Keys are lowercase Swedish place names as they appear in Assistanskoll listings.
LOCATION_COORDS = {
    # ── Within ~50 km ✅ ──────────────────────────────────────────────
    "gunsta":        (59.9878, 17.7542),
    "storvreta":     (59.9647, 17.7103),
    "vattholma":     (59.9906, 17.6856),
    "björklinge":    (60.0003, 17.5353),
    "älavsjö":       (59.8700, 17.7200),
    "uppsala":       (59.8586, 17.6389),
    "rickomberga":   (59.8500, 17.5800),
    "eriksberg":     (59.8600, 17.5400),
    "almunge":       (59.9297, 18.0736),
    "sävja":         (59.8200, 17.6800),
    "gottsunda":     (59.8300, 17.6000),
    "stenhagen":     (59.8400, 17.5600),
    "vänge":         (59.8453, 17.5094),
    "länna":         (59.7800, 17.7600),
    "knivsta":       (59.7267, 17.7847),
    "örbyhus":       (60.2500, 17.7100),
    "gimo":          (60.1731, 18.1847),
    "tobo":          (60.3100, 17.6300),   # 36.5 km – just inside
    "tierp":         (60.3426, 17.5140),   # 41.6 km – inside 50
    "tärnsjö":       (60.1486, 16.9356),   # 48.8 km – inside 50
    "heby":          (59.9256, 16.8792),   # 49.2 km – inside 50
    # ── Outside ~50 km ❌ ─────────────────────────────────────────────
    "märsta":        (59.6231, 17.8567),   # 41 km but south, keep
    "sigtuna":       (59.6178, 17.7228),   # 41 km – inside
    "enköping":      (59.6350, 17.0769),   # 54.5 km – OUTSIDE
    "håbo":          (59.5694, 17.5314),   # 48 km – inside
    "bålsta":        (59.5694, 17.5314),   # 48 km – inside
    "rimbo":         (59.7467, 18.3753),   # 43.8 km – inside
    "östhammar":     (60.2563, 18.3700),   # 45.3 km – inside
    "norrtälje":     (59.7580, 18.7061),   # 59 km – OUTSIDE
    "västerås":      (59.6162, 16.5528),   # 79 km – OUTSIDE
    "sala":          (59.9211, 16.6036),   # 64.5 km – OUTSIDE
    "sollentuna":    (59.4281, 17.9511),   # 63 km – OUTSIDE
    "stockholm":     (59.3293, 18.0686),   # 75 km – OUTSIDE
    "täby":          (59.4439, 18.0686),   # 62 km – OUTSIDE
    "upplands väsby":(59.5156, 17.9119),   # 54 km – OUTSIDE
    "nyköping":      (58.7534, 17.0096),   # far – OUTSIDE
    "södertälje":    (59.1956, 17.6253),   # far – OUTSIDE
}

# ── Distance helper ───────────────────────────────────────────────────────────
def haversine(lat1, lon1, lat2, lon2) -> float:
    """Great-circle distance in km."""
    R = 6371
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1))
         * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.asin(math.sqrt(a))

def distance_from_gunsta(ort: str) -> float | None:
    """
    Returns km from Gunsta for a given ort string, or None if unknown.
    Tries substring matching so e.g. 'Knivsta kommun' still matches 'knivsta'.
    """
    ort_low = ort.lower().strip()
    # Exact match first
    if ort_low in LOCATION_COORDS:
        lat, lon = LOCATION_COORDS[ort_low]
        return haversine(GUNSTA_LAT, GUNSTA_LON, lat, lon)
    # Substring match
    for key, (lat, lon) in LOCATION_COORDS.items():
        if key in ort_low or ort_low in key:
            return haversine(GUNSTA_LAT, GUNSTA_LON, lat, lon)
    return None   # unknown location → do NOT exclude (give benefit of doubt)

def distance_label(ort: str) -> str:
    """Returns '~15 km' string for Excel, or '? km' if unknown."""
    d = distance_from_gunsta(ort)
    return f"~{round(d)} km" if d is not None else "? km"

# ── Female exclusion keywords ─────────────────────────────────────────────────
FEMALE_EXCLUDE = [
    "kvinna", "tjej", "flicka", "dam ", "tös",
    " hon ", " hon,", " hon.", "hon är", "hon vill", "hon bor",
    "henne ", " henne,", "hennes ",
    "kvinnlig assistent", "kvinnliga assistenter",
    "söker kvinna", "söker en kvinna", "söker dig som är kvinna",
    "du är kvinna", "du som är kvinna",
    "till kvinna", "åt kvinna", "hos kvinna",
    "till en kvinna", "åt en kvinna", "hos en kvinna",
    "till tjej", "åt tjej", "hos tjej",
    "till en tjej", "åt en tjej", "hos en tjej",
    "till flicka", "till en flicka",
    "kvinna med ms", "ms i centrala", "ms-kvinna",
    "dotter", "syster",
]

MALE_INCLUDE = [
    "kille", "man ", " man,", " man.", "man i ", "man med ", "man som ",
    "till man", "åt man", "hos man", "till en man", "åt en man",
    "pojke", "grabben", "honom ", "hans ", "herr ", "son ", "bror ",
    "årig kille", "årsåldern", "ung man",
    "killar", "manliga sökande", "manlig assistent",
]

# ── Driving-licence keywords ───────────────────────────────────────────────────
LICENCE_REQUIRED = [
    "körkort krävs", "körkort är ett krav", "körkort krav",
    "b-körkort krävs", "b körkort krävs",
    "måste ha körkort", "kräver körkort",
]
LICENCE_MERIT = [
    "körkort är meriterande", "körkort meriterande",
    "körkort önskas", "körkort är en merit", "meriterande med körkort",
    "b-körkort meriterande", "har körkort", "körkort",
]

# ── Weekend / Friday-night keywords ───────────────────────────────────────────
WEEKEND_STRONG = [
    "helg", "helger", "helgpass", "veckoslut",
    "lördag", "söndag", "lördagar", "söndagar",
    "fredag kväll", "fredagskväll", "fredagskvällar",
    "kväll och helg", "kväll & helg", "kvällar och helger",
    "extrajobb",
]
WEEKEND_SOFT = ["kväll", "kvällar", "kvällspass", "deltid", "vid sidan"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def load_seen() -> dict:
    return json.loads(SEEN_FILE.read_text()) if SEEN_FILE.exists() else {}

def save_seen(seen: dict):
    SEEN_FILE.write_text(json.dumps(seen, ensure_ascii=False, indent=2))

def normalise(text: str) -> str:
    return " ".join(text.lower().split())

def is_female(title: str, raw: str = "") -> bool:
    text = normalise(title + " " + raw)
    return any(kw in text for kw in FEMALE_EXCLUDE)

def is_male(title: str, raw: str = "") -> bool:
    if is_female(title, raw):
        return False
    text = normalise(title + " " + raw)
    return any(kw in text for kw in MALE_INCLUDE)

def licence_status(title: str, raw: str = "") -> str:
    text = normalise(title + " " + raw)
    if any(kw in text for kw in LICENCE_REQUIRED): return "required"
    if any(kw in text for kw in LICENCE_MERIT):    return "merit"
    return "none"

def weekend_status(title: str, raw: str = "") -> str:
    text = normalise(title + " " + raw)
    if any(kw in text for kw in WEEKEND_STRONG): return "strong"
    if any(kw in text for kw in WEEKEND_SOFT):   return "soft"
    return "none"

def is_expired(deadline_str: str) -> bool:
    try:
        return date.fromisoformat(deadline_str[:10]) < date.today()
    except Exception:
        return False

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def rate_job(title: str, ort: str, raw: str = "") -> str:
    lic  = licence_status(title, raw)
    wk   = weekend_status(title, raw)
    loc  = normalise(title + " " + ort)
    near = any(k in loc for k in ["knivsta","almunge","storvreta","gunsta","uppsala"])
    if lic == "required":                      return "⭐⭐⭐"
    if lic == "merit" and wk == "strong":      return "⭐⭐⭐"
    if wk == "strong" and near:               return "⭐⭐⭐"
    if wk in ("strong","soft") or lic == "merit" or near: return "⭐⭐"
    return "⭐"

def weekend_label(title: str, raw: str = "") -> str:
    wk = weekend_status(title, raw)
    if wk == "strong": return "✅ Helg/kväll"
    if wk == "soft":   return "🟡 Möjligt"
    return "❓ Okänt"

def licence_label(title: str, raw: str = "") -> str:
    lic = licence_status(title, raw)
    if lic == "required": return "🚗 KRÄVS"
    if lic == "merit":    return "🚗 Merit"
    return "–"

def semester_flag(raw: str, deadline: str, title: str):
    t = normalise(raw + " " + title)
    if "sommar" in t or "feriejobb" in t:
        return "🔴", "Sommarvikariat – troligen behövs under semestern."
    if "tills vidare" in t or "löpande" in t:
        return "🟢", "Tills vidare – semester förhandlas normalt."
    if "behovsanst" in t or "timvikari" in t:
        return "🟢", "Behovsanst. – ta inga pass 8 jun–3 jul."
    if "6 mån" in t:
        if deadline and deadline < "2026-06-01":
            return "🟠", "Deadline före semester – sök INNAN 8 juni."
        return "🟡", "6 mån+ – diskutera semester vid intervju."
    return "🟡", "Kontrollera anställningsform med arbetsgivaren."

# ── Scraping ──────────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )
}

def fetch_assistanskoll_uppsala() -> list[dict]:
    url = "https://assistanskoll.se/platsannonser-i-Uppsala-lan.html"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        print(f"[WARN] Assistanskoll fetch failed: {e}")
        return []

    soup = BeautifulSoup(r.text, "html.parser")
    jobs = []
    for li in soup.select("ul li"):
        a = li.find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        href  = a.get("href", "")
        text  = li.get_text(" ", strip=True)
        m = re.search(r"/annonser/(\d+)", href)
        annons_id = m.group(1) if m else None
        if not annons_id:
            continue
        dead_m = re.search(r"sista ansökningsdag (\d{4}-\d{2}-\d{2})", text)
        pub_m  = re.search(r"Inlämnad till Arbetsförmedlingen (\d{4}-\d{2}-\d{2})", text)
        ort_m  = re.search(r"\(([^.]+)\. Inlämnad", text)
        jobs.append({
            "id":       annons_id,
            "title":    title,
            "url":      f"https://arbetsformedlingen.se/platsbanken/annonser/{annons_id}",
            "deadline": dead_m.group(1) if dead_m else "",
            "pub_date": pub_m.group(1)  if pub_m  else "",
            "ort":      ort_m.group(1).strip() if ort_m else "Uppsala",
            "source":   f"Platsbanken #{annons_id}",
            "raw_text": text,
        })
    print(f"[INFO] Assistanskoll: {len(jobs)} annonser totalt")
    return jobs


def filter_jobs(jobs: list[dict]) -> tuple[list[dict], list[dict]]:
    """
    Returns (kept_jobs, distance_excluded_jobs).
    distance_excluded_jobs is passed to the email so the user can see what was cut.
    """
    kept             = []
    dist_excluded    = []   # jobs rejected only because of distance
    n_female = n_expired = n_no_male = 0

    for j in jobs:
        title = j["title"]
        raw   = j.get("raw_text", "")
        ort   = j.get("ort", "")

        if j["deadline"] and is_expired(j["deadline"]):
            n_expired += 1
            continue

        if is_female(title, raw):
            n_female += 1
            print(f"[SKIP female]   {title[:60]}")
            continue

        if not is_male(title, raw):
            n_no_male += 1
            print(f"[SKIP no-male]  {title[:60]}")
            continue

        # Distance check
        d = distance_from_gunsta(ort)
        if d is not None and d > MAX_DISTANCE:
            dist_excluded.append({**j, "distance_km": round(d)})
            print(f"[SKIP distance] {title[:50]}  ({ort}, {round(d)} km)")
            continue

        # Passed all filters
        j["distance_km"] = round(d) if d is not None else None
        kept.append(j)

    print(f"[INFO] Filter summary: {len(kept)} kept | "
          f"{n_female} female | {n_expired} expired | "
          f"{n_no_male} no-male | {len(dist_excluded)} too far (>{MAX_DISTANCE} km)")
    return kept, dist_excluded


# ── Excel helpers ─────────────────────────────────────────────────────────────
MATCH_COLORS = {
    "⭐⭐⭐": ("1B4332", "D8F3DC"),
    "⭐⭐":   ("1A3A5C", "DBEAFE"),
    "⭐":     ("4A1D1D", "FEE2E2"),
}
SEM_BG = {"🟢":"C8E6C9","🟠":"FFE0B2","🔴":"FFCDD2","🟡":"FFF9C4"}
SEM_FG = {"🟢":"1B5E20","🟠":"E65100","🔴":"B71C1C","🟡":"F57F17"}

def safe_unmerge(ws, rn):
    for ref in [str(m) for m in ws.merged_cells.ranges
                if m.min_row <= rn <= m.max_row]:
        try: ws.unmerge_cells(ref)
        except: pass

def write_cell(ws, row, col, value, bg, fg,
               bold=False, align="left", wrap=True,
               underline=False, hyperlink=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", size=9, bold=bold, color=fg,
                  underline="single" if underline else None)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    c.border = thin_border()
    if hyperlink:
        c.hyperlink = hyperlink

def write_job_row(ws, rn, job, is_new=False):
    stars = job.get("stars", "⭐⭐")
    hbg, rbg = MATCH_COLORS[stars]
    sf   = job.get("sem_flag", "🟡")
    sn   = job.get("sem_note", "Kontrollera med arbetsgivaren.")
    sbg  = SEM_BG.get(sf, "FFFFFF")
    sfg  = SEM_FG.get(sf, "000000")
    d    = job.get("distance_km")
    dist = f"~{d} km" if d is not None else "? km"
    pub  = f"{'🆕 NYTT' + chr(10) if is_new else ''}{job.get('pub_date','?')}\n{job.get('source','')}"

    safe_unmerge(ws, rn)
    write_cell(ws, rn, 1,  ("🆕 " if is_new else "") + stars, hbg,"FFFFFF",True,"center")
    write_cell(ws, rn, 2,  job["title"],                       rbg,"1A3A8C",True,"left",True,True,job["url"])
    write_cell(ws, rn, 3,  job.get("company","–"),             rbg,"000000")
    write_cell(ws, rn, 4,  job.get("ort","Uppsala"),           rbg,"000000")
    write_cell(ws, rn, 5,  job.get("anst","Deltid"),           rbg,"000000")
    write_cell(ws, rn, 6,  job.get("tider","❓ Kontrollera"),  rbg,"000000")
    write_cell(ws, rn, 7,  job.get("korkort","–"),             rbg,"000000",False,"center")
    write_cell(ws, rn, 8,  job.get("deadline","Löpande"),      rbg,"000000",False,"center")
    write_cell(ws, rn, 9,  dist,                               rbg,"000000",False,"center")
    write_cell(ws, rn, 10, job.get("varfor","Ny annons."),     rbg,"000000")
    write_cell(ws, rn, 11, "♂️ Man",                          "DBEAFE","1E40AF",True,"center")
    write_cell(ws, rn, 12, f"{sf}\n{sn}",                     sbg,sfg)
    write_cell(ws, rn, 13, pub,                                rbg,"333333")
    ws.row_dimensions[rn].height = 55


def update_excel(new_jobs, removed_ids):
    if not EXCEL_FILE.exists():
        print(f"[ERROR] {EXCEL_FILE} not found"); return
    wb = load_workbook(str(EXCEL_FILE))
    if SHEET_NAME not in wb.sheetnames:
        print(f"[ERROR] Sheet not found"); return
    ws = wb[SHEET_NAME]

    last_data_row = 5
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        v = row[1].value
        if v and "BORTTAGNA" in str(v): break
        if v: last_data_row = row[0].row

    footer_row = last_data_row + 1
    for job in new_jobs:
        ws.insert_rows(footer_row)
        write_job_row(ws, footer_row, job, is_new=True)
        footer_row += 1

    safe_unmerge(ws, 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    c = ws.cell(row=1, column=1,
        value=(f"📋  Lediga PA-jobb – MANLIGA KUNDER – Uppsala  |  "
               f"✅ UPPDATERAD {TODAY}  |  "
               f"Radie: <{MAX_DISTANCE} km från Gunsta  |  "
               f"{len(new_jobs)} nya tillkomna  |  {len(removed_ids)} utgångna borttagna"))
    c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="1B4332")
    c.alignment = Alignment(horizontal="left", vertical="center")
    wb.save(str(EXCEL_FILE))
    print(f"[INFO] Excel uppdaterad: {len(new_jobs)} nya, {len(removed_ids)} borttagna")


# ── Email digest ──────────────────────────────────────────────────────────────
def send_email(new_jobs: list[dict], removed_ids: list[str],
               dist_excluded: list[dict]):
    sender   = os.environ.get("EMAIL_FROM")
    password = os.environ.get("EMAIL_PASSWORD")
    receiver = os.environ.get("EMAIL_TO")
    if not all([sender, password, receiver]):
        print("[WARN] E-postuppgifter saknas – ingen mejl skickas."); return

    sorted_jobs = sorted(new_jobs, key=lambda j: j.get("stars","⭐"), reverse=True)
    star_color  = {"⭐⭐⭐":"#1B4332","⭐⭐":"#1A3A5C","⭐":"#991B1B"}

    # ── New jobs table rows ───────────────────────────────────────────────────
    job_rows = ""
    for j in sorted_jobs:
        sc  = star_color.get(j.get("stars","⭐"), "#333")
        d   = j.get("distance_km")
        dlabel = f"~{d} km" if d is not None else "? km"
        job_rows += f"""
        <tr>
          <td style="padding:8px 10px;border-bottom:1px solid #eee">
            <span style="font-size:11px;font-weight:bold;color:{sc}">{j.get('stars','⭐')}</span>
            <b> <a href="{j['url']}" style="color:#1A3A8C;text-decoration:none">{j['title']}</a></b><br>
            <span style="color:#888;font-size:11px">{j.get('ort','?')} &middot; {dlabel} &middot; pub. {j.get('pub_date','?')}</span>
          </td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('korkort','–')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('tider','–')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('deadline','Löpande')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center">{j.get('sem_flag','🟡')}</td>
        </tr>"""

    # ── Distance-excluded section ─────────────────────────────────────────────
    # Group by town for a clean summary
    town_summary: dict[str, list] = {}
    for j in dist_excluded:
        key = f"{j.get('ort','?')} ({j.get('distance_km','?')} km)"
        town_summary.setdefault(key, []).append(j)

    dist_rows = ""
    for town, jobs_list in sorted(town_summary.items(),
                                   key=lambda x: int(x[0].split("(")[1].split(" ")[0]) if x[0].split("(")[1].split(" ")[0].isdigit() else 99):
        titles_html = "".join(
            f"<li style='margin:2px 0'>"
            f"<a href='{j['url']}' style='color:#555;font-size:11px'>{j['title']}</a>"
            f"</li>"
            for j in jobs_list
        )
        dist_rows += f"""
        <tr>
          <td style="padding:7px 10px;border-bottom:1px solid #fde8e8;vertical-align:top">
            <b style="color:#7f1d1d">{town}</b>
            <ul style="margin:4px 0 0;padding-left:16px">{titles_html}</ul>
          </td>
        </tr>"""

    dist_section = ""
    if dist_excluded:
        dist_section = f"""
        <h3 style="color:#7f1d1d;margin:24px 0 8px">
          📍 {len(dist_excluded)} jobb exkluderade – för långt bort (&gt;{MAX_DISTANCE} km från Gunsta)
        </h3>
        <table style="width:100%;border-collapse:collapse;background:#fff5f5;border:1px solid #fca5a5;border-radius:6px">
          {dist_rows}
        </table>
        <p style="color:#888;font-size:11px;margin:6px 0 0">
          Dessa annonser matchade man-filter och helg/körkort men ligger utanför {MAX_DISTANCE} km-radien.
          Ändra MAX_DISTANCE i update_jobs.py om du vill inkludera dem.
        </p>"""

    # ── Removed jobs ──────────────────────────────────────────────────────────
    removed_html = "".join(
        f"<li style='color:#888;font-size:12px'>{rid}</li>"
        for rid in removed_ids
    ) or "<li style='color:#aaa'>–</li>"

    # ── Full email HTML ───────────────────────────────────────────────────────
    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:740px;margin:auto;color:#1C1917">
      <!-- Header -->
      <div style="background:#1B4332;color:white;padding:18px 20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0;font-size:18px">📋 PA-jobb Uppsala – {TODAY}</h2>
        <p style="margin:6px 0 0;opacity:0.75;font-size:12px">
          Daglig uppdatering &middot; Manliga kunder &middot; Max {MAX_DISTANCE} km från Gunsta
          &middot; Körkort &amp; helg prioriterat
        </p>
      </div>

      <div style="padding:16px 20px;background:#f9f9f9;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px">

        <!-- New jobs -->
        <h3 style="color:#1B4332;margin:0 0 10px">🆕 {len(new_jobs)} nya jobb sedan igår</h3>
        {'<p style="color:#888;font-size:13px">Inga nya jobb idag.</p>' if not new_jobs else f"""
        <table style="width:100%;border-collapse:collapse;background:white;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden">
          <tr style="background:#1B4332;color:white;font-size:11px">
            <th style="padding:8px 10px;text-align:left">Annons &amp; ort</th>
            <th style="padding:8px;white-space:nowrap">Körkort</th>
            <th style="padding:8px;white-space:nowrap">Helg/kväll</th>
            <th style="padding:8px;white-space:nowrap">Deadline</th>
            <th style="padding:8px;white-space:nowrap">Semester</th>
          </tr>
          {job_rows}
        </table>"""}

        <!-- Distance excluded -->
        {dist_section}

        <!-- Removed (expired) -->
        <h3 style="color:#991B1B;margin:24px 0 6px">🗑️ {len(removed_ids)} utgångna / borttagna</h3>
        <ul style="margin:0;padding-left:18px">{removed_html}</ul>

        <!-- Legend -->
        <div style="margin-top:20px;padding:10px 14px;background:#fff;border:1px solid #e5e7eb;border-radius:6px;font-size:11px;color:#555;line-height:1.8">
          <b>Förklaring:</b><br>
          ⭐⭐⭐ Perfekt match &nbsp;|&nbsp; ⭐⭐ Bra match &nbsp;|&nbsp; ⭐ Möjlig match<br>
          🚗 KRÄVS = körkort obligatoriskt &nbsp;|&nbsp; 🚗 Merit = körkort meriterande<br>
          ✅ Helg/kväll = helgpass nämns explicit &nbsp;|&nbsp; 🟡 Möjligt = kväll/deltid nämns<br>
          📍 Avstånd beräknat från Gunsta ({GUNSTA_LAT}°N, {GUNSTA_LON}°E)
        </div>

        <p style="margin-top:14px;font-size:11px;color:#aaa">
          Excel-fil bifogad &middot; Automatisk körning via GitHub Actions &middot; PA-agent v3
        </p>
      </div>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = (f"PA-jobb Uppsala {TODAY} – "
                      f"{len(new_jobs)} nya jobb"
                      + (f" ({len(dist_excluded)} exkluderade pga avstånd)" if dist_excluded else ""))
    msg["From"]    = sender
    msg["To"]      = receiver
    msg.attach(MIMEText(html, "html"))

    if EXCEL_FILE.exists():
        with open(EXCEL_FILE, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f"attachment; filename={EXCEL_FILE.name}")
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        print(f"[INFO] Mejl skickat till {receiver}")
    except Exception as e:
        print(f"[ERROR] Mejl misslyckades: {e}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print(f"\n{'='*55}")
    print(f"PA-agent v3 · {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"Home: Gunsta · Max radius: {MAX_DISTANCE} km")
    print(f"{'='*55}\n")

    seen = load_seen()
    raw_jobs = fetch_assistanskoll_uppsala()
    filtered_jobs, dist_excluded = filter_jobs(raw_jobs)

    new_jobs    = []
    current_ids = {}

    for j in filtered_jobs:
        aid = j["id"]
        current_ids[aid] = j
        if aid not in seen:
            raw = j.get("raw_text", "")
            j["stars"]    = rate_job(j["title"], j["ort"], raw)
            j["korkort"]  = licence_label(j["title"], raw)
            j["tider"]    = weekend_label(j["title"], raw)
            j["anst"]     = "Deltid"
            sf, sn        = semester_flag(raw, j["deadline"], j["title"])
            j["sem_flag"] = sf
            j["sem_note"] = sn
            j["company"]  = "–"
            j["varfor"]   = (
                "Ny annons! "
                + (f"🚗 Körkort {j['korkort']} · " if j["korkort"] != "–" else "")
                + f"{j['tider']} · Kolla detaljer på Platsbanken."
            )
            new_jobs.append(j)
            d = j.get("distance_km")
            print(f"[NEW {j['stars']}] {j['title'][:50]}  "
                  f"| {j['ort']} {('~'+str(d)+' km') if d else '?'}  "
                  f"| {j['korkort']} | {j['tider']}")

    removed_ids = [aid for aid in seen if aid not in current_ids]
    if removed_ids:
        print(f"[REMOVED] {len(removed_ids)} annonser utgångna")

    new_jobs.sort(key=lambda j: j.get("stars","⭐"), reverse=True)

    seen.update({j["id"]: {"title": j["title"], "seen_date": TODAY}
                 for j in new_jobs})
    for aid in removed_ids:
        seen.pop(aid, None)
    save_seen(seen)

    if new_jobs or removed_ids:
        update_excel(new_jobs, removed_ids)
    else:
        print("[INFO] Inga ändringar – Excel oförändrad.")

    send_email(new_jobs, removed_ids, dist_excluded)

    perfekta = sum(1 for j in new_jobs if j.get("stars") == "⭐⭐⭐")
    print(f"\n✅ Klar! {len(new_jobs)} nya ({perfekta} perfekta), "
          f"{len(dist_excluded)} exkluderade pga avstånd, "
          f"{len(removed_ids)} borttagna.")


if __name__ == "__main__":
    main()
