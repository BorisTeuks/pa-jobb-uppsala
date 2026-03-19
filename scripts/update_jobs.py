"""
PA-jobb Uppsala – Daglig uppdatering
Söker manliga kunder i Uppsala-regionen på Platsbanken/Assistanskoll.
Uppdaterar Excel-filen och skickar ett mejldigest.
"""

import re
import json
import smtplib
import os
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Configuration ────────────────────────────────────────────────────────────
EXCEL_FILE   = Path("PA_Jobb_Uppsala_Gunsta.xlsx")
SEEN_FILE    = Path("scripts/seen_jobs.json")   # tracks known annons-IDs
TODAY        = date.today().isoformat()
SHEET_NAME   = "📋 Lediga Jobb"

# Männens nyckelord – annonser som matchar är intressanta
MALE_KEYWORDS = [
    "kille", "man", "pojke", "grabben", "honom", "hans",
    "herr", "människa", "son", "bror"
]
# Ord som signalerar KVINNA → skip
FEMALE_SKIP = [
    "kvinna", "tjej", "flicka", "henne", "hennes", "dam",
    "dotter", "syster", "ms-kvinna", "hon ", " hon,",
    "kvinnlig", "kvinnliga", "söker kvinna", "söker en kvinna"
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_seen() -> dict:
    if SEEN_FILE.exists():
        return json.loads(SEEN_FILE.read_text())
    return {}

def save_seen(seen: dict):
    SEEN_FILE.write_text(json.dumps(seen, ensure_ascii=False, indent=2))

def is_male(title: str, desc: str = "") -> bool:
    text = (title + " " + desc).lower()
    if any(w in text for w in FEMALE_SKIP):
        return False
    return any(w in text for w in MALE_KEYWORDS)

def is_expired(deadline_str: str) -> bool:
    """Returns True if deadline has passed."""
    try:
        d = date.fromisoformat(deadline_str[:10])
        return d < date.today()
    except Exception:
        return False

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

# ── Scraping ──────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )
}

def fetch_assistanskoll_uppsala() -> list[dict]:
    """Hämtar alla annonser från Assistanskoll Uppsala-länet."""
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

        # Extract Platsbanken annons-nr from href
        m = re.search(r"/annonser/(\d+)", href)
        annons_id = m.group(1) if m else None
        if not annons_id:
            continue

        # Extract deadline
        dead_match = re.search(r"sista ansökningsdag (\d{4}-\d{2}-\d{2})", text)
        deadline   = dead_match.group(1) if dead_match else ""

        # Extract pub date
        pub_match  = re.search(r"Inlämnad till Arbetsförmedlingen (\d{4}-\d{2}-\d{2})", text)
        pub_date   = pub_match.group(1) if pub_match else ""

        # Extract ort
        ort_match  = re.search(r"\(([^.]+)\. Inlämnad", text)
        ort        = ort_match.group(1).strip() if ort_match else "Uppsala"

        jobs.append({
            "id":        annons_id,
            "title":     title,
            "url":       f"https://arbetsformedlingen.se/platsbanken/annonser/{annons_id}",
            "deadline":  deadline,
            "pub_date":  pub_date,
            "ort":       ort,
            "source":    f"Platsbanken #{annons_id}",
            "raw_text":  text,
        })

    print(f"[INFO] Assistanskoll: {len(jobs)} annonser hittade")
    return jobs


def filter_male_active(jobs: list[dict]) -> list[dict]:
    """Behåll bara manliga kunder med aktiv deadline."""
    result = []
    for j in jobs:
        if j["deadline"] and is_expired(j["deadline"]):
            continue
        if is_male(j["title"], j.get("raw_text", "")):
            result.append(j)
    print(f"[INFO] Efter filtrering (man + aktiv): {len(result)} annonser")
    return result


# ── Excel-uppdatering ─────────────────────────────────────────────────────────

MATCH_COLORS = {
    "⭐⭐⭐": ("1B4332", "D8F3DC"),
    "⭐⭐":   ("1A3A5C", "DBEAFE"),
    "⭐":     ("4A1D1D", "FEE2E2"),
}
SEM_BG = {"🟢": "C8E6C9", "🟠": "FFE0B2", "🔴": "FFCDD2", "🟡": "FFF9C4"}
SEM_FG = {"🟢": "1B5E20", "🟠": "E65100", "🔴": "B71C1C", "🟡": "F57F17"}

# Keywords → star rating for NEW jobs
PERFEKT = ["körkort", "knivsta", "almunge", "storvreta", "gunsta", "extrajobb helg"]
BRA     = ["Uppsala", "helg", "kväll", "deltid", "tills vidare"]

def rate_job(title: str, ort: str) -> str:
    t = title.lower()
    if any(k in t for k in ["körkort krävs", "körkort krav"]):
        return "⭐⭐⭐"
    if any(k in (t + ort.lower()) for k in ["knivsta", "almunge", "storvreta", "extrajobb"]):
        return "⭐⭐⭐"
    if any(k in t for k in ["helg", "kväll", "tills vidare", "deltid"]):
        return "⭐⭐"
    return "⭐"

def semester_flag(anst: str, deadline: str, title: str) -> tuple[str, str]:
    t = (anst + title).lower()
    if "sommar" in t or "feriejobb" in t:
        return "🔴", "Sommarvikariat = troligen behövs under din semester (8 jun–3 jul)."
    if "tills vidare" in t or "löpande" in t:
        return "🟢", "Tills vidare = semester förhandlas normalt."
    if "behovsanst" in t or "timvikari" in t:
        return "🟢", "Behovsanst. = ta inga pass 8 jun–3 jul."
    if "6 mån" in t:
        if deadline and deadline < "2026-06-01":
            return "🟠", "Deadline före semester – sök INNAN 8 juni."
        return "🟡", "6 mån+ – diskutera semester vid intervju."
    return "🟡", "Kontrollera anställningsform med arbetsgivaren."

def safe_unmerge(ws, rn: int):
    to_remove = [str(m) for m in ws.merged_cells.ranges
                 if m.min_row <= rn <= m.max_row]
    for ref in to_remove:
        try:
            ws.unmerge_cells(ref)
        except Exception:
            pass

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
    sf, sn   = job.get("sem_flag", "🟡"), job.get("sem_note", "Kontrollera med arbetsgivaren.")
    sbg, sfg = SEM_BG.get(sf, "FFFFFF"), SEM_FG.get(sf, "000000")

    new_tag = "🆕 NYTT\n" if is_new else ""
    pub_text = f"{new_tag}{job.get('pub_date','?')}\n{job.get('source','')}"

    safe_unmerge(ws, rn)
    write_cell(ws, rn, 1,  ("🆕 " if is_new else "") + stars, hbg,"FFFFFF",True,"center")
    write_cell(ws, rn, 2,  job["title"],  rbg,"1A3A8C",True,"left",True,True, job["url"])
    write_cell(ws, rn, 3,  job.get("company","–"), rbg,"000000")
    write_cell(ws, rn, 4,  job.get("ort","Uppsala"), rbg,"000000")
    write_cell(ws, rn, 5,  job.get("anst","Deltid"), rbg,"000000")
    write_cell(ws, rn, 6,  job.get("tider","🗓️ Helg + kväll"), rbg,"000000")
    write_cell(ws, rn, 7,  job.get("korkort","–"), rbg,"000000",False,"center")
    write_cell(ws, rn, 8,  job.get("deadline","Löpande"), rbg,"000000",False,"center")
    write_cell(ws, rn, 9,  job.get("avstand","–"), rbg,"000000",False,"center")
    write_cell(ws, rn, 10, job.get("varfor","Ny annons – kolla detaljer på Platsbanken."), rbg,"000000")
    write_cell(ws, rn, 11, "♂️ Man","DBEAFE","1E40AF",True,"center")
    write_cell(ws, rn, 12, f"{sf}\n{sn}", sbg, sfg)
    write_cell(ws, rn, 13, pub_text, rbg,"333333")
    ws.row_dimensions[rn].height = 55


def update_excel(new_jobs: list[dict], removed_ids: list[str]):
    """Adds new jobs to Excel and marks expired ones."""
    if not EXCEL_FILE.exists():
        print(f"[ERROR] {EXCEL_FILE} not found – skipping Excel update.")
        return

    wb = load_workbook(str(EXCEL_FILE))
    if SHEET_NAME not in wb.sheetnames:
        print(f"[ERROR] Sheet '{SHEET_NAME}' not found.")
        return
    ws = wb[SHEET_NAME]

    # Find last data row (before footer)
    last_data_row = 5
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        v = row[1].value
        if v and "BORTTAGNA" in str(v):
            break
        if v:
            last_data_row = row[0].row

    # Insert new jobs just before footer
    footer_row = last_data_row + 1
    for i, job in enumerate(new_jobs):
        rn = footer_row + i
        ws.insert_rows(rn)
        write_job_row(ws, rn, job, is_new=True)
        footer_row += 1  # shift for next insertion

    # Update title
    title_val = (
        f"📋  Lediga PA-jobb – MANLIGA KUNDER – Uppsala  |  "
        f"✅ UPPDATERAD {TODAY}  |  "
        f"{len(new_jobs)} nya tillkomna  |  "
        f"{len(removed_ids)} utgångna borttagna"
    )
    safe_unmerge(ws, 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    c = ws.cell(row=1, column=1, value=title_val)
    c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="1B4332")
    c.alignment = Alignment(horizontal="left", vertical="center")

    wb.save(str(EXCEL_FILE))
    print(f"[INFO] Excel uppdaterad: {len(new_jobs)} nya, {len(removed_ids)} borttagna")


# ── E-postdigest ───────────────────────────────────────────────────────────────

def send_email(new_jobs: list[dict], removed_ids: list[str]):
    sender   = os.environ.get("EMAIL_FROM")
    password = os.environ.get("EMAIL_PASSWORD")
    receiver = os.environ.get("EMAIL_TO")

    if not all([sender, password, receiver]):
        print("[WARN] E-postuppgifter saknas – ingen mejl skickas.")
        return

    subject = f"PA-jobb Uppsala {TODAY} – {len(new_jobs)} nya jobb"

    # HTML-body
    new_rows = ""
    for j in new_jobs:
        sf = j.get("sem_flag", "🟡")
        new_rows += f"""
        <tr>
          <td style="padding:8px;border-bottom:1px solid #eee">
            <b><a href="{j['url']}" style="color:#1A3A8C">{j['title']}</a></b><br>
            <span style="color:#666;font-size:12px">{j.get('ort','?')} · {j.get('pub_date','?')} · {j.get('source','')}</span>
          </td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center">{j.get('deadline','Löpande')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center">{sf}</td>
        </tr>"""

    removed_list = "".join(f"<li>{rid}</li>" for rid in removed_ids) or "<li>–</li>"

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:700px;margin:auto">
      <div style="background:#1B4332;color:white;padding:16px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">📋 PA-jobb Uppsala – {TODAY}</h2>
        <p style="margin:4px 0 0;opacity:0.8">Daglig uppdatering · Manliga kunder · Gunsta-området</p>
      </div>

      <div style="padding:16px;background:#f9f9f9">
        <h3 style="color:#1B4332">🆕 {len(new_jobs)} nya jobb sedan igår</h3>
        {'<p style="color:#666">Inga nya jobb idag.</p>' if not new_jobs else f'''
        <table style="width:100%;border-collapse:collapse">
          <tr style="background:#1B4332;color:white">
            <th style="padding:8px;text-align:left">Annons</th>
            <th style="padding:8px">Deadline</th>
            <th style="padding:8px">Semester</th>
          </tr>
          {new_rows}
        </table>'''}

        <h3 style="color:#B71C1C;margin-top:24px">🗑️ {len(removed_ids)} utgångna (deadline passerad)</h3>
        <ul style="color:#666">{removed_list}</ul>

        <p style="margin-top:24px;font-size:12px;color:#999">
          Se bifogad Excel för fullständig lista med semester-flaggor och avstånd.<br>
          Automatisk körning via GitHub Actions · Boris Teuks PA-agent
        </p>
      </div>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = sender
    msg["To"]      = receiver
    msg.attach(MIMEText(html, "html"))

    # Bifoga Excel
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


# ── Huvudflöde ────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*50}")
    print(f"PA-agent körning: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*50}\n")

    seen = load_seen()

    # 1. Hämta och filtrera annonser
    raw_jobs  = fetch_assistanskoll_uppsala()
    male_jobs = filter_male_active(raw_jobs)

    # 2. Hitta NYA (inte sedda förut)
    new_jobs = []
    current_ids = {}

    for j in male_jobs:
        aid = j["id"]
        current_ids[aid] = j

        if aid not in seen:
            # Berika jobbet
            j["stars"]    = rate_job(j["title"], j["ort"])
            j["anst"]     = "Deltid"  # default – kan förbättras med annonsdetails
            sf, sn        = semester_flag(j.get("raw_text",""), j["deadline"], j["title"])
            j["sem_flag"] = sf
            j["sem_note"] = sn
            j["company"]  = "–"  # Assistanskoll visar ej alltid bolag i listelement
            j["varfor"]   = f"Ny annons! Kolla detaljer på Platsbanken."
            j["avstand"]  = "–"
            j["korkort"]  = "✅ KRÄVS" if "körkort" in j["title"].lower() else "–"
            j["tider"]    = "🗓️ Kontrollera annons"
            new_jobs.append(j)
            print(f"[NEW] {j['title'][:60]}")

    # 3. Hitta UTGÅNGNA (var i seen, finns inte längre)
    removed_ids = [
        aid for aid in seen
        if aid not in current_ids
    ]
    if removed_ids:
        print(f"[REMOVED] {len(removed_ids)} annonser utgångna: {removed_ids}")

    # 4. Uppdatera seen-listan
    seen.update({j["id"]: {"title": j["title"], "seen_date": TODAY}
                 for j in new_jobs})
    for aid in removed_ids:
        seen.pop(aid, None)
    save_seen(seen)

    # 5. Uppdatera Excel
    if new_jobs or removed_ids:
        update_excel(new_jobs, removed_ids)
    else:
        print("[INFO] Inga ändringar – Excel oförändrad.")

    # 6. Skicka mejl (alltid, för daglig rapport)
    send_email(new_jobs, removed_ids)

    print(f"\n✅ Klar! {len(new_jobs)} nya, {len(removed_ids)} borttagna.")


if __name__ == "__main__":
    main()
