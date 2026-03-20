"""
PA-jobb Uppsala – Daglig uppdatering  (v4)
──────────────────────────────────────────
Changes from v3:
  • Generates docs/index.html  – live dashboard published via GitHub Pages
  • Email becomes a short digest + link to dashboard
  • Still-open section added (jobs seen before today still active)
  • seen_jobs.json now stores first_seen date per job
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

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_FILE   = Path("PA_Jobb_Uppsala_Gunsta.xlsx")
SEEN_FILE    = Path("scripts/seen_jobs.json")
DASHBOARD    = Path("docs/index.html")
TODAY        = date.today().isoformat()
SHEET_NAME   = "📋 Lediga Jobb"
PAGES_URL    = "https://boristeuks.github.io/pa-jobb-uppsala"
GUNSTA_LAT   = 59.9878
GUNSTA_LON   = 17.7542
MAX_KM       = 50

# ── Location coords ───────────────────────────────────────────────────────────
LOCATION_COORDS = {
    "gunsta":(59.9878,17.7542),"storvreta":(59.9647,17.7103),
    "vattholma":(59.9906,17.6856),"björklinge":(60.0003,17.5353),
    "älavsjö":(59.8700,17.7200),"uppsala":(59.8586,17.6389),
    "rickomberga":(59.8500,17.5800),"eriksberg":(59.8600,17.5400),
    "almunge":(59.9297,18.0736),"sävja":(59.8200,17.6800),
    "gottsunda":(59.8300,17.6000),"stenhagen":(59.8400,17.5600),
    "vänge":(59.8453,17.5094),"länna":(59.7800,17.7600),
    "knivsta":(59.7267,17.7847),"örbyhus":(60.2500,17.7100),
    "gimo":(60.1731,18.1847),"tobo":(60.3100,17.6300),
    "tierp":(60.3426,17.5140),"tärnsjö":(60.1486,16.9356),
    "heby":(59.9256,16.8792),"märsta":(59.6231,17.8567),
    "sigtuna":(59.6178,17.7228),"håbo":(59.5694,17.5314),
    "bålsta":(59.5694,17.5314),"rimbo":(59.7467,18.3753),
    "östhammar":(60.2563,18.3700),"enköping":(59.6350,17.0769),
    "norrtälje":(59.7580,18.7061),"västerås":(59.6162,16.5528),
    "sala":(59.9211,16.6036),"stockholm":(59.3293,18.0686),
}

FEMALE_EXCLUDE = [
    "kvinna","tjej","flicka","dam ","tös"," hon "," hon,"," hon.",
    "hon är","hon vill","hon bor","henne "," henne,","hennes ",
    "kvinnlig assistent","kvinnliga assistenter","söker kvinna",
    "söker en kvinna","du är kvinna","du som är kvinna",
    "till kvinna","åt kvinna","hos kvinna","till en kvinna",
    "åt en kvinna","hos en kvinna","till tjej","åt tjej","hos tjej",
    "till en tjej","åt en tjej","hos en tjej","till flicka",
    "till en flicka","kvinna med ms","ms i centrala","ms-kvinna",
    "dotter","syster",
]
MALE_INCLUDE = [
    "kille","man "," man,"," man.","man i ","man med ","man som ",
    "till man","åt man","hos man","till en man","åt en man",
    "pojke","grabben","honom ","hans ","herr ","son ","bror ",
    "årig kille","årsåldern","ung man","killar",
    "manliga sökande","manlig assistent",
]
LICENCE_REQUIRED = [
    "körkort krävs","körkort är ett krav","körkort krav",
    "b-körkort krävs","måste ha körkort","kräver körkort",
]
LICENCE_MERIT = [
    "körkort är meriterande","körkort meriterande","körkort önskas",
    "körkort är en merit","meriterande med körkort",
    "b-körkort meriterande","körkort",
]
WEEKEND_STRONG = [
    "helg","helger","helgpass","lördag","söndag","lördagar","söndagar",
    "fredag kväll","fredagskväll","kväll och helg","kväll & helg",
    "kvällar och helger","extrajobb",
]
WEEKEND_SOFT = ["kväll","kvällar","kvällspass","deltid","vid sidan"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def load_seen():
    return json.loads(SEEN_FILE.read_text()) if SEEN_FILE.exists() else {}

def save_seen(seen):
    SEEN_FILE.write_text(json.dumps(seen, ensure_ascii=False, indent=2))

def norm(t): return " ".join(t.lower().split())

def is_female(title, raw=""):
    text = norm(title+" "+raw)
    return any(k in text for k in FEMALE_EXCLUDE)

def is_male(title, raw=""):
    if is_female(title, raw): return False
    text = norm(title+" "+raw)
    return any(k in text for k in MALE_INCLUDE)

def lic_status(title, raw=""):
    text = norm(title+" "+raw)
    if any(k in text for k in LICENCE_REQUIRED): return "required"
    if any(k in text for k in LICENCE_MERIT):    return "merit"
    return "none"

def wk_status(title, raw=""):
    text = norm(title+" "+raw)
    if any(k in text for k in WEEKEND_STRONG): return "strong"
    if any(k in text for k in WEEKEND_SOFT):   return "soft"
    return "none"

def haversine(lat2, lon2):
    R=6371; dlat=math.radians(lat2-GUNSTA_LAT); dlon=math.radians(lon2-GUNSTA_LON)
    a=math.sin(dlat/2)**2+math.cos(math.radians(GUNSTA_LAT))*math.cos(math.radians(lat2))*math.sin(dlon/2)**2
    return R*2*math.asin(math.sqrt(a))

def dist_km(ort):
    o = ort.lower().strip()
    for key,(lat,lon) in LOCATION_COORDS.items():
        if key in o or o in key:
            return round(haversine(lat,lon))
    return None

def is_expired(d):
    try: return date.fromisoformat(d[:10]) < date.today()
    except: return False

def lic_label(title, raw=""):
    s = lic_status(title, raw)
    return "🚗 KRÄVS" if s=="required" else ("🚗 Merit" if s=="merit" else "–")

def wk_label(title, raw=""):
    s = wk_status(title, raw)
    return "✅ Helg/kväll" if s=="strong" else ("🟡 Möjligt" if s=="soft" else "❓")

def rate(title, ort, raw=""):
    lic = lic_status(title,raw); wk = wk_status(title,raw)
    loc = norm(title+" "+ort)
    near = any(k in loc for k in ["knivsta","almunge","storvreta","gunsta","uppsala"])
    if lic=="required": return "⭐⭐⭐"
    if lic=="merit" and wk=="strong": return "⭐⭐⭐"
    if wk=="strong" and near: return "⭐⭐⭐"
    if wk in ("strong","soft") or lic=="merit" or near: return "⭐⭐"
    return "⭐"

def sem_flag(raw, deadline, title):
    t = norm(raw+" "+title)
    if "sommar" in t or "feriejobb" in t: return "🔴","Sommarvikariat."
    if "tills vidare" in t or "löpande" in t: return "🟢","Tills vidare."
    if "behovsanst" in t or "timvikari" in t: return "🟢","Behovsanst."
    if "6 mån" in t:
        if deadline and deadline<"2026-06-01": return "🟠","Sök INNAN semestern."
        return "🟡","Diskutera semester."
    return "🟡","Kontrollera."

def days_since(date_str):
    try:
        d = date.fromisoformat(date_str[:10])
        return (date.today()-d).days
    except: return 0

# ── Scraping ──────────────────────────────────────────────────────────────────
HEADERS={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36"}

def fetch():
    url="https://assistanskoll.se/platsannonser-i-Uppsala-lan.html"
    try:
        r=requests.get(url,headers=HEADERS,timeout=15); r.raise_for_status()
    except Exception as e:
        print(f"[WARN] fetch failed: {e}"); return []
    soup=BeautifulSoup(r.text,"html.parser"); jobs=[]
    for li in soup.select("ul li"):
        a=li.find("a")
        if not a: continue
        title=a.get_text(strip=True); href=a.get("href",""); text=li.get_text(" ",strip=True)
        m=re.search(r"/annonser/(\d+)",href)
        if not m: continue
        aid=m.group(1)
        dead=re.search(r"sista ansökningsdag (\d{4}-\d{2}-\d{2})",text)
        pub=re.search(r"Inlämnad till Arbetsförmedlingen (\d{4}-\d{2}-\d{2})",text)
        ort=re.search(r"\(([^.]+)\. Inlämnad",text)
        jobs.append({"id":aid,"title":title,
            "url":f"https://arbetsformedlingen.se/platsbanken/annonser/{aid}",
            "deadline":dead.group(1) if dead else "",
            "pub_date":pub.group(1) if pub else "",
            "ort":ort.group(1).strip() if ort else "Uppsala",
            "source":f"Platsbanken #{aid}","raw_text":text})
    print(f"[INFO] Fetched {len(jobs)} listings")
    return jobs

def filter_jobs(jobs):
    kept=[]; dist_excl=[]; nf=ne=nm=0
    for j in jobs:
        if j["deadline"] and is_expired(j["deadline"]): ne+=1; continue
        if is_female(j["title"],j.get("raw_text","")): nf+=1; continue
        if not is_male(j["title"],j.get("raw_text","")): nm+=1; continue
        d=dist_km(j.get("ort",""))
        if d is not None and d>MAX_KM:
            dist_excl.append({**j,"distance_km":d}); continue
        j["distance_km"]=d; kept.append(j)
    print(f"[INFO] Kept {len(kept)} | female={nf} expired={ne} no-male={nm} far={len(dist_excl)}")
    return kept, dist_excl

def enrich(j, seen):
    raw=j.get("raw_text","")
    j["stars"]   = rate(j["title"],j["ort"],raw)
    j["korkort"] = lic_label(j["title"],raw)
    j["tider"]   = wk_label(j["title"],raw)
    j["anst"]    = "Deltid"
    sf,sn        = sem_flag(raw,j["deadline"],j["title"])
    j["sem_flag"]= sf; j["sem_note"]=sn
    j["company"] = "–"; j["avstand"]=f"~{j['distance_km']} km" if j.get("distance_km") else "? km"
    j["first_seen"] = seen.get(j["id"],{}).get("first_seen", TODAY)
    return j

# ── Excel ─────────────────────────────────────────────────────────────────────
MCOL={"⭐⭐⭐":("1B4332","D8F3DC"),"⭐⭐":("1A3A5C","DBEAFE"),"⭐":("4A1D1D","FEE2E2")}
SBG={"🟢":"C8E6C9","🟠":"FFE0B2","🔴":"FFCDD2","🟡":"FFF9C4"}
SFG={"🟢":"1B5E20","🟠":"E65100","🔴":"B71C1C","🟡":"F57F17"}

def tb():
    s=Side(style="thin",color="CCCCCC"); return Border(left=s,right=s,top=s,bottom=s)

def wc(ws,row,col,v,bg,fg,bold=False,align="left",wrap=True,ul=False,hl=None):
    c=ws.cell(row=row,column=col,value=v)
    c.font=Font(name="Arial",size=9,bold=bold,color=fg,underline="single" if ul else None)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(wrap_text=wrap,vertical="top",horizontal=align)
    c.border=tb()
    if hl: c.hyperlink=hl

def safe_unmerge(ws,rn):
    for ref in [str(m) for m in ws.merged_cells.ranges if m.min_row<=rn<=m.max_row]:
        try: ws.unmerge_cells(ref)
        except: pass

def write_row(ws,rn,j,is_new=False):
    stars=j.get("stars","⭐⭐"); hbg,rbg=MCOL[stars]
    sf=j.get("sem_flag","🟡"); sn=j.get("sem_note","")
    sbg=SBG.get(sf,"FFFFFF"); sfg=SFG.get(sf,"000000")
    pub=f"{'🆕 NYTT'+chr(10) if is_new else ''}{j.get('pub_date','?')}\n{j.get('source','')}"
    safe_unmerge(ws,rn)
    wc(ws,rn,1,("🆕 " if is_new else "")+stars,hbg,"FFFFFF",True,"center")
    wc(ws,rn,2,j["title"],rbg,"1A3A8C",True,"left",True,True,j["url"])
    wc(ws,rn,3,j.get("company","–"),rbg,"000000")
    wc(ws,rn,4,j.get("ort","Uppsala"),rbg,"000000")
    wc(ws,rn,5,j.get("anst","Deltid"),rbg,"000000")
    wc(ws,rn,6,j.get("tider","❓"),rbg,"000000")
    wc(ws,rn,7,j.get("korkort","–"),rbg,"000000",False,"center")
    wc(ws,rn,8,j.get("deadline","Löpande"),rbg,"000000",False,"center")
    wc(ws,rn,9,j.get("avstand","–"),rbg,"000000",False,"center")
    wc(ws,rn,10,f"Ny annons! {j.get('korkort','–')} · {j.get('tider','–')}",rbg,"000000")
    wc(ws,rn,11,"♂️ Man","DBEAFE","1E40AF",True,"center")
    wc(ws,rn,12,f"{sf}\n{sn}",sbg,sfg)
    wc(ws,rn,13,pub,rbg,"333333")
    ws.row_dimensions[rn].height=55

def update_excel(new_jobs,removed_ids):
    if not EXCEL_FILE.exists(): print("[ERROR] Excel not found"); return
    wb=load_workbook(str(EXCEL_FILE))
    if SHEET_NAME not in wb.sheetnames: print("[ERROR] Sheet not found"); return
    ws=wb[SHEET_NAME]
    last=5
    for row in ws.iter_rows(min_row=5,max_row=ws.max_row):
        v=row[1].value
        if v and "BORTTAGNA" in str(v): break
        if v: last=row[0].row
    footer=last+1
    for j in new_jobs:
        ws.insert_rows(footer); write_row(ws,footer,j,is_new=True); footer+=1
    safe_unmerge(ws,1); ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=13)
    c=ws.cell(row=1,column=1,
        value=f"📋 PA-jobb Uppsala · UPPDATERAD {TODAY} · Radie <{MAX_KM} km · {len(new_jobs)} nya · {len(removed_ids)} borttagna")
    c.font=Font(name="Arial",size=11,bold=True,color="FFFFFF")
    c.fill=PatternFill("solid",start_color="1B4332")
    c.alignment=Alignment(horizontal="left",vertical="center")
    wb.save(str(EXCEL_FILE))
    print(f"[INFO] Excel updated")

# ── HTML Dashboard ────────────────────────────────────────────────────────────
def star_color(s): return {"⭐⭐⭐":"#1B4332","⭐⭐":"#1A3A5C","⭐":"#991B1B"}.get(s,"#333")

def job_card(j, badge_html=""):
    sc   = star_color(j.get("stars","⭐"))
    d    = j.get("distance_km"); dist = f"~{d} km" if d else "?"
    days = days_since(j.get("first_seen", TODAY))
    dl   = j.get("deadline","Löpande")
    deadline_style = ' style="color:#991B1B;font-weight:600"' if dl < TODAY else ""
    return f"""
    <div class="card">
      <div class="card-top">
        <span class="stars" style="color:{sc}">{j.get('stars','⭐')}</span>
        <a href="{j['url']}" target="_blank" class="job-title">{j['title']}</a>
        {badge_html}
      </div>
      <div class="card-meta">
        <span>📍 {j.get('ort','?')} &middot; {dist}</span>
        <span>🚗 {j.get('korkort','–')}</span>
        <span>{j.get('tider','–')}</span>
        <span{deadline_style}>⏰ {dl}</span>
        <span>{j.get('sem_flag','🟡')}</span>
        {"<span class='days-open'>öppen sedan "+str(days)+" dagar</span>" if days>0 else ""}
      </div>
    </div>"""

def build_dashboard(new_jobs, open_jobs, closed_jobs, dist_excl):
    DASHBOARD.parent.mkdir(exist_ok=True)

    new_html   = "".join(job_card(j,'<span class="badge badge-new">NY IDAG</span>') for j in new_jobs) or "<p class='empty'>Inga nya jobb idag.</p>"
    open_html  = "".join(job_card(j) for j in open_jobs)  or "<p class='empty'>Inga öppna jobb från tidigare dagar.</p>"
    closed_html= "".join(
        f'<div class="card card-closed"><span class="job-title-closed">{j.get("title","?")}</span>'
        f'<span class="badge badge-closed">STÄNGD</span></div>'
        for j in closed_jobs) or "<p class='empty'>Inga nyligen stängda.</p>"

    dist_rows=""
    for j in dist_excl:
        dist_rows+=f'<div class="dist-row"><a href="{j["url"]}" target="_blank">{j["title"]}</a> <span class="dist-badge">{j.get("ort","?")} · ~{j.get("distance_km","?")} km</span></div>'

    html = f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PA-jobb Uppsala – {TODAY}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f4;color:#1c1917;line-height:1.5}}
  header{{background:#1B4332;color:white;padding:20px 24px}}
  header h1{{font-size:20px;font-weight:600}}
  header p{{opacity:.75;font-size:13px;margin-top:4px}}
  .stats{{display:flex;gap:12px;margin-top:14px;flex-wrap:wrap}}
  .stat{{background:rgba(255,255,255,.15);border-radius:8px;padding:8px 14px;font-size:13px;font-weight:500}}
  .stat span{{font-size:20px;font-weight:700;display:block}}
  main{{max-width:860px;margin:24px auto;padding:0 16px 40px}}
  .section{{margin-bottom:28px}}
  .section-title{{font-size:15px;font-weight:600;margin-bottom:12px;display:flex;align-items:center;gap:8px}}
  .card{{background:white;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:8px}}
  .card-top{{display:flex;align-items:flex-start;gap:8px;flex-wrap:wrap}}
  .stars{{font-size:12px;font-weight:700;flex-shrink:0;padding-top:2px}}
  .job-title{{color:#1A3A5C;font-weight:600;font-size:14px;text-decoration:none;flex:1}}
  .job-title:hover{{text-decoration:underline}}
  .card-meta{{display:flex;gap:10px;margin-top:8px;font-size:12px;color:#57534e;flex-wrap:wrap}}
  .days-open{{color:#854F0B;background:#FEF9C3;padding:1px 7px;border-radius:99px;font-size:11px}}
  .badge{{padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600;flex-shrink:0}}
  .badge-new{{background:#D8F3DC;color:#1B4332}}
  .badge-closed{{background:#FEE2E2;color:#991B1B}}
  .card-closed{{display:flex;align-items:center;gap:10px;opacity:.6}}
  .job-title-closed{{text-decoration:line-through;color:#57534e;font-size:13px;flex:1}}
  .dist-section{{background:#FFF5F5;border:1px solid #FECACA;border-radius:10px;padding:14px 16px;margin-bottom:28px}}
  .dist-title{{font-size:13px;font-weight:600;color:#991B1B;margin-bottom:8px}}
  .dist-row{{padding:5px 0;border-bottom:1px solid #FEE2E2;font-size:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}}
  .dist-row:last-child{{border-bottom:none}}
  .dist-row a{{color:#1A3A5C}}
  .dist-badge{{background:#FEE2E2;color:#991B1B;padding:1px 7px;border-radius:99px;font-size:11px;white-space:nowrap}}
  .empty{{color:#a8a29e;font-size:13px;padding:8px 0}}
  .legend{{background:white;border:1px solid #e5e7eb;border-radius:10px;padding:12px 16px;font-size:12px;color:#57534e;margin-top:8px;line-height:2}}
  footer{{text-align:center;font-size:11px;color:#a8a29e;padding:16px}}
  @media(max-width:600px){{.stats{{gap:8px}}.stat{{padding:6px 10px}}}}
</style>
</head>
<body>
<header>
  <h1>📋 PA-jobb Uppsala</h1>
  <p>Uppdaterad {TODAY} &middot; Manliga kunder &middot; Max {MAX_KM} km från Gunsta &middot; Körkort &amp; helg prioriterat</p>
  <div class="stats">
    <div class="stat"><span>{len(new_jobs)}</span>Nya idag</div>
    <div class="stat"><span>{len(open_jobs)}</span>Fortfarande öppna</div>
    <div class="stat"><span>{len(closed_jobs)}</span>Nyligen stängda</div>
    <div class="stat"><span>{len(dist_excl)}</span>Exkluderade (avstånd)</div>
  </div>
</header>

<main>

  <!-- NEW TODAY -->
  <div class="section">
    <div class="section-title" style="color:#1B4332">🆕 Nya jobb idag ({len(new_jobs)})</div>
    {new_html}
  </div>

  <!-- STILL OPEN -->
  <div class="section">
    <div class="section-title" style="color:#1A3A5C">📂 Fortfarande öppna ({len(open_jobs)})</div>
    {open_html}
  </div>

  <!-- CLOSED -->
  <div class="section">
    <div class="section-title" style="color:#991B1B">🗑️ Nyligen stängda ({len(closed_jobs)})</div>
    {closed_html}
  </div>

  <!-- DISTANCE EXCLUDED -->
  {"" if not dist_excl else f'''
  <div class="dist-section">
    <div class="dist-title">📍 Exkluderade – för långt bort (&gt;{MAX_KM} km från Gunsta) ({len(dist_excl)})</div>
    {dist_rows}
  </div>'''}

  <!-- LEGEND -->
  <div class="legend">
    <strong>Förklaring:</strong>
    ⭐⭐⭐ Perfekt match &nbsp;|&nbsp; ⭐⭐ Bra match &nbsp;|&nbsp; ⭐ Möjlig match &nbsp;|&nbsp;
    🚗 KRÄVS = körkort obligatoriskt &nbsp;|&nbsp; 🚗 Merit = körkort meriterande &nbsp;|&nbsp;
    ✅ Helg/kväll = helgpass nämns &nbsp;|&nbsp; 🟡 Möjligt = kväll/deltid nämns &nbsp;|&nbsp;
    🟢 Tills vidare &nbsp;|&nbsp; 🟠 Sök innan semester &nbsp;|&nbsp; 🔴 Sommarvikariat
  </div>

</main>

<footer>PA-agent v4 &middot; GitHub Actions &middot; Boris Teuks &middot; {TODAY}</footer>
</body>
</html>"""

    DASHBOARD.write_text(html, encoding="utf-8")
    print(f"[INFO] Dashboard written → {DASHBOARD}")

# ── Email ─────────────────────────────────────────────────────────────────────
def send_email(new_jobs, open_jobs, closed_jobs, dist_excl):
    sender=os.environ.get("EMAIL_FROM"); pw=os.environ.get("EMAIL_PASSWORD"); rcv=os.environ.get("EMAIL_TO")
    if not all([sender,pw,rcv]): print("[WARN] No email credentials"); return

    sc={"⭐⭐⭐":"#1B4332","⭐⭐":"#1A3A5C","⭐":"#991B1B"}
    rows=""
    for j in sorted(new_jobs,key=lambda x:x.get("stars","⭐"),reverse=True):
        c=sc.get(j.get("stars","⭐"),"#333")
        d=j.get("distance_km"); dl=f"~{d} km" if d else "?"
        rows+=f"""<tr>
          <td style="padding:8px 10px;border-bottom:1px solid #eee">
            <span style="color:{c};font-weight:700;font-size:11px">{j.get('stars','⭐')}</span>
            <b> <a href="{j['url']}" style="color:#1A3A5C;text-decoration:none">{j['title']}</a></b><br>
            <span style="color:#888;font-size:11px">{j.get('ort','?')} &middot; {dl} &middot; pub.{j.get('pub_date','?')}</span>
          </td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('korkort','–')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('tider','–')}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;font-size:12px">{j.get('deadline','Löpande')}</td>
        </tr>"""

    new_section = f"""
    <table style="width:100%;border-collapse:collapse;background:white;border:1px solid #e5e7eb;border-radius:6px">
      <tr style="background:#1B4332;color:white;font-size:11px">
        <th style="padding:8px 10px;text-align:left">Annons</th>
        <th style="padding:8px">Körkort</th><th style="padding:8px">Helg/kväll</th><th style="padding:8px">Deadline</th>
      </tr>{rows}</table>""" if new_jobs else "<p style='color:#888;font-size:13px'>Inga nya jobb idag.</p>"

    html=f"""<html><body style="font-family:Arial,sans-serif;max-width:700px;margin:auto">
    <div style="background:#1B4332;color:white;padding:18px 20px;border-radius:8px 8px 0 0">
      <h2 style="margin:0;font-size:18px">📋 PA-jobb Uppsala – {TODAY}</h2>
      <p style="margin:4px 0 0;opacity:.75;font-size:12px">Daglig uppdatering &middot; Manliga kunder &middot; Max {MAX_KM} km från Gunsta</p>
    </div>
    <div style="padding:16px 20px;background:#f9f9f9;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px">

      <!-- Stats row -->
      <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap">
        <div style="background:#D8F3DC;border-radius:8px;padding:8px 14px;text-align:center;min-width:80px">
          <div style="font-size:22px;font-weight:700;color:#1B4332">{len(new_jobs)}</div>
          <div style="font-size:11px;color:#1B4332">Nya idag</div>
        </div>
        <div style="background:#DBEAFE;border-radius:8px;padding:8px 14px;text-align:center;min-width:80px">
          <div style="font-size:22px;font-weight:700;color:#1A3A5C">{len(open_jobs)}</div>
          <div style="font-size:11px;color:#1A3A5C">Fortfarande öppna</div>
        </div>
        <div style="background:#FEE2E2;border-radius:8px;padding:8px 14px;text-align:center;min-width:80px">
          <div style="font-size:22px;font-weight:700;color:#991B1B">{len(closed_jobs)}</div>
          <div style="font-size:11px;color:#991B1B">Stängda</div>
        </div>
        <div style="flex:1;background:#1B4332;border-radius:8px;padding:10px 16px;display:flex;align-items:center;justify-content:center">
          <a href="{PAGES_URL}" style="color:white;font-weight:600;font-size:14px;text-decoration:none">
            🔗 Visa fullständig dashboard →
          </a>
        </div>
      </div>

      <!-- New jobs table -->
      <h3 style="color:#1B4332;margin:0 0 10px">🆕 Nya jobb idag</h3>
      {new_section}

      <p style="margin-top:16px;font-size:12px;color:#888;text-align:center">
        Alla {len(open_jobs)} öppna jobb, detaljer och historik →
        <a href="{PAGES_URL}" style="color:#1A3A5C">{PAGES_URL}</a>
      </p>
      <p style="margin-top:6px;font-size:11px;color:#aaa;text-align:center">
        Excel bifogad &middot; PA-agent v4 &middot; GitHub Actions
      </p>
    </div></body></html>"""

    subj=(f"PA-jobb Uppsala {TODAY} – {len(new_jobs)} nya"
          +(f", {len(open_jobs)} öppna" if open_jobs else "")
          +(f", {len(closed_jobs)} stängda" if closed_jobs else ""))
    msg=MIMEMultipart("alternative"); msg["Subject"]=subj; msg["From"]=sender; msg["To"]=rcv
    msg.attach(MIMEText(html,"html"))
    if EXCEL_FILE.exists():
        with open(EXCEL_FILE,"rb") as f:
            p=MIMEBase("application","octet-stream"); p.set_payload(f.read())
        encoders.encode_base64(p); p.add_header("Content-Disposition",f"attachment; filename={EXCEL_FILE.name}")
        msg.attach(p)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com",465) as s:
            s.login(sender,pw); s.sendmail(sender,rcv,msg.as_string())
        print(f"[INFO] Email sent to {rcv}")
    except Exception as e:
        print(f"[ERROR] Email failed: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print(f"\n{'='*55}\nPA-agent v4 · {datetime.now().strftime('%Y-%m-%d %H:%M')}\n{'='*55}\n")

    seen      = load_seen()
    raw       = fetch()
    filtered, dist_excl = filter_jobs(raw)

    current_ids = {j["id"]:j for j in filtered}
    new_jobs    = []   # seen for the first time today
    open_jobs   = []   # already in seen, still active
    closed_jobs = []   # were in seen, no longer on site

    for j in filtered:
        aid = j["id"]
        enrich(j, seen)
        if aid not in seen:
            j["first_seen"] = TODAY
            new_jobs.append(j)
            print(f"[NEW  {j['stars']}] {j['title'][:50]} | {j['ort']} | {j['korkort']} | {j['tider']}")
        else:
            open_jobs.append(j)

    for aid, meta in seen.items():
        if aid not in current_ids:
            closed_jobs.append({"id":aid,"title":meta.get("title","?"),"url":"#"})

    # Sort
    for lst in [new_jobs, open_jobs]:
        lst.sort(key=lambda j:j.get("stars","⭐"),reverse=True)

    # Update memory
    seen.update({j["id"]:{"title":j["title"],"first_seen":j["first_seen"]} for j in new_jobs})
    for j in open_jobs:
        if j["id"] in seen:
            seen[j["id"]]["title"] = j["title"]
    for j in closed_jobs:
        seen.pop(j["id"],None)
    save_seen(seen)

    # Outputs
    build_dashboard(new_jobs, open_jobs, closed_jobs, dist_excl)
    if new_jobs or closed_jobs:
        update_excel(new_jobs, [j["id"] for j in closed_jobs])
    else:
        print("[INFO] No changes – Excel unchanged.")
    send_email(new_jobs, open_jobs, closed_jobs, dist_excl)

    print(f"\n✅ Done! New={len(new_jobs)} Open={len(open_jobs)} Closed={len(closed_jobs)} Far={len(dist_excl)}")

if __name__=="__main__":
    main()
