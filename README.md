# 📋 PA-jobb Uppsala – Daglig agent

Automatisk daglig uppdatering av PA-jobb för manliga kunder i Uppsala-regionen.
Körs gratis via GitHub Actions kl 07:00 varje morgon.

---

## 🗂️ Filstruktur

```
pa-agent/
├── .github/
│   └── workflows/
│       └── daily_update.yml    ← GitHub Actions schema
├── scripts/
│   ├── update_jobs.py          ← Huvudskript (söker, filtrerar, uppdaterar)
│   └── seen_jobs.json          ← Skapas automatiskt (håller koll på sedda jobb)
├── PA_Jobb_Uppsala_Gunsta.xlsx ← Din Excel-fil (lägg in här!)
├── requirements.txt
└── README.md
```

---

## 🚀 Setup – steg för steg

### Steg 1 – Skapa ett GitHub-konto
Gå till **github.com** och registrera dig (gratis).

---

### Steg 2 – Skapa ett nytt repo

1. Klicka på **"New repository"** (gröna knappen)
2. Namn: `pa-jobb-uppsala`
3. Välj **Private** (så ingen annan ser dina jobb)
4. Klicka **"Create repository"**

---

### Steg 3 – Ladda upp filerna

Du kan antingen använda GitHub-webbgränssnittet eller Git.

**Enklaste sättet (webbläsaren):**

1. I ditt nya repo, klicka **"uploading an existing file"**
2. Dra och släpp **alla filer** från denna mapp:
   - `PA_Jobb_Uppsala_Gunsta.xlsx`
   - `requirements.txt`
   - `README.md`
3. Skapa sedan mapparna manuellt:
   - Klicka **"Add file" → "Create new file"**
   - Skriv `scripts/update_jobs.py` som filnamn
   - Klistra in innehållet från `scripts/update_jobs.py`
   - Upprepa för `.github/workflows/daily_update.yml`

---

### Steg 4 – Konfigurera Gmail för mejlutskick

Agenten skickar ett dagligt mejldigest till dig. Du behöver ett **App Password** från Gmail.

1. Gå till **myaccount.google.com**
2. Klicka **"Säkerhet"** → **"Tvåstegsverifiering"** (aktivera om det inte är på)
3. Gå till **"App-lösenord"** (söka efter det)
4. Välj app: **"E-post"**, enhet: **"Annan"** → skriv `PA-agent`
5. Klicka **"Generera"** → kopiera det 16-siffriga lösenordet

---

### Steg 5 – Lägg in hemligheter i GitHub

1. I ditt repo: **Settings → Secrets and variables → Actions**
2. Klicka **"New repository secret"** tre gånger:

| Namn | Värde |
|------|-------|
| `EMAIL_FROM` | Din Gmail-adress, t.ex. `boris@gmail.com` |
| `EMAIL_PASSWORD` | Det 16-siffriga App Password från steg 4 |
| `EMAIL_TO` | Din e-post dit mejlet ska skickas (kan vara samma) |

---

### Steg 6 – Testa att det fungerar

1. Gå till **Actions** (fliken längst upp i repot)
2. Klicka på **"PA-jobb Uppsala – Daglig uppdatering"**
3. Klicka **"Run workflow"** → **"Run workflow"**
4. Vänta ~1 minut → grön bock = det fungerar!
5. Kolla din e-post – du borde ha fått ett mejl med sammanfattning

---

### Steg 7 – Automatisk körning

Agenten kör nu automatiskt varje dag kl **07:00** (svensk tid).

Du behöver inte göra något mer. Varje morgon:
- ✅ Söker efter nya manliga kund-jobb i Uppsala
- ✅ Uppdaterar Excel-filen i repot
- ✅ Skickar mejl med sammanfattning
- ✅ Tar bort utgångna jobb automatiskt

---

## 📥 Hämta uppdaterad Excel

Den uppdaterade Excel-filen sparas direkt i repot varje dag.

För att ladda ner den:
1. Gå till repot på GitHub
2. Klicka på `PA_Jobb_Uppsala_Gunsta.xlsx`
3. Klicka **"Download raw file"**

Eller: filen bifogas automatiskt i det dagliga mejlet.

---

## ⚙️ Ändra körtid

I filen `.github/workflows/daily_update.yml`:

```yaml
- cron: "0 5 * * *"   # 07:00 svensk vintertid (05:00 UTC)
```

Byt `0 5` till t.ex. `0 4` för 06:00 eller `30 5` för 07:30.

---

## 🛑 Pausa agenten

Gå till **Actions → "PA-jobb Uppsala"** → **"..."** → **"Disable workflow"**.

---

## 💡 Tips

- Agenten hittar **nya** jobb baserat på Platsbanken-ID – den vet vad den redan sett
- `scripts/seen_jobs.json` är minnet – ta inte bort den
- Vill du lägga till fler källor? Öppna ett ärende eller ändra i `update_jobs.py`
