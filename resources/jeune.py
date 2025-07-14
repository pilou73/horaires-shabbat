import requests
from datetime import datetime, timedelta
from astral import LocationInfo
from astral.sun import sun

CITY = "Ramat Gan"
REGION = "Israel"
TZ = "Asia/Jerusalem"
LAT = 32.0680
LON = 34.8248
GEONAMEID = "293397"
ALOT_OFFSET = 72
TZEIT_OFFSET = 18

def fetch_fast_days(start, end, geonameid=GEONAMEID):
    url = "https://www.hebcal.com/hebcal"
    params = {
        "v": 1,
        "cfg": "json",
        "maj": "on",
        "ss": "on",
        "mf": "on",
        "start": start,
        "end": end,
        "geonameid": geonameid,
        "lg": "he"
    }
    r = requests.get(url, params=params)
    items = r.json()["items"]

    fasts = []
    for item in items:
        if item.get("subcat") == "fast" and item.get("category") == "holiday":
            fast_name = item.get("title_orig", item.get("title", ""))
            fast_name_he = item.get("hebrew", "")
            fast_date = item["date"][:10]
            # Exclure Kippour et 9 Av
            if (
                "Yom Kippur" in fast_name or "יום כיפור" in fast_name_he or
                "Tisha B'Av" in fast_name or "תשעה באב" in fast_name_he or
                "Fast of Av" in fast_name or "Fast of Ninth of Av" in fast_name
            ):
                continue
            fasts.append({
                "date": fast_date,
                "nom": fast_name,
                "nom_heb": fast_name_he,
            })
    return fasts

def get_fast_times(g_date, city=CITY, region=REGION, tz=TZ, latitude=LAT, longitude=LON, alot_offset=ALOT_OFFSET, tzeit_offset=TZEIT_OFFSET):
    loc = LocationInfo(city, region, tz, latitude, longitude)
    s = sun(loc.observer, date=g_date, tzinfo=loc.timezone)
    sunrise = s["sunrise"]
    sunset = s["sunset"]
    alot_dt = sunrise - timedelta(minutes=alot_offset)
    tzeit_dt = sunset + timedelta(minutes=tzeit_offset)
    return alot_dt, tzeit_dt

def format_dt(dt):
    return dt.strftime('%Y%m%dT%H%M%S')

if __name__ == "__main__":
    start_date = "2024-09-01"
    end_date = "2030-09-30"
    fasts = fetch_fast_days(start_date, end_date)
    fasts.sort(key=lambda f: f["date"])

    # Affichage console et génération ICS
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Jeûnes juifs diurnes (Ramat Gan)",
        "X-WR-TIMEZONE:Asia/Jerusalem"
    ]

    for f in fasts:
        g_date = datetime.strptime(f['date'], "%Y-%m-%d").date()
        debut_dt, fin_dt = get_fast_times(g_date)
        summary = f"{f['nom']} ({f['nom_heb']})"
        description = "Jeûne diurne. Début à עלות השחר, fin à צאת הכוכבים (+18min)."

        # Affichage console
        print(f"{f['date']} | {f['nom']} ({f['nom_heb']})")
        print(f"  Début du jeûne (עלות השחר): {debut_dt.strftime('%H:%M')}")
        print(f"  Fin du jeûne   (צאת הכוכבים + 18 min): {fin_dt.strftime('%H:%M')}")
        print("----------------------------------------")

        # ICS
        ics_lines.append("BEGIN:VEVENT")
        ics_lines.append(f"DTSTART;TZID=Asia/Jerusalem:{format_dt(debut_dt)}")
        ics_lines.append(f"DTEND;TZID=Asia/Jerusalem:{format_dt(fin_dt)}")
        ics_lines.append(f"SUMMARY:{summary}")
        ics_lines.append(f"DESCRIPTION:{description}")
        ics_lines.append("END:VEVENT")

    ics_lines.append("END:VCALENDAR")

    with open("jeunes.ics", "w", encoding="utf-8") as f_ics:
        f_ics.write('\n'.join(ics_lines))

    print("Fichier jeunes.ics généré (sans 9 Av ni Kippour).")