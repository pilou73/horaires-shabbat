"""
Identification de la lecture du Chabbat (Israël), calculée hors-ligne avec la bibliothèque hdate.

- Chabbat ordinaire : la parasha de la semaine (y compris parashiot jumelées).
- Chabbat qui tombe un jour de fête ou pendant Hol HaMoed : il n'y a pas de parasha
  régulière -> on renvoie le nom de la fête (ex. « שבת חול המועד סוכות »).

Si hdate n'est pas disponible, on garde le nom fourni par les données (Excel / liste intégrée).
"""
import datetime as dt

# fête (nom hdate en hébreu) -> (titre affiché, clé des illustrations dans ai_images.THEMES)
HOLIDAY_TITLES = {
    "א' ראש השנה": ("שבת ראש השנה", "ראש השנה"),
    "ראש השנה": ("שבת ראש השנה", "ראש השנה"),
    "יום הכפורים": ("שבת יום הכיפורים", "יום הכיפורים"),
    "סוכות": ("שבת חג הסוכות", "סוכות"),
    "חול המועד סוכות": ("שבת חול המועד סוכות", "חול המועד סוכות"),
    "שמיני עצרת": ("שמיני עצרת · שמחת תורה", "וזאת הברכה"),
    "שמחת תורה": ("שמיני עצרת · שמחת תורה", "וזאת הברכה"),
    "פסח": ("שבת חג הפסח", "פסח"),
    "חול המועד פסח": ("שבת חול המועד פסח", "חול המועד פסח"),
    "שביעי פסח": ("שביעי של פסח", "שביעי של פסח"),
    "שבועות": ("שבת חג השבועות", "שבועות"),
}


def saturday_of(day):
    """Samedi de la semaine d'un vendredi (date stockée dans les données) ou du samedi lui-même."""
    if isinstance(day, dt.datetime):
        day = day.date()
    return day + dt.timedelta(days=(5 - day.weekday()) % 7)


def shabbat_reading(day, fallback_parasha=""):
    """
    Retourne un dict :
      title      : texte du titre de l'affiche (« פרשת נח », « שבת חול המועד סוכות »…)
      theme      : clé pour les 3 illustrations (nom de parasha ou de fête)
      parasha    : nom de la parasha régulière ("" si Chabbat de fête)
      is_holiday : True si Chabbat de fête / Hol HaMoed (lecture spéciale)
      source     : "hdate" ou "data"
    """
    sat = saturday_of(day)
    try:
        from hdate import HDateInfo
        from hdate.translator import set_language
        set_language("he")
        info = HDateInfo(sat, diaspora=False)  # Israël
        parasha = str(info.parasha)
        parasha = "" if parasha.lower() == "none" else parasha
        holidays = [str(h) for h in info.holidays]
    except Exception as e:  # hdate absent ou erreur : on garde les données telles quelles
        print(f"⚠️ hdate indisponible ({e}) — nom de parasha pris dans les données")
        name = str(fallback_parasha or "").replace("=", "-").strip()
        return {"title": f"פרשת {name}" if name else "שבת שלום", "theme": name,
                "parasha": name, "is_holiday": False, "source": "data"}

    for h in holidays:
        if h in HOLIDAY_TITLES and (not parasha or h in ("שמיני עצרת", "שמחת תורה")):
            title, theme = HOLIDAY_TITLES[h]
            return {"title": title, "theme": theme, "parasha": parasha, "is_holiday": True, "source": "hdate"}
    if not parasha:  # autre fête sans parasha (cas imprévu) : nom de la fête
        name = holidays[0] if holidays else ""
        return {"title": f"שבת {name}".strip() if name else "שבת שלום", "theme": name,
                "parasha": "", "is_holiday": True, "source": "hdate"}
    return {"title": f"פרשת {parasha}", "theme": parasha, "parasha": parasha,
            "is_holiday": False, "source": "hdate"}
