from datetime import datetime, date, timedelta
from zmanim.hebrew_calendar.jewish_calendar import JewishCalendar

HEBREW_MONTHS = {
    1: 'Nissan',  2: 'Iyar',     3: 'Sivan',    4: 'Tamouz',
    5: 'Av',      6: 'Eloul',    7: 'Tishrei',  8: 'Heshvan',
    9: 'Kislev', 10: 'Tevet',   11: 'Shevat',  12: 'Adar',
   13: 'Adar II'  # pour années embolismiques
}

def get_jewish_month_name(jm, jy):
    if jm == 12 and JewishCalendar.is_jewish_leap_year(jy):
        return 'Adar I'
    if jm == 13:
        return 'Adar II'
    return HEBREW_MONTHS.get(jm, 'Mois-inconnu')

def calculate_molad_for_date(gregorian_date):
    """
    Calcule le molad pour LA date du Rosh Chodesh fournie en entrée.
    """
    jc = JewishCalendar(gregorian_date)
    molad_obj = jc.molad()  # méthode qui renvoie l'objet Molad

    # extraction des parts du molad
    hour     = molad_obj.molad_hours
    minute   = molad_obj.molad_minutes
    chalakim = molad_obj.molad_chalakim

    jm = jc.jewish_month
    jy = jc.jewish_year
    month_name = get_jewish_month_name(jm, jy)

    return {
        "rosh_chodesh_date": gregorian_date,
        "molad": f"{hour}:{minute}:+ {chalakim} chalakim",
        "hebrew_month": month_name,
        "hebrew_year": jy
    }

def find_next_rosh_chodesh(start_date=None):
    """
    À partir de start_date (ou d'aujourd'hui), cherche le prochain date
    où jewish_day == 1 (Rosh Chodesh) et renvoie cette date.
    """
    current = start_date or date.today()
    # on cherchera au plus sur 60 jours pour couvrir tous les cas
    for _ in range(60):
        jc = JewishCalendar(datetime.combine(current, datetime.min.time()))
        if jc.jewish_day == 1:
            return current
        current += timedelta(days=1)
    raise RuntimeError("No Rosh Chodesh in 60 next days.")

if __name__ == "__main__":
    import sys

    # Lecture d'une date passée en argument au format YYYY-MM-DD
    if len(sys.argv) >= 2:
        try:
            start = date.fromisoformat(sys.argv[1])
        except ValueError:
            print("❌ Invalible Format. Use YYYY-MM-DD.")
            sys.exit(1)
    else:
        start = None

    # Recherche et calcul
    rc_date = find_next_rosh_chodesh(start)
    result = calculate_molad_for_date(rc_date)

    # Affichage
    print("🔍 Search from :", start or date.today())
    print("📅 Rosh Chodesh found :", result["rosh_chodesh_date"])
    print("🌙 Hebrew Month   :", result["hebrew_month"])
    print("🕰️  Molad           :", result["molad"])
    print("🗓️  Hebrew Year:", result["hebrew_year"])
	