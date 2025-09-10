import json
import os
import sys
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageFont
import requests
import re

# Ajout des imports pour le calcul du molad
try:
    from zmanim.hebrew_calendar.jewish_calendar import JewishCalendar
    ZMANIM_AVAILABLE = True
except ImportError:
    ZMANIM_AVAILABLE = False
    print("⚠️ Bibliothèque zmanim non disponible. Le molad ne sera pas calculé.")

def get_font(size=24, bold=False):
    """
    Charge les polices FreeSans depuis le dossier resources.
    Supporte hébreu et latin.
    """
    try:
        if bold:
            return ImageFont.truetype("resources/FreeSansBold.ttf", size)
        else:
            return ImageFont.truetype("resources/FreeSans.ttf", size)
    except OSError as e:
        print(f"⚠️ Erreur chargement police FreeSans: {e}")
        print("Utilisation de la police par défaut.")
        return ImageFont.load_default()

def is_hebrew(text):
    """Retourne True si le texte contient des caractères hébreux"""
    return any('\u0590' <= c <= '\u05FF' for c in text)

def reverse_hebrew_text(text):
    pattern = r'(\d{1,2}/\d{1,2}/\d{4}|\d{1,2}:\d{2}|\d+|[\u0590-\u05FF]+|[^\w\s]+|\s+)'
    tokens = re.findall(pattern, text)
    tokens = tokens[::-1]
    result = []
    for token in tokens:
        if re.match(r'^[\u0590-\u05FF]+$', token):
            result.append(token[::-1])
        else:
            result.append(token)
    return ''.join(result).strip()

def round_to_nearest_5(dt):
    if dt is None:
        return None
    minute = dt.minute
    rounded_minute = 5 * round(minute / 5)
    if rounded_minute == 60:
        dt = dt.replace(minute=0) + timedelta(hours=1)
    else:
        dt = dt.replace(minute=rounded_minute)
    return dt

def get_next_friday_saturday_tuesday(today=None):
    if today is None:
        today = datetime.now()
    days_to_friday = (4 - today.weekday()) % 7
    friday = today + timedelta(days=days_to_friday)
    saturday = friday + timedelta(days=1)
    days_to_tuesday = (1 - today.weekday()) % 7
    tuesday = today + timedelta(days=days_to_tuesday)
    return friday, saturday, tuesday

def get_zmanim_for_day(lat, lon, date):
    url = (
        f"https://www.hebcal.com/zmanim?cfg=json"
        f"&latitude={lat}&longitude={lon}"
        f"&date={date.strftime('%Y-%m-%d')}"
    )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        return data.get('times', {})
    except Exception as e:
        print(f"Erreur lors de la récupération des zmanim pour {date.strftime('%A %Y-%m-%d')}: {e}")
        return {}

def get_all_zmanim(lat, lon, today=None):
    friday, saturday, tuesday = get_next_friday_saturday_tuesday(today)
    zmanim = {}

    times_fri = get_zmanim_for_day(lat, lon, friday)
    if 'candle_lighting' in times_fri:
        zmanim['candle_lighting'] = datetime.fromisoformat(times_fri['candle_lighting'])

    times_sat = get_zmanim_for_day(lat, lon, saturday)
    if 'tzeit85deg' in times_sat:
        zmanim['fin_shabbat'] = datetime.fromisoformat(times_sat['tzeit85deg'])

    times_tue = get_zmanim_for_day(lat, lon, tuesday)
    if 'sunset' in times_tue:
        zmanim['shkiya'] = datetime.fromisoformat(times_tue['sunset'])

    if 'candle_lighting' not in zmanim and 'shkiya' in zmanim:
        zmanim['candle_lighting'] = zmanim['shkiya'] - timedelta(minutes=18)

    return zmanim

def get_reference_time(reference, zmanim):
    ref_map = {
        'candle_lighting': 'candle_lighting',
        'fin_shabbat': 'fin_shabbat',
        'shkiya': 'shkiya'
    }
    ref_key = ref_map.get(reference, reference)
    return zmanim.get(ref_key)

def calculer_horaire_struct(activite, zmanim):
    if activite['type'] == 'fixe':
        try:
            heure_str = activite['heure']
            h, m = map(int, heure_str.split(':'))
            dt = datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)
            return round_to_nearest_5(dt)
        except Exception:
            return None
    elif activite['type'] == 'calculee':
        try:
            minutes_offset = int(activite['minutes_offset'])
            avant_apres = activite['avant_apres']
            reference = activite['reference']
            heure_reference = get_reference_time(reference, zmanim)
            if heure_reference is None:
                return None
            if avant_apres == 'avant':
                dt = heure_reference - timedelta(minutes=abs(minutes_offset))
            else:
                dt = heure_reference + timedelta(minutes=abs(minutes_offset))
            return round_to_nearest_5(dt)
        except Exception:
            return None
    else:
        return None

def get_parasha_name_hebrew(lat, lon, shabbat_date):
    week_start = shabbat_date - timedelta(days=6)
    week_end = shabbat_date + timedelta(days=2)
    url = (
        f"https://www.hebcal.com/hebcal?cfg=json"
        f"&v=1"
        f"&maj=on&min=on&mod=on&s=on"
        f"&start={week_start.strftime('%Y-%m-%d')}"
        f"&end={week_end.strftime('%Y-%m-%d')}"
        f"&geo=pos&latitude={lat}&longitude={lon}"
        f"&lg=he"
    )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        for item in data.get('items', []):
            if item.get('category') == 'parashat' and item.get('date', '').startswith(shabbat_date.strftime('%Y-%m-%d')):
                return item.get('hebrew', item.get('title'))
    except Exception as e:
        print(f"Erreur lors de la récupération de la parasha: {e}")
    return ""

def is_shabbat_mevarchim(lat, lon, shabbat_date):
    sunday = shabbat_date + timedelta(days=1)
    next_saturday = shabbat_date + timedelta(days=7)
    url = (
        f"https://www.hebcal.com/hebcal?cfg=json"
        f"&v=1"
        f"&maj=on&min=on&mod=on&nx=on"
        f"&start={sunday.strftime('%Y-%m-%d')}"
        f"&end={next_saturday.strftime('%Y-%m-%d')}"
        f"&geo=pos&latitude={lat}&longitude={lon}"
        f"&lg=he"
    )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        for item in data.get('items', []):
            if item.get('category') == 'roshchodesh':
                return True
    except Exception as e:
        print(f"Erreur lors de la vérification de Rosh Hodesh: {e}")
    return False

HEBREW_DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת']
HEBREW_MONTHS = {
    1: 'ניסן', 2: 'אייר', 3: 'סיון', 4: 'תמוז',
    5: 'אב', 6: 'אלול', 7: 'תשרי', 8: 'חשוון',
    9: 'כסלו', 10: 'טבת', 11: 'שבט', 12: 'אדר',
    13: 'אדר ב׳'
}

def get_weekday_name_hebrew(dt):
    return HEBREW_DAYS[(dt.weekday() + 1) % 7]

def get_jewish_month_name_hebrew(jm, jy):
    if not ZMANIM_AVAILABLE:
        return ""
    if jm == 12 and JewishCalendar.is_jewish_leap_year(jy):
        return 'אדר א׳'
    if jm == 13:
        return 'אדר ב׳'
    return HEBREW_MONTHS.get(jm, 'חודש לא ידוע')

def get_next_month_molad(shabbat_date):
    if not ZMANIM_AVAILABLE:
        return ""
    try:
        jc = JewishCalendar(datetime.combine(shabbat_date, datetime.min.time()))
        current_jyear = jc.jewish_year
        current_jmonth = jc.jewish_month

        if current_jmonth == 13:
            next_jmonth = 1
            next_jyear = current_jyear + 1
        else:
            next_jmonth = current_jmonth + 1
            next_jyear = current_jyear

        jc_next_month = JewishCalendar()
        jc_next_month.set_jewish_date(next_jyear, next_jmonth, 1)
        molad_obj = jc_next_month.molad()

        molad_date = jc_next_month.gregorian_date - timedelta(days=1)
        hour = molad_obj.molad_hours
        minute = molad_obj.molad_minutes
        chalakim = molad_obj.molad_chalakim
        weekday_he = get_weekday_name_hebrew(molad_date)

        return f"המולד ביום {weekday_he} בשעה {hour:02d}:{minute:02d} {chalakim} חלקים"
    except Exception as e:
        print(f"Erreur lors du calcul du molad: {e}")
        return ""

def get_rosh_hodesh_days_for_next_month(shabbat_date):
    if not ZMANIM_AVAILABLE:
        return []
    try:
        jc = JewishCalendar(datetime.combine(shabbat_date, datetime.min.time()))
        current_jyear = jc.jewish_year
        current_jmonth = jc.jewish_month

        if current_jmonth == 13:
            next_jmonth = 1
            next_jyear = current_jyear + 1
        else:
            next_jmonth = current_jmonth + 1
            next_jyear = current_jyear

        jc_current = JewishCalendar()
        jc_current.set_jewish_date(current_jyear, current_jmonth, 1)

        if hasattr(jc_current, "days_in_jewish_month"):
            days_in_current_month = jc_current.days_in_jewish_month()
        elif hasattr(JewishCalendar, "getLastDayOfJewishMonth"):
            days_in_current_month = JewishCalendar.getLastDayOfJewishMonth(current_jmonth, current_jyear)
        else:
            days_in_current_month = 29  # fallback

        jc_next = JewishCalendar()
        jc_next.set_jewish_date(next_jyear, next_jmonth, 1)
        rc1_gdate = jc_next.gregorian_date

        month_name = get_jewish_month_name_hebrew(next_jmonth, next_jyear)

        if days_in_current_month == 30:
            jc_current.set_jewish_date(current_jyear, current_jmonth, 30)
            rc0_gdate = jc_current.gregorian_date
            day0_he = get_weekday_name_hebrew(rc0_gdate)
            day1_he = get_weekday_name_hebrew(rc1_gdate)
            return f"ראש חודש {month_name} יהיה ביום {day0_he} ו{day1_he}"
        else:
            day1_he = get_weekday_name_hebrew(rc1_gdate)
            return f"ראש חודש {month_name} יהיה ביום {day1_he}"
    except Exception as e:
        print(f"Erreur lors du calcul des jours de Rosh Hodesh: {e}")
        return ""

def get_birkat_halevana_announcement(saturday, is_mevarchim):
    if not ZMANIM_AVAILABLE or is_mevarchim:
        return ""
    try:
        shabbat_date = saturday.replace(hour=0, minute=0, second=0, microsecond=0)
        jc = JewishCalendar(shabbat_date)
        jyear = jc.jewish_year
        jmonth = jc.jewish_month

        jc_8 = JewishCalendar()
        jc_8.set_jewish_date(jyear, jmonth, 8)
        gdate_8 = jc_8.gregorian_date

        jc_15 = JewishCalendar()
        jc_15.set_jewish_date(jyear, jmonth, 15)
        gdate_15 = jc_15.gregorian_date

        today = shabbat_date.date()
        start = gdate_8
        end = gdate_15

        if today < start:
            return f"ברכת הלבנה ניתן לומר מ־{start.strftime('%d/%m/%Y')} עד {end.strftime('%d/%m/%Y')}"
        elif start <= today <= end:
            return f"ברכת הלבנה ניתן לומר עד {end.strftime('%d/%m/%Y')}"
        elif today > end:
            return "הזמן לברכת הלבנה עבר"
        else:
            return ""
    except Exception as e:
        print(f"Erreur Birkat HaLevana: {e}")
        return ""

def get_birkat_halevana_icon(birkat_text):
    """
    Retourne le chemin de l'icône à utiliser pour Birkat HaLevana selon le texte d'annonce.
    """
    if not birkat_text:
        return None
    if "מ־" in birkat_text:  # "de ... à ..." → début de la période
        return "resources/first_moon.png"
    if "עד" in birkat_text:  # "jusqu'à ..." → période en cours
        return "resources/full_moon.png"
    return None

def get_next_tekufa_announcement(shabbat_date, ics_path="tekufa_2025_2035.ics"):
    """
    Cherche une tekoufa qui tombe dans la semaine qui suit le shabbat donné.
    Traduit automatiquement le nom en hébreu et retourne (titre_hébreu, date, heure).
    """
    if not os.path.exists(ics_path):
        return None

    # Dictionnaire de traduction des tekoufot
    tekufa_translations = {
        'tevet': 'תקופת טבת',
        'nisan': 'תקופת ניסן',
        'nissan': 'תקופת ניסן',
        'tammuz': 'תקופת תמוז',
        'tamuz': 'תקופת תמוז',
        'tishri': 'תקופת תשרי',
        'tishrei': 'תקופת תשרי'
    }

    start = (shabbat_date + timedelta(days=1)).date()
    end = (shabbat_date + timedelta(days=7)).date()
    try:
        with open(ics_path, encoding="utf-8") as f:
            content = f.read()
        events = re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", content, re.DOTALL)
        for event in events:
            # Cherche DTSTART (date seule ou date+heure)
            m = re.search(r"DTSTART(?:;VALUE=DATE)?(?::|;TZID=[^:]+:)(\d{8})(T(\d{2})(\d{2}))?", event)
            if not m:
                continue
            event_date = datetime.strptime(m.group(1), "%Y%m%d").date()
            event_hour = None
            if m.group(2):  # Il y a une heure
                hour = int(m.group(3))
                minute = int(m.group(4))
                event_hour = f"{hour:02d}:{minute:02d}"
            m2 = re.search(r"SUMMARY:(.+)", event)
            if not m2:
                continue
            summary = m2.group(1).strip()

            # Traduction automatique en hébreu
            summary_lower = summary.lower()
            hebrew_title = summary  # Par défaut, garde le titre original
            for eng_key, heb_value in tekufa_translations.items():
                if eng_key in summary_lower:
                    hebrew_title = heb_value
                    break

            print(f"DEBUG TEKUFA: {event_date} - {summary} -> {hebrew_title} - {event_hour}")
            if start <= event_date <= end:
                return hebrew_title, event_date, event_hour
        return None
    except Exception as e:
        print(f"Erreur parsing tekufa: {e}")
        return None

def generer_image(config, horaires_activites, horaires_semaine, parasha_he, template_file, is_mevarchim, molad_text=None, rosh_hodesh_text=None, birkat_halevana_text=None, tekufa_announcement=None):
    if not os.path.exists('output'):
        os.makedirs('output')
    img = Image.open(template_file)
    draw = ImageDraw.Draw(img)

    # Chargement des polices FreeSans (hébreu + latin)
    font_hebrew = get_font(24, bold=False)    # Texte normal
    font_hebrew_bold = get_font(48, bold=True)    # Parasha (gros et gras)
    font_hebrew_title = get_font(32, bold=True)    # Titres (gras)
    font_hebrew_green = get_font(24, bold=False)    # Texte vert (semaine)

    # 1. Parasha dans le carré bleu (centré)
    if parasha_he:
        parasha_text = reverse_hebrew_text(parasha_he)
        bbox = draw.textbbox((0, 0), parasha_text, font=font_hebrew_bold)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        x = 70 + (400 - w) // 2
        y = 220 + (80 - h) // 2
        draw.text((x, y), parasha_text, font=font_hebrew_bold, fill=(0, 51, 204))
        if is_mevarchim:
            mevarchim_text = reverse_hebrew_text("שבת מברכין")
            bbox2 = draw.textbbox((0, 0), mevarchim_text, font=font_hebrew_title)
            w2 = bbox2[2] - bbox2[0]
            h2 = bbox2[3] - bbox2[1]
            x2 = 70 + (400 - w2) // 2
            y2 = y + h + 10
            draw.text((x2, y2), mevarchim_text, font=font_hebrew_title, fill=(0, 51, 204))

    # 2. Titre "זמני השבת" avant les horaires de Chabbat (centré, bleu)
    titre_shabbat = reverse_hebrew_text("זמני השבת")
    y_titre_shabbat = 355
    largeur_img = img.width
    bbox_titre_shabbat = draw.textbbox((0, 0), titre_shabbat, font=font_hebrew_title)
    w_titre_shabbat = bbox_titre_shabbat[2] - bbox_titre_shabbat[0]
    x_titre_shabbat = (largeur_img - w_titre_shabbat) // 2
    draw.text((x_titre_shabbat, y_titre_shabbat), titre_shabbat, font=font_hebrew_title, fill=(0,0,0))

    # 3. Horaires des activités de Chabbat (zone blanche centrale)
    y_start = 400
    line_height = 38
    x_nom_heb = 820    # Position pour noms hébreux (à droite)
    x_nom_lat = 170    # Position pour noms latins (à gauche)
    x_horaire_heb = 170  # Position horaires quand nom hébreu (à gauche)
    x_horaire_lat = 820  # Position horaires quand nom latin (à droite)

    activites_valides = [act for act in horaires_activites if act['horaire'] is not None]
    for idx, act in enumerate(activites_valides):
        nom_original = act['nom']
        horaire = act['horaire'].strftime('%H:%M')
        y = y_start + idx * line_height
        
        if is_hebrew(nom_original):
            # Texte hébreu : inverse + aligne à droite
            nom_affiche = reverse_hebrew_text(nom_original)
            draw.text((x_nom_heb, y), nom_affiche, font=font_hebrew, fill=(0,0,0), anchor="ra")
            draw.text((x_horaire_heb, y), horaire, font=font_hebrew, fill=(0,0,0), anchor="la")
        else:
            # Texte latin : normal + aligne à gauche
            draw.text((x_nom_lat, y), nom_original, font=font_hebrew, fill=(0,0,0), anchor="la")
            draw.text((x_horaire_lat, y), horaire, font=font_hebrew, fill=(0,0,0), anchor="ra")

    # 4. Titre avant les horaires de semaine (en vert, gras, centré)
    titre_semaine = reverse_hebrew_text("זמני אמצע השבוע")
    y_titre = y_start + line_height * (len(activites_valides) + 1)
    bbox_titre = draw.textbbox((0, 0), titre_semaine, font=font_hebrew_title)
    w_titre = bbox_titre[2] - bbox_titre[0]
    x_titre = (largeur_img - w_titre) // 2
    draw.text((x_titre, y_titre), titre_semaine, font=font_hebrew_title, fill=(0, 128, 0))

    # 5. Horaires de semaine (en vert, bien plus bas) - CORRECTION ICI
    x_nom_semaine_heb = 800   # Position pour noms hébreux (à droite)
    x_nom_semaine_lat = 170   # Position pour noms latins (à gauche)
    x_horaire_semaine_heb = 170  # Position horaires quand nom hébreu (à gauche)
    x_horaire_semaine_lat = 800  # Position horaires quand nom latin (à droite)
    line_height_semaine = 44
    y_semaine = y_titre + line_height_semaine + 40
    
    # CORRECTION: Parcourir toutes les entrées de horaires_semaine dynamiquement
    idx_semaine = 0
    for key, horaire in horaires_semaine.items():
        if horaire is not None:
            # Récupérer le nom original depuis la config
            nom_original = config['horaires_semaine'][key]['nom']
            y = y_semaine + idx_semaine * line_height_semaine
            
            if is_hebrew(nom_original):
                # Texte hébreu : inverse + aligne à droite
                nom_affiche = reverse_hebrew_text(nom_original)
                draw.text((x_nom_semaine_heb, y), nom_affiche, font=font_hebrew_green, fill=(0,128,0), anchor="ra")
                draw.text((x_horaire_semaine_heb, y), horaire.strftime('%H:%M'), font=font_hebrew_green, fill=(0,128,0), anchor="la")
            else:
                # Texte latin : normal + aligne à gauche
                draw.text((x_nom_semaine_lat, y), nom_original, font=font_hebrew_green, fill=(0,128,0), anchor="la")
                draw.text((x_horaire_semaine_lat, y), horaire.strftime('%H:%M'), font=font_hebrew_green, fill=(0,128,0), anchor="ra")
            idx_semaine += 1

    # 6. Tekoufa (si applicable) - bleu ciel, centré, plus haut que les autres annonces
    if tekufa_announcement:
        y_tekufa = img.height - 200
        tekufa_text = reverse_hebrew_text(tekufa_announcement)
        bbox = draw.textbbox((0, 0), tekufa_text, font=font_hebrew)
        w = bbox[2] - bbox[0]
        x_text = (img.width - w) // 2
        draw.text((x_text, y_tekufa), tekufa_text, font=font_hebrew, fill=(0, 180, 255))
        
        # Icône eau.png à côté du texte
        icon_path = "resources/eau.png"
        if os.path.exists(icon_path):
            try:
                icon = Image.open(icon_path).convert("RGBA")
                icon = icon.resize((32, 32))
                # Place l'icône à gauche du texte
                img.paste(icon, (x_text - 40, y_tekufa), icon)
            except Exception as e:
                print(f"Erreur affichage icône tekoufa: {e}")
        else:
            print(f"⚠️ Icône non trouvée: {icon_path}")

    # 7. Birkat HaLevana (si applicable)
    if birkat_halevana_text:
        y_birkat = img.height - 120
        birkat_text = reverse_hebrew_text(birkat_halevana_text)
        bbox = draw.textbbox((0, 0), birkat_text, font=font_hebrew)
        w = bbox[2] - bbox[0]
        x_text = (img.width - w) // 2
        draw.text((x_text, y_birkat), birkat_text, font=font_hebrew, fill=(0, 51, 204))
        
        # Ajout de l'icône correspondante
        icon_path = get_birkat_halevana_icon(birkat_halevana_text)
        if icon_path and os.path.exists(icon_path):
            try:
                icon = Image.open(icon_path).convert("RGBA")
                icon = icon.resize((32, 32))
                img.paste(icon, (x_text - 40, y_birkat), icon)
            except Exception as e:
                print(f"Erreur affichage icône Birkat HaLevana: {e}")
        elif icon_path:
            print(f"⚠️ Icône non trouvée: {icon_path}")

    # 8. Ajout du molad et des jours de Rosh Hodesh en bas de page si Shabbat Mevarchim
    if is_mevarchim and (molad_text or rosh_hodesh_text):
        y_bas = img.height - 150
        if birkat_halevana_text:
            y_bas -= 40
        if molad_text:
            molad_text_r = reverse_hebrew_text(molad_text)
            bbox_molad = draw.textbbox((0, 0), molad_text_r, font=font_hebrew)
            w_molad = bbox_molad[2] - bbox_molad[0]
            x_molad = (img.width - w_molad) // 2
            draw.text((x_molad, y_bas), molad_text_r, font=font_hebrew, fill=(0, 51, 204))
            y_bas += 32
        if rosh_hodesh_text:
            rosh_hodesh_text_r = reverse_hebrew_text(rosh_hodesh_text)
            bbox_rh = draw.textbbox((0, 0), rosh_hodesh_text_r, font=font_hebrew)
            w_rh = bbox_rh[2] - bbox_rh[0]
            x_rh = (img.width - w_rh) // 2
            draw.text((x_rh, y_bas), rosh_hodesh_text_r, font=font_hebrew, fill=(0, 51, 204))

    img.save('output/latest-schedule.jpg')
    print("Image sauvegardée: output/latest-schedule.jpg")

def main():
    print("=== Générateur d'horaires de communauté ===\n")
    date_test = None
    if len(sys.argv) > 1:
        try:
            date_test = datetime.strptime(sys.argv[1], "%Y-%m-%d")
            print(f"Mode test : date simulée = {date_test.strftime('%A %d/%m/%Y')}")
        except Exception:
            print("Format de date incorrect, attendu : YYYY-MM-DD")
            return

    if not os.path.exists('config.json'):
        print("Erreur: Le fichier config.json n'existe pas.")
        return
    with open('config.json', encoding='utf-8') as f:
        config = json.load(f)
    lat = float(config['latitude'])
    lon = float(config['longitude'])
    print(f"Localisation: {config.get('nom_communaute', 'Communauté')} ({lat}, {lon})")
    print("\n--- Récupération des horaires halakhiques ---")
    zmanim = get_all_zmanim(lat, lon, date_test)
    print(zmanim)
    print("\n--- Calcul des horaires des activités ---")
    horaires_activites = []
    for act in config.get('activites_shabbat', []):
        horaire = calculer_horaire_struct(act, zmanim)
        horaires_activites.append({'nom': act['nom'], 'horaire': horaire})
        if horaire:
            print(f"✓ {act['nom']}: {horaire.strftime('%H:%M')}")
        else:
            print(f"✗ {act['nom']}: Impossible de calculer")
    
    # CORRECTION: Calcul dynamique des horaires de semaine
    horaires_semaine = {}
    print("\n--- Calcul des horaires de semaine ---")
    for key, office_struct in config.get('horaires_semaine', {}).items():
        horaire = calculer_horaire_struct(office_struct, zmanim)
        horaires_semaine[key] = horaire
        if horaire:
            print(f"✓ {office_struct['nom']}: {horaire.strftime('%H:%M')}")
        else:
            print(f"✗ {office_struct['nom']}: Impossible de calculer")

    friday, saturday, _ = get_next_friday_saturday_tuesday(date_test)
    parasha_he = get_parasha_name_hebrew(lat, lon, saturday)
    print(f"\nParasha de la semaine (hébreu) : {parasha_he}")

    is_mevarchim = is_shabbat_mevarchim(lat, lon, saturday)
    template_file = 'resources/template.jpg'

    molad_text = None
    rosh_hodesh_text = None
    if is_mevarchim:
        print("Shabbat Mevarchim détecté.")
        if ZMANIM_AVAILABLE:
            molad_text = get_next_month_molad(saturday)
            rosh_hodesh_text = get_rosh_hodesh_days_for_next_month(saturday)
            print("Molad :", molad_text)
            print("Rosh Hodesh :", rosh_hodesh_text)
        else:
            print("⚠️ Bibliothèque zmanim non disponible pour le calcul du molad.")
    else:
        print("Shabbat classique.")

    # Gestion de Birkat HaLevana
    birkat_halevana_text = get_birkat_halevana_announcement(saturday, is_mevarchim)
    if birkat_halevana_text:
        print("Birkat HaLevana :", birkat_halevana_text)
    else:
        print("Pas d'annonce Birkat HaLevana cette semaine.")

    # Annonce Tekoufa
    tekufa_announcement = None
    tekufa_info = get_next_tekufa_announcement(saturday)
    if tekufa_info:
        tekufa_title, tekufa_date, tekufa_hour = tekufa_info
        date_str = tekufa_date.strftime('%d/%m/%Y')
        if tekufa_hour:
            tekufa_announcement = f"{tekufa_title} תחול ביום {date_str} בשעה {tekufa_hour}"
        else:
            tekufa_announcement = f"{tekufa_title} תחול ביום {date_str}"
        print("Tekoufa à annoncer :", tekufa_announcement)
    else:
        print("Pas de tekoufa à annoncer cette semaine.")

    print("\n--- Génération de l'image ---")
    generer_image(
        config, horaires_activites, horaires_semaine, parasha_he, template_file,
        is_mevarchim, molad_text, rosh_hodesh_text, birkat_halevana_text, tekufa_announcement
    )
    print("\n✅ Image générée avec succès dans output/latest-schedule.jpg")

if __name__ == "__main__":
    main()