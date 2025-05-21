import os
import sys
from pathlib import Path
from datetime import datetime, date, timedelta
import math
import requests
import pytz
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from astral import LocationInfo
from astral.sun import sun
import unicodedata
import re
import shutil
from zmanim.hebrew_calendar.jewish_calendar import JewishCalendar

# ===== CONFIGURATION MOLAD =====
HEBREW_MONTHS = {
    1: 'Nissan', 2: 'Iyar', 3: 'Sivan', 4: 'Tamouz',
    5: 'Av', 6: 'Eloul', 7: 'Tishrei', 8: 'Heshvan',
    9: 'Kislev', 10: 'Tevet', 11: 'Shevat', 12: 'Adar',
    13: 'Adar II'
}

def get_jewish_month_name(jm, jy):
    if jm == 12 and JewishCalendar.is_jewish_leap_year(jy):
        return 'Adar I'
    if jm == 13:
        return 'Adar II'
    return HEBREW_MONTHS.get(jm, 'Mois-inconnu')

def calculate_molad_for_date(gregorian_date):
    jc = JewishCalendar(datetime.combine(gregorian_date, datetime.min.time()))
    molad_obj = jc.molad()
    return {
        "rosh_chodesh_date": gregorian_date,
        "molad": f"{molad_obj.molad_hours:02d}:{molad_obj.molad_minutes:02d} +{molad_obj.molad_chalakim} chalakim",
        "hebrew_month": get_jewish_month_name(jc.jewish_month, jc.jewish_year),
        "hebrew_year": jc.jewish_year
    }

def find_next_rosh_chodesh(start_date=None):
    current = start_date or date.today()
    for _ in range(60):
        jc = JewishCalendar(datetime.combine(current, datetime.min.time()))
        if jc.jewish_day == 1:
            return current
        current += timedelta(days=1)
    raise RuntimeError("No Rosh Chodesh in next 60 days.")

# ===== CLASSE PRINCIPALE =====
class ShabbatScheduleGenerator:
    def __init__(self, template_path, font_path, arial_bold_path, output_dir):
        self.template_path = Path(template_path)
        self.font_path = Path(font_path)
        self.arial_bold_path = Path(arial_bold_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Vérification des fichiers
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template introuvable: {self.template_path}")
        if not self.font_path.exists():
            raise FileNotFoundError(f"Police introuvable: {self.font_path}")
        if not self.arial_bold_path.exists():
            raise FileNotFoundError(f"Police Arial Bold introuvable: {self.arial_bold_path}")

        # Chargement des polices
        try:
            self._font = ImageFont.truetype(str(self.font_path), 30)
            self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)
        except Exception as e:
            raise Exception(f"Erreur de chargement de la police: {e}")

        # Configuration initiale
        self.season = self.determine_season()
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)
        
        # Données Shabbat
        self.yearly_shabbat_data = [
            {'day': '2024-12-06 00:00:00', 'פרשה': 'ויצא', 'כנסית שבת': '16:17', 'צאת שבת': '17:16'},
            {'day': '2024-12-13 00:00:00', 'פרשה': 'וישלח', 'כנסית שבת': '16:19', 'צאת שבת': '17:17'},
            {'day': '2024-12-20 00:00:00', 'פרשה': 'וישב', 'כנסית שבת': '16:22', 'צאת שבת': '17:20'},
            {'day': '2024-12-27 00:00:00', 'פרשה': 'מקץ', 'כנסית שבת': '16:25', 'צאת שבת': '17:24'},
            {'day': '2025-01-03 00:00:00', 'פרשה': 'ויגש', 'כנסית שבת': '16:30', 'צאת שבת': '17:29'},
            {'day': '2025-01-10 00:00:00', 'פרשה': 'ויחי', 'כנסית שבת': '16:36', 'צאת שבת': '17:35'},
            {'day': '2025-01-17 00:00:00', 'פרשה': 'שמות', 'כנסית שבת': '16:42', 'צאת שבת': '17:41'},
            {'day': '2025-01-24 00:00:00', 'פרשה': 'וארא', 'כנסית שבת': '16:49', 'צאת שבת': '17:47'},
            {'day': '2025-01-31 00:00:00', 'פרשה': 'בא', 'כנסית שבת': '16:55', 'צאת שבת': '17:53'},
            {'day': '2025-02-07 00:00:00', 'פרשה': 'בשלח', 'כנסית שבת': '17:02', 'צאת שבת': '17:58'},
            {'day': '2025-02-14 00:00:00', 'פרשה': 'יתרו', 'כנסית שבת': '17:08', 'צאת שבת': '18:04'},
            {'day': '2025-02-21 00:00:00', 'פרשה': 'משפטים', 'כנסית שבת': '17:14', 'צאת שבת': '18:10'},
            {'day': '2025-02-28 00:00:00', 'פרשה': 'תרומה', 'כנסית שבת': '17:19', 'צאת שבת': '18:15'},
            {'day': '2025-03-07 00:00:00', 'פרשה': 'תצוה', 'כנסית שבת': '17:25', 'צאת שבת': '18:20'},
            {'day': '2025-03-14 00:00:00', 'פרשה': 'כי-תשא', 'כנסית שבת': '17:30', 'צאת שבת': '18:25'},
            {'day': '2025-03-21 00:00:00', 'פרשה': 'ויקהל', 'כנסית שבת': '17:34', 'צאת שבת': '18:30'},
            {'day': '2025-03-28 00:00:00', 'פרשה': 'פקודי', 'כנסית שבת': '18:39', 'צאת שבת': '19:35'},
            {'day': '2025-04-04 00:00:00', 'פרשה': 'ויקרא', 'כנסית שבת': '18:44', 'צאת שבת': '19:40'},
            {'day': '2025-04-11 00:00:00', 'פרשה': 'צו', 'כנסית שבת': '18:49', 'צאת שבת': '19:45'},
            {'day': '2025-04-18 00:00:00', 'פרשה': 'פסח', 'כנסית שבת': '18:54', 'צאת שבת': '19:51'},
            {'day': '2025-04-25 00:00:00', 'פרשה': 'שמיני', 'כנסית שבת': '18:59', 'צאת שבת': '19:56'},
            {'day': '2025-05-02 00:00:00', 'פרשה': 'תזריע-מצורע', 'כנסית שבת': '19:04', 'צאת שבת': '20:02'},
            {'day': '2025-05-09 00:00:00', 'פרשה': 'אחרי-מות קדושים', 'כנסית שבת': '19:09', 'צאת שבת': '20:08'},
            {'day': '2025-05-16 00:00:00', 'פרשה': 'אמור', 'כנסית שבת': '19:14', 'צאת שבת': '20:13'},
            {'day': '2025-05-23 00:00:00', 'פרשה': 'בהר-בחוקותי', 'כנסית שבת': '19:18', 'צאת שבת': '20:19'},
            {'day': '2025-05-30 00:00:00', 'פרשה': 'במדבר', 'כנסית שבת': '19:23', 'צאת שבת': '20:23'},
            {'day': '2025-06-06 00:00:00', 'פרשה': 'נשא', 'כנסית שבת': '19:26', 'צאת שבת': '20:28'},
            {'day': '2025-06-13 00:00:00', 'פרשה': 'בהעלותך', 'כנסית שבת': '19:29', 'צאת שבת': '20:31'},
            {'day': '2025-06-20 00:00:00', 'פרשה': 'שלח', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
            {'day': '2025-06-27 00:00:00', 'פרשה': 'קרח', 'כנסית שבת': '19:33', 'צאת שבת': '20:33'},
            {'day': '2025-07-04 00:00:00', 'פרשה': 'חוקת', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
            {'day': '2025-07-11 00:00:00', 'פרשה': 'בלק', 'כנסית שבת': '19:31', 'צאת שבת': '20:31'},
            {'day': '2025-07-18 00:00:00', 'פרשה': 'פינחס', 'כנסית שבת': '19:28', 'צאת שבת': '20:27'},
            {'day': '2025-07-25 00:00:00', 'פרשה': 'מטות-מסעי', 'כנסית שבת': '19:24', 'צאת שבת': '20:23'},
            {'day': '2025-08-01 00:00:00', 'פרשה': 'דברים', 'כנסית שבת': '19:19', 'צאת שבת': '20:17'},
            {'day': '2025-08-08 00:00:00', 'פרשה': 'ואתחנן', 'כנסית שבת': '19:13', 'צאת שבת': '20:10'},
            {'day': '2025-08-15 00:00:00', 'פרשה': 'עקב', 'כנסית שבת': '19:06', 'צאת שבת': '20:02'},
            {'day': '2025-08-22 00:00:00', 'פרשה': 'ראה', 'כנסית שבת': '18:59', 'צאת שבת': '19:54'},
            {'day': '2025-08-29 00:00:00', 'פרשה': 'שופטים', 'כנסית שבת': '18:50', 'צאת שבת': '19:45'},
            {'day': '2025-09-05 00:00:00', 'פרשה': 'כי-תצא', 'כנסית שבת': '18:41', 'צאת שבת': '19:35'},
            {'day': '2025-09-12 00:00:00', 'פרשה': 'כי-תבוא', 'כנסית שבת': '18:32', 'צאת שבת': '19:26'},
            {'day': '2025-09-19 00:00:00', 'פרשה': 'ניצבים', 'כנסית שבת': '18:23', 'צאת שבת': '19:16'},
        ]

    # ===== FONCTIONS DE BASE =====
    def determine_season(self):
        """
        Détermine dynamiquement si nous sommes en heure d'été ou d'hiver en Israël.
        Par défaut, l'heure d'été est appliquée du dernier vendredi de mars jusqu'au dernier dimanche d'octobre.
        Pour tester en hiver, vous pouvez temporairement remplacer le contenu par :
            return "winter"
        """
        today = datetime.now()
        year = today.year
        start_summer = datetime(year, 3, 29)
        end_summer = datetime(year, 10, 26)
        return "summer" if start_summer <= today <= end_summer else "winter"

    def sanitize_filename(self, value: str) -> str:
        """
        Transforme une chaîne en un slug ASCII-safe pour les noms de fichiers:
        - Décompose les accents et supprime les caractères non-ASCII
        - Ne conserve que lettres, chiffres et espaces
        - Remplace les espaces par des underscores
        """
        # 1. Normalisation Unicode
        nfkd = unicodedata.normalize('NFKD', value)
        # 2. Encodage ASCII
        ascii_str = nfkd.encode('ascii', 'ignore').decode('ascii')
        # 3. Ne garder que alphanumériques et espaces
        ascii_str = re.sub(r'[^\w\s-]', '', ascii_str).strip()
        # 4. Remplacer les espaces par underscore
        return re.sub(r'\s+', '_', ascii_str)

    # ===== CALCULS HORAIRES =====
    def round_to_nearest_five(self, minutes):
        return (minutes // 5) * 5

    def format_time(self, minutes):
        if minutes is None or minutes < 0:
            return ""
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"

    def calculate_times(self, shabbat_start, shabbat_end):
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute

        times = {
            "mincha_kabbalat": start_minutes,
            "shir_hashirim": self.round_to_nearest_five(start_minutes - 10),
            "shacharit": self.round_to_nearest_five(7 * 60 + 45),
            "mincha_gdola": self.round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": self.round_to_nearest_five(14 * 60 + (0 if self.season == "winter" else 30)),
            "shiur_nashim": 16 * 60,
            "arvit_hol": None,
            "arvit_motsach": None,
            "mincha_2": None,
            "shiur_rav": None,
            "parashat_hashavua": None,
            "mincha_hol": None
        }

        # Garantir 10 minutes avant Shabbat
        if start_minutes - times["shir_hashirim"] < 10:
            times["shir_hashirim"] = start_minutes - 10

        # Calculs des autres horaires
        times["mincha_2"] = self.round_to_nearest_five(end_minutes - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        times["parashat_hashavua"] = self.round_to_nearest_five(times["shiur_rav"] - 45)
        
        # ... (calculs restants identiques à vos fichiers)

        return times

    # ===== GESTION MEVARCHIM =====
    def fetch_roshchodesh_dates(self, start_date, end_date):
        # ... (identique à votre code)

    def get_mevarchim_friday(self, rosh_date):
        # ... (identique à votre code)

    def identify_shabbat_mevarchim(self, shabbat_df, rosh_dates):
        # ... (identique à votre code)

    # ===== GENERATION IMAGE =====
    def reverse_hebrew_text(self, text):
        return text  # Ajouter une logique d'inversion si nécessaire

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting, shabbat_date, is_mevarchim=False):
        try:
            # Sélection template
            template = self.template_path
            if is_mevarchim:
                rc_template = self.template_path.parent / "template_rosh_hodesh.jpg"
                if rc_template.exists():
                    template = rc_template

            with Image.open(template) as img:
                draw = ImageDraw.Draw(img)
                font = self._font
                bold = self._arial_bold_font

                # Positions des éléments (ajuster selon votre disposition originale)
                positions = [
                    (120, 400, 'shir_hashirim'),
                    (120, 475, 'mincha_kabbalat'),
                    (120, 510, 'shacharit'),
                    (120, 550, 'mincha_gdola'),
                    (120, 590, 'tehilim'),
                    # ... autres positions
                ]

                # Dessin des horaires
                for x, y, key in positions:
                    draw.text((x, y), self.format_time(times[key]), fill="black", font=font)

                # Parasha
                reversed_parasha = self.reverse_hebrew_text(parasha_hebrew)
                draw.text((300, 280), reversed_parasha, fill="blue", font=bold, anchor="mm")

                # Molad pour Mevarchim
                if is_mevarchim:
                    rc_date = find_next_rosh_chodesh(shabbat_date)
                    molad_info = calculate_molad_for_date(rc_date)
                    molad_text = f"מולד: {molad_info['molad']}"
                    draw.text((100, 950), molad_text, fill="blue", font=font)

                # Sauvegarde
                safe_parasha = self.sanitize_filename(parasha)
                output_path = self.output_dir / f"horaires_{safe_parasha}.jpeg"
                img.save(str(output_path))
                
                return output_path

        except Exception as e:
            print(f"Erreur création image: {e}")
            return None

    # ===== GESTION EXCEL =====
    def update_excel_with_mevarchim_column(self, excel_path: Path):
        # ... (identique à votre code)

    def update_excel(self, shabbat_data, times):
        # ... (code de mise à jour Excel avec colonne supplémentaire)

    # ===== FLUX PRINCIPAL =====
    def generate(self):
        # ... (logique principale combinée de vos fichiers)

# ===== POINT D'ENTREE =====
def main():
    try:
        # ... (configuration des chemins)
        generator = ShabbatScheduleGenerator(template_path, font_path, arial_bold_path, output_dir)
        generator.update_excel_with_mevarchim_column(generator.output_dir / "horaires_shabbat.xlsx")
        generator.generate()
    except Exception as e:
        print(f"Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")

if __name__ == "__main__":
    main()