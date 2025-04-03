import os
import sys
from pathlib import Path
from datetime import datetime, timedelta
import math
import requests
import pytz
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from astral import LocationInfo
from astral.sun import sun

class ShabbatScheduleGenerator:
    def __init__(self, template_path, font_path, arial_bold_path, output_dir):
        self.template_path = Path(template_path)
        self.font_path = Path(font_path)
        self.arial_bold_path = Path(arial_bold_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)  # Création du dossier de sortie

        # Vérifie que les fichiers nécessaires existent
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template introuvable: {self.template_path}")
        if not self.font_path.exists():
            raise FileNotFoundError(f"Police introuvable: {self.font_path}")
        if not self.arial_bold_path.exists():
            raise FileNotFoundError(f"Police Arial Bold introuvable: {self.arial_bold_path}")

        try:
            # Chargement des polices
            self._font = ImageFont.truetype(str(self.font_path), 30)
            self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)
        except Exception as e:
            raise Exception(f"Erreur de chargement de la police: {e}")

        # Détermination automatique de la saison (été ou hiver)
        self.season = self.determine_season()

        # Localisation pour Ramat Gan, Israël (pour les calculs solaires)
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

        # Données de l'onglet "שבתות השנה" (pour Excel)
        self.yearly_shabbat_data = [
            {'day': '2024-12-06 00:00:00', 'פרשה': 'Vayetzei', 'כנסית שבת': '16:17', 'צאת שבת': '17:16'},
            {'day': '2024-12-13 00:00:00', 'פרשה': 'Vayishlach', 'כנסית שבת': '16:19', 'צאת שבת': '17:17'},
            {'day': '2024-12-20 00:00:00', 'פרשה': 'Vayeshev', 'כנסית שבת': '16:22', 'צאת שבת': '17:20'},
            {'day': '2024-12-27 00:00:00', 'פרשה': 'Miketz', 'כנסית שבת': '16:25', 'צאת שבת': '17:24'},
            {'day': '2025-01-03 00:00:00', 'פרשה': 'Vayigash', 'כנסית שבת': '16:30', 'צאת שבת': '17:29'},
            {'day': '2025-01-10 00:00:00', 'פרשה': 'Vayechi', 'כנסית שבת': '16:36', 'צאת שבת': '17:35'},
            {'day': '2025-01-17 00:00:00', 'פרשה': 'Shemot', 'כנסית שבת': '16:42', 'צאת שבת': '17:41'},
            {'day': '2025-01-24 00:00:00', 'פרשה': 'Vaera', 'כנסית שבת': '16:49', 'צאת שבת': '17:47'},
            {'day': '2025-01-31 00:00:00', 'פרשה': 'Bo', 'כנסית שבת': '16:55', 'צאת שבת': '17:53'},
            {'day': '2025-02-07 00:00:00', 'פרשה': 'Beshalach', 'כנסית שבת': '17:02', 'צאת שבת': '17:58'},
            {'day': '2025-02-14 00:00:00', 'פרשה': 'Yitro', 'כנסית שבת': '17:08', 'צאת שבת': '18:04'},
            {'day': '2025-02-21 00:00:00', 'פרשה': 'Mishpatim', 'כנסית שבת': '17:14', 'צאת שבת': '18:10'},
            {'day': '2025-02-28 00:00:00', 'פרשה': 'Terumah', 'כנסית שבת': '17:19', 'צאת שבת': '18:15'},
            {'day': '2025-03-07 00:00:00', 'פרשה': 'Tetzaveh', 'כנסית שבת': '17:25', 'צאת שבת': '18:20'},
            {'day': '2025-03-14 00:00:00', 'פרשה': 'Ki Tisa', 'כנסית שבת': '17:30', 'צאת שבת': '18:25'},
            {'day': '2025-03-21 00:00:00', 'פרשה': 'Vayakhel', 'כנסית שבת': '17:34', 'צאת שבת': '18:30'},
            {'day': '2025-03-28 00:00:00', 'פרשה': 'Pekudei', 'כנסית שבת': '18:39', 'צאת שבת': '19:35'},
            {'day': '2025-04-04 00:00:00', 'פרשה': 'Vayikra', 'כנסית שבת': '18:44', 'צאת שבת': '19:40'},
            {'day': '2025-04-11 00:00:00', 'פרשה': 'Tzav', 'כנסית שבת': '18:49', 'צאת שבת': ''},
            {'day': '2025-04-18 00:00:00', 'פרשה': 'Pesach', 'כנסית שבת': '18:54', 'צאת שבת': '19:51'},
            {'day': '2025-04-25 00:00:00', 'פרשה': 'Shmini', 'כנסית שבת': '18:59', 'צאת שבת': '19:56'},
            {'day': '2025-05-02 00:00:00', 'פרשה': 'Tazria-Metzora', 'כנסית שבת': '19:04', 'צאת שבת': '20:02'},
            {'day': '2025-05-09 00:00:00', 'פרשה': 'Achrei Mot-Kedoshim', 'כנסית שבת': '19:09', 'צאת שבת': '20:08'},
            {'day': '2025-05-16 00:00:00', 'פרשה': 'Emor', 'כנסית שבת': '19:14', 'צאת שבת': '20:13'},
            {'day': '2025-05-23 00:00:00', 'פרשה': 'Behar-Bechukotai', 'כנסית שבת': '19:18', 'צאת שבת': '20:19'},
            {'day': '2025-05-30 00:00:00', 'פרשה': 'Bamidbar', 'כנסית שבת': '19:23', 'צאת שבת': '20:23'},
            {'day': '2025-06-06 00:00:00', 'פרשה': 'Nasso', 'כנסית שבת': '19:26', 'צאת שבת': '20:28'},
            {'day': '2025-06-13 00:00:00', 'פרשה': 'Beha’alotcha', 'כנסית שבת': '19:29', 'צאת שבת': '20:31'},
            {'day': '2025-06-20 00:00:00', 'פרשה': 'Sh’lach', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
            {'day': '2025-06-27 00:00:00', 'פרשה': 'Korach', 'כנסית שבת': '19:33', 'צאת שבת': '20:33'},
            {'day': '2025-07-04 00:00:00', 'פרשה': 'Chukat', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
            {'day': '2025-07-11 00:00:00', 'פרשה': 'Balak', 'כנסית שבת': '19:31', 'צאת שבת': '20:31'},
            {'day': '2025-07-18 00:00:00', 'פרשה': 'Pinchas', 'כנסית שבת': '19:28', 'צאת שבת': '20:27'},
            {'day': '2025-07-25 00:00:00', 'פרשה': 'Matot-Masei', 'כנסית שבת': '19:24', 'צאת שבת': '20:23'},
            {'day': '2025-08-01 00:00:00', 'פרשה': 'Devarim', 'כנסית שבת': '19:19', 'צאת שבת': '20:17'},
            {'day': '2025-08-08 00:00:00', 'פרשה': 'Vaetchanan', 'כנסית שבת': '19:13', 'צאת שבת': '20:10'},
            {'day': '2025-08-15 00:00:00', 'פרשה': 'Eikev', 'כנסית שבת': '19:06', 'צאת שבת': '20:02'},
            {'day': '2025-08-22 00:00:00', 'פרשה': 'Re’eh', 'כנסית שבת': '18:59', 'צאת שבת': '19:54'},
            {'day': '2025-08-29 00:00:00', 'פרשה': 'Shoftim', 'כנסית שבת': '18:50', 'צאת שבת': '19:45'},
            {'day': '2025-09-05 00:00:00', 'פרשה': 'Ki Teitzei', 'כנסית שבת': '18:41', 'צאת שבת': '19:35'},
            {'day': '2025-09-12 00:00:00', 'פרשה': 'Ki Tavo', 'כנסית שבת': '18:32', 'צאת שבת': '19:26'},
            {'day': '2025-09-19 00:00:00', 'פרשה': 'Nitzavim', 'כנסית שבת': '18:23', 'צאת שבת': '19:16'},
        ]

    def determine_season(self):
        """
        Détermine dynamiquement si nous sommes en heure d'été ou d'hiver en Israël.
        L'heure d'été (summer) est appliquée généralement du dernier vendredi de mars
        jusqu'au dernier dimanche d'octobre.
        """
        today = datetime.now()
        year = today.year
        start_summer = datetime(year, 3, 29)
        end_summer = datetime(year, 10, 26)
        return "summer" if start_summer <= today <= end_summer else "winter"

    def get_hebcal_times(self, start_date, end_date):
        """
        Récupère via l’API Hebcal les horaires du Shabbat (candles et havdalah)
        et retourne une liste d'événements avec le nom de la parasha, l'heure de fin, etc.
        """
        tz = pytz.timezone("Asia/Jerusalem")
        base_url = "https://www.hebcal.com/shabbat"
        params = {
            "cfg": "json",
            "geonameid": "293397",  # Ramat Gan
            "b": "18",            # Heure de fin (Havdalah)
            "M": "on",            # Inclure la parasha
            "start": start_date.strftime("%Y-%m-%d"),
            "end": end_date.strftime("%Y-%m-%d")
        }
        try:
            response = requests.get(base_url, params=params)
            response.raise_for_status()
            data = response.json()
            shabbat_times = []
            for item in data["items"]:
                if item["category"] == "candles":
                    start_time = datetime.fromisoformat(item["date"]).astimezone(tz)
                    havdalah_items = [i for i in data["items"] if i["category"] == "havdalah"]
                    parasha_items = [i for i in data["items"] if i["category"] == "parashat"]
                    if havdalah_items and parasha_items:
                        end_time = datetime.fromisoformat(havdalah_items[0]["date"]).astimezone(tz)
                        parasha = parasha_items[0]["title"].replace("Parashat ", "")
                        parasha_hebrew = parasha_items[0].get("hebrew", parasha)
                        shabbat_times.append({
                            "date": start_time.date(),
                            "start": start_time,
                            "end": end_time,
                            "parasha": parasha,
                            "parasha_hebrew": parasha_hebrew,
                            "candle_lighting": start_time.strftime("%H:%M")
                        })
            return shabbat_times
        except requests.RequestException as e:
            print(f"Erreur lors de la récupération des données: {e}")
            return []

    def round_to_nearest_five(self, minutes):
        """Arrondit les minutes à la baisse au multiple de 5."""
        return (minutes // 5) * 5

    def format_time(self, minutes):
        """Transforme un nombre total de minutes en chaîne au format HH:MM."""
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"

    def compute_weekend_times(self, shabbat_event):
        """
        Calcule les horaires du Shabbat en se basant sur l'heure d'entrée (candles)
        et la fin du Shabbat (havdalah) selon les règles suivantes :
          • La "מנחה של הערב" (veille) est l'heure d'entrée.
          • "שיר השירים" est 10 minutes avant l'entrée (arrondi).
          • "שחרית" est fixé à 07:45.
          • "מנחה גדולה" est fixée à 12:30 en hiver et 13:00 en été.
          • "תהילים" est fixé 45 minutes après (en hiver, il n'y a qu'un affichage unique, alors qu'en été on affichera "17:00/HH:MM").
          • Le cours de פרשת השבוע est 3 heures avant la fin.
          • "ערבית" final est 10 minutes avant la fin.
          • "מנחה 2" est 1h30 avant "ערבית".
          • Le cours du Rav est 45 minutes avant "מנחה 2".
          • Le cours des femmes est fixé à 16:00.
        """
        candle_dt = shabbat_event["start"]
        shabbat_end_dt = shabbat_event["end"]
        candle_minutes = candle_dt.hour * 60 + candle_dt.minute
        end_minutes = shabbat_end_dt.hour * 60 + shabbat_end_dt.minute

        times = {
            "mincha_kabbalat": candle_minutes,
            "shir_hashirim": self.round_to_nearest_five(candle_minutes - 10),
            "shacharit": 7 * 60 + 45,
            "mincha_gdola": self.round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": self.round_to_nearest_five(13 * 60 + 45),
            "parashat_hashavua": self.round_to_nearest_five(end_minutes - 180),
            "arvit": self.round_to_nearest_five(end_minutes - (5 if self.season == "winter" else 10)),
            "mincha_2": None,  # Calculé ci-dessous
            "shiur_rav": None,  # Calculé ci-dessous
            "shiur_nashim": 16 * 60
        }
        times["mincha_2"] = self.round_to_nearest_five(times["arvit"] - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        return times

    def get_next_shabbat_time(self, current_shabbat_date):
        """
        Recherche, dans les données annuelles, la date et l'heure (arrondie)
        du prochain Shabbat.
        """
        try:
            current_date = datetime(current_shabbat_date.year, current_shabbat_date.month, current_shabbat_date.day)
            change_time_date = datetime(2025, 3, 27)
            last_shabbat = None
            for shabbat in self.yearly_shabbat_data:
                shabbat_date = datetime.strptime(shabbat["day"], "%Y-%m-%d %H:%M:%S")
                if shabbat_date > current_date:
                    if shabbat_date > change_time_date and last_shabbat:
                        shabbat = last_shabbat
                    shabbat_entry_time = shabbat["כנסית שבת"]
                    hours, minutes = map(int, shabbat_entry_time.split(':'))
                    total_minutes = hours * 60 + minutes
                    mincha_weekday = self.round_to_nearest_five(total_minutes)
                    return shabbat_date.strftime("%d/%m/%Y"), self.format_time(mincha_weekday)
                last_shabbat = shabbat
            return None, None
        except Exception as e:
            print(f"Erreur lors de la récupération du Shabbat suivant: {e}")
            return None, None

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting, shabbat_date):
        """
        Crée l'image des horaires du Shabbat avec le design suivant :
          • Affichage des horaires aux positions prédéfinies (X=120) :
              - 400px : 'שיר השירים'
              - 475px : 'mincha_kabbalat'
              - 510px : 'שחרית'
              - 550px : 'מנחה גדולה'
              - 590px : 'פרשת השבוע'
              - 630px : 'תהילים' (si été, affiché sous le format "17:00/HH:MM", sinon affiché normalement)
              - 670px : 'שיעור עם הרב'
              - 710px : 'שיעור לנשים'
              - 750px : 'מנחה 2'
              - 790px : 'ערבית'
          • L'heure de fin de Shabbat ("מוצאי שבת קודש") est affichée à 830px.
          • Le nom de la parasha en hébreu est affiché en bleu dans un carré en haut à gauche (position 300,280, centré).
          • L'heure de כניסת שבת est affichée en haut (à 440px).
          • En vert, on affiche :
              - "מנחה ביניים" à 950px (calculé à partir du coucher de soleil du dimanche et du jeudi),
              - et l'heure d'ערבית en milieu de semaine à 990px.
        """
        try:
            print("Ouverture du template...")
            print(f"Chemin du template : {self.template_path}")
            print(f"Chemin de la police : {self.font_path}")
            print(f"Chemin de la police Arial Bold : {self.arial_bold_path}")

            with Image.open(self.template_path) as img:
                print("Template ouvert avec succès.")
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(str(self.font_path), 30)

                time_x = 120
                time_positions = [
                    (time_x, 400, 'shir_hashirim'),
                    (time_x, 475, 'mincha_kabbalat'),
                    (time_x, 510, 'shacharit'),
                    (time_x, 550, 'mincha_gdola'),
                    (time_x, 590, 'tehilim'),  # Pour tehilim : affichage conditionnel
                    (time_x, 630, 'shiur_nashim'),
                    (time_x, 670, 'parashat_hashavua'),
                    (time_x, 710, 'shiur_rav'),
                    (time_x, 750, 'mincha_2'),
                    (time_x, 790, 'arvit')
                ]
                for x, y, key in time_positions:
                    if key == 'tehilim':
                        if self.season == "summer":
                            # En été, on affiche "17:00/HH:MM" avec décalage
                            formatted = "17:00/" + self.format_time(times['tehilim'])
                            draw.text((x - 40, y), formatted, fill="black", font=font)
                        else:
                            # En hiver, on affiche uniquement le temps calculé
                            formatted = self.format_time(times['tehilim'])
                            draw.text((x, y), formatted, fill="black", font=font)
                    else:
                        formatted = self.format_time(times[key])
                        draw.text((x, y), formatted, fill="black", font=font)

                # Affichage de l'heure de fin de Shabbat ("מוצאי שבת קודש") à Y=830
                end_time_str = shabbat_end.strftime("%H:%M")
                draw.text((time_x, 830), end_time_str, fill="black", font=font)

                # Affichage du nom de la parasha en hébreu dans le carré en haut à gauche (en bleu, centré)
                draw.text((300, 280), parasha_hebrew, fill="blue", font=self._arial_bold_font, anchor="mm")

                # Affichage de l'heure de כניסת שבת en haut (Y=440)
                draw.text((time_x, 440), candle_lighting, fill="black", font=font)

                # Calcul et affichage de "מנחה ביניים"
                # Utilise le coucher de soleil du dimanche (shabbat_date + 2 jours) et du jeudi (dimanche+4 jours)
                sunday_date = shabbat_date + timedelta(days=2)
                s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
                sunday_sunset = s_sunday["sunset"].strftime("%H:%M")

                thursday_date = sunday_date + timedelta(days=4)
                s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
                thursday_sunset = s_thursday["sunset"].strftime("%H:%M")

                def to_minutes(t):
                    h, m = map(int, t.split(':'))
                    return h * 60 + m
                if sunday_sunset and thursday_sunset:
                    base = min(to_minutes(sunday_sunset), to_minutes(thursday_sunset)) - 20
                    if base < 0:
                        minha_beynayim = ""
                    else:
                        minha_beynayim = self.format_time(self.round_to_nearest_five(base))
                else:
                    minha_beynayim = ""
                draw.text((time_x, 950), minha_beynayim, fill="green", font=font)

                # Calcul et affichage de l'heure d'ערבית en milieu de semaine à 990px
                arvit_time = self.round_to_nearest_five(times['mincha_kabbalat'] + 45)
                arvit_str = self.format_time(arvit_time)
                draw.text((time_x, 990), arvit_str, fill="green", font=font)

                output_path = self.output_dir / f"horaires_{parasha}.jpeg"
                print(f"Chemin de sortie de l'image : {output_path}")
                img.save(str(output_path))
                print("Image sauvegardée avec succès")
                return output_path

        except Exception as e:
            print(f"Erreur lors de la création de l'image: {e}")
            return None

    def update_excel(self, shabbat_data, times):
        """
        Met à jour (ou crée) le fichier Excel avec deux feuilles :
          - "Sheet1" pour les horaires du Shabbat courant,
          - "שבתות השנה" pour les horaires annuels enrichis (avec מנחה ביניים et צאת הכוכבים).
        """
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        next_shabbat_date, next_shabbat_time = self.get_next_shabbat_time(shabbat_data["date"])

        row = {
            "תאריך": shabbat_data["date"].strftime("%d/%m/%Y"),
            "פרשה": shabbat_data["parasha"],
            "שיר השירים": self.format_time(times["shir_hashirim"]),
            "כנסית שבת": shabbat_data["candle_lighting"],
            "מנחה": self.format_time(times["mincha_kabbalat"]),
            "שחרית": self.format_time(times["shacharit"]),
            "מנחה גדולה": self.format_time(times["mincha_gdola"]),
            "שיעור לנשים": self.format_time(times["shiur_nashim"]),
            "תהילים לילדים": self.format_time(times["tehilim"]),
            "שיעור פרשת השבוע": self.format_time(times["parashat_hashavua"]),
            "שיעור עם הרב": self.format_time(times["shiur_rav"]),
            "מנחה 2": self.format_time(times["mincha_2"]),
            "ערבית מוצ\"ש": self.format_time(times["arvit"]),
            "מוצאי שבת קודש": shabbat_data["end"].strftime("%H:%M"),
            "שבת הבאה (Date)": next_shabbat_date if next_shabbat_date else "N/A",
            "שבת הבאה (Heure)": next_shabbat_time if next_shabbat_time else "N/A"
        }

        try:
            # Création du DataFrame annuel pour l'onglet "שבתות השנה"
            yearly_df = pd.DataFrame(self.yearly_shabbat_data)
            
            def compute_times(row):
                row_date = datetime.strptime(row["day"], "%Y-%m-%d %H:%M:%S").date()
                sunday_date = row_date + timedelta(days=2)
                s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
                sunday_sunset = s_sunday["sunset"].strftime("%H:%M")
                sunday_dusk = s_sunday["dusk"].strftime("%H:%M")
                thursday_date = sunday_date + timedelta(days=4)
                s_thu = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
                thursday_sunset = s_thu["sunset"].strftime("%H:%M")
                thursday_dusk = s_thu["dusk"].strftime("%H:%M")
                def to_minutes(t):
                    h, m = map(int, t.split(":"))
                    return h * 60 + m
                if sunday_sunset and thursday_sunset:
                    base = min(to_minutes(sunday_sunset), to_minutes(thursday_sunset)) - 20
                    if base < 0:
                        minha_midweek = ""
                    else:
                        minha_midweek = self.format_time(self.round_to_nearest_five(base))
                else:
                    minha_midweek = ""
                return pd.Series({
                    "שקיעה Dimanche": sunday_sunset,
                    "שקיעה Jeudi": thursday_sunset,
                    "צאת הכוכבים Dimanche": sunday_dusk,
                    "צאת הכוכבים Jeudi": thursday_dusk,
                    "מנחה ביניים": minha_midweek
                })
            
            times_df = yearly_df.apply(compute_times, axis=1)
            yearly_df = pd.concat([yearly_df, times_df], axis=1)
            
            if excel_path.exists():
                with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_sheet1 = pd.DataFrame([row])
                    df_sheet1.to_excel(writer, sheet_name="Sheet1", index=False,
                                       startrow=writer.sheets["Sheet1"].max_row)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            else:
                with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:
                    df_sheet1 = pd.DataFrame([row])
                    df_sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            print(f"Excel mis à jour: {excel_path}")
        except Exception as e:
            print(f"Erreur lors de la mise à jour de l'Excel: {e}")

    def generate(self):
        """Point d'entrée : récupère les horaires du Shabbat, génère l'image et met à jour l'Excel."""
        current_date = datetime.now()
        end_date = current_date + timedelta(days=14)
        shabbat_times = self.get_hebcal_times(current_date, end_date)
        if not shabbat_times:
            print("Aucun horaire trouvé pour cette semaine")
            return

        shabbat = shabbat_times[0]
        times = self.compute_weekend_times(shabbat)
        image_path = self.create_image(times,
                                       shabbat["parasha"],
                                       shabbat["parasha_hebrew"],
                                       shabbat["end"],
                                       shabbat["candle_lighting"],
                                       shabbat["date"])
        if not image_path:
            print("Échec de la génération de l'image")
        self.update_excel(shabbat, times)

def main():
    try:
        if getattr(sys, "frozen", False):
            base_path = Path(sys.executable).parent
        elif "__file__" in globals():
            base_path = Path(__file__).parent
        else:
            base_path = Path.cwd()

        template_path = base_path / "resources" / "template.jpg"
        font_path = base_path / "resources" / "mriamc_0.ttf"
        arial_bold_path = base_path / "resources" / "ARIALBD_0.TTF"
        output_dir = base_path / "output"

        generator = ShabbatScheduleGenerator(template_path, font_path, arial_bold_path, output_dir)
        generator.generate()

    except Exception as e:
        print(f"Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")

if __name__ == "__main__":
    main()