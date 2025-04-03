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

        # Vérification de l'existence des fichiers
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

        # Configuration de la localisation pour Ramat Gan, Israël (calculs solaires)
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

        # Données de l'onglet "שבתות השנה" (pour Excel)
        self.yearly_shabbat_data = [
            {'day': '2024-12-06 00:00:00', 'פרשה': 'Vayetzei', 'כנסית שבת': '16:17', 'צאת שבת': '17:16'},
            {'day': '2024-12-13 00:00:00', 'פרשה': 'Vayishlach', 'כנסית שבת': '16:19', 'צאת שבת': '17:17'},
            {'day': '2024-12-20 00:00:00', 'פרשה': 'Vayeshev', 'כנסית שבת': '16:22', 'צאת שבת': '17:20'},
            {'day': '2024-12-27 00:00:00', 'פרשה': 'Miketz', 'כנסית שבת': '16:25', 'צאת שבת': '17:24'},
            {'day': '2025-01-03 00:00:00', 'פרשה': 'Vayigash', 'כנסית שבת': '16:30', 'צאת שבת': '17:29'},
            # ... (les autres données annuelles)
        ]

    def determine_season(self):
        """
        Détermine dynamiquement si nous sommes en heure d'été ou d'hiver en Israël.
        Par défaut, l'heure d'été est appliquée du dernier vendredi de mars
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
            "b": "18",             # Heure de fin (Havdalah)
            "M": "on",             # Inclure la parasha
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

    def reverse_hebrew_text(self, text):
        """Inverse le texte hébreu pour un affichage correct (notamment sur GitHub)."""
        return text[::-1]

    def compute_weekend_times(self, shabbat_event):
        """
        Calcule les horaires du Shabbat en se basant sur l'heure d'entrée (candles)
        et la fin du Shabbat (havdalah) selon les règles suivantes :
          • "מנחה של הערב" est l'heure d'entrée.
          • "שיר השירים" est 10 minutes avant l'entrée (arrondi).
          • "שחרית" est fixé à 07:45.
          • "מנחה גדולה" est fixée à 12:30 en hiver et 13:00 en été.
          • "תהילים" est fixé 45 minutes après (en hiver, affiché uniquement, en été sous le format "17:00/HH:MM").
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
            "mincha_2": None,   # Calculé ci-dessous
            "shiur_rav": None,   # Calculé ci-dessous
            "shiur_nashim": 16 * 60
        }
        times["mincha_2"] = self.round_to_nearest_five(times["arvit"] - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        return times

    def get_next_shabbat_time(self, current_shabbat_date):
        """
        Recherche, dans les données annuelles, la date et l'heure (arrondie) du prochain Shabbat.
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
          • Affichage des horaires à des positions fixes (X = 120) :
              - 400px : "שיר השירים"
              - 475px : "mincha_kabbalat"
              - 510px : "שחרית"
              - 550px : "מנחה גדולה"
              - 590px : "פרשת השבוע"
              - 630px : "תהילים" (en été : "17:00/HH:MM", en hiver : temps calculé)
              - 670px : "שיעור עם הרב"
              - 710px : "שיעור לנשים"
              - 750px : "מנחה 2"
              - 790px : "ערבית"
          • À 830px, on affiche l'heure de fin du Shabbat ("מוצאי שבת קודש").
          • En haut à gauche (position 300,280, centré), on affiche le nom de la parasha en bleu.
          • À 440px, l'heure de כניסת שבת est affichée.
          • En vert, on affiche :
                - "מנחה ביניים" à 950px (calculée à partir du coucher du soleil du dimanche et du jeudi),
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
                    (time_x, 590, 'tehilim'),
                    (time_x, 630, 'shiur_nashim'),
                    (time_x, 670, 'parashat_hashavua'),
                    (time_x, 710, 'shiur_rav'),
                    (time_x, 750, 'mincha_2'),
                    (time_x, 790, 'arvit')
                ]
                for x, y, key in time_positions:
                    if key == 'tehilim':
                        if self.season == "summer":
                            # En été, affichage sous le format "17:00/HH:MM" décalé
                            formatted = "17:00/" + self.format_time(times['tehilim'])
                            draw.text((x - 40, y), formatted, fill="black", font=font)
                        else:
                            # En hiver, affichage simple du temps calculé
                            formatted = self.format_time(times['tehilim'])
                            draw.text((x, y), formatted, fill="black", font=font)
                    else:
                        formatted = self.format_time(times[key])
                        draw.text((x, y), formatted, fill="black", font=font)

                # Affichage de l'heure de fin ("מוצאי שבת קודש") à 830px
                end_time_str = shabbat_end.strftime("%H:%M")
                draw.text((time_x, 830), end_time_str, fill="black", font=font)

                # Inverser le nom de la parasha pour un affichage correct
                reversed_parasha = self.reverse_hebrew_text(parasha_hebrew)
                # Afficher le nom de la parasha en bleu dans le carré en haut à gauche, centré
                draw.text((300, 280), reversed_parasha, fill="blue", font=self._arial_bold_font, anchor="mm")

                # Affichage de l'heure de כניסת שבת en haut (440px)
                draw.text((time_x, 440), candle_lighting, fill="black", font=font)

                # Calcul et affichage de "מנחה ביניים"
                # Basé sur le coucher du soleil du dimanche (shabbat_date + 2 jours) et du jeudi (dimanche + 4 jours)
                sunday_date = shabbat_date + timedelta(days=2)
                s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
                sunday_sunset = s_sunday["sunset"].strftime("%H:%M")

                thursday_date = sunday_date + timedelta(days=4)
                s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
                thursday_sunset = s_thursday["sunset"].strftime("%H:%M")

                def to_minutes(t):
                    h, m = map(int, t.split(":"))
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

                # Calcul et affichage de l'heure d'ערבית en milieu de semaine (990px)
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
            "תהילים לילדים": self.format_time(times["tehilim"]),
            "שיעור לנשים": self.format_time(times["shiur_nashim"]),
            "שיעור פרשת השבוע": self.format_time(times["parashat_hashavua"]),
            "שיעור עם הרב": self.format_time(times["shiur_rav"]),
            "מנחה 2": self.format_time(times["mincha_2"]),
            "ערבית מוצ\"ש": self.format_time(times["arvit"]),
            "מוצאי שבת קודש": shabbat_data["end"].strftime("%H:%M"),
            "שבת הבאה (Date)": next_shabbat_date if next_shabbat_date else "N/A",
            "שבת הבאה (Heure)": next_shabbat_time if next_shabbat_time else "N/A"
        }

        try:
            if excel_path.exists():
                with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=writer.sheets["Sheet1"].max_row)
                    yearly_df = pd.DataFrame(self.yearly_shabbat_data)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            else:
                with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name="Sheet1", index=False)
                    yearly_df = pd.DataFrame(self.yearly_shabbat_data)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            print(f"Excel mis à jour: {excel_path}")
        except Exception as e:
            print(f"Erreur lors de la mise à jour de l'Excel: {e}")

    def generate(self):
        current_date = datetime.now()
        end_date = current_date + timedelta(days=14)
        shabbat_times = self.get_hebcal_times(current_date, end_date)
        if not shabbat_times:
            print("Aucun horaire trouvé pour cette semaine")
            return

        shabbat = shabbat_times[0]
        # Utiliser compute_weekend_times (au lieu de calculate_times) pour obtenir les horaires
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