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
import unicodedata
import re
import shutil

class ShabbatScheduleGenerator:
    def __init__(self, template_path, font_path, arial_bold_path, output_dir):
        """
        Initialise le générateur de planning du Chabbat.
        """
        self.template_path = Path(template_path)
        self.font_path = Path(font_path)
        self.arial_bold_path = Path(arial_bold_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Vérifie que les fichiers nécessaires existent
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template introuvable: {self.template_path}")
        if not self.font_path.exists():
            raise FileNotFoundError(f"Police introuvable: {self.font_path}")
        if not self.arial_bold_path.exists():
            raise FileNotFoundError(f"Police Arial Bold introuvable: {self.arial_bold_path}")

        try:
            # Charge les polices
            self._font = ImageFont.truetype(str(self.font_path), 30)
            self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)
        except Exception as e:
            raise Exception(f"Erreur de chargement de la police: {e}")

        # Détermination automatique de la saison (été ou hiver)
        self.season = self.determine_season()

        # Configuration de la localisation pour Ramat Gan, Israël (pour les calculs astronomiques)
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

        # Données intégrées pour l'onglet "שבתות השנה" (pour Excel)
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

    def sanitize_filename(self, value: str) -> str:
        """Transforme une chaîne en un slug ASCII-safe pour les noms de fichiers."""
        nfkd = unicodedata.normalize('NFKD', value)
        ascii_str = nfkd.encode('ascii', 'ignore').decode('ascii')
        ascii_str = re.sub(r'[^\w\s-]', '', ascii_str).strip()
        return re.sub(r'\s+', '_', ascii_str)

    def determine_season(self):
        """Détermine si c’est été ou hiver."""
        today = datetime.now()
        year = today.year
        start_summer = datetime(year, 3, 29)
        end_summer = datetime(year, 10, 26)
        return "summer" if start_summer <= today <= end_summer else "winter"

    def fetch_roshchodesh_dates(self, start_date, end_date):
        """Récupère les premiers jours de Rosh Chodesh via l'API Hebcal."""
        url = "https://www.hebcal.com/hebcal "
        params = {
            "v": 1,
            "cfg": "json",
            "geonameid": "293397",
            "start": start_date.strftime("%Y-%m-%d"),
            "end": end_date.strftime("%Y-%m-%d"),
            "category": "roshchodesh",
            "nx": "on"
        }
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()
            rosh_dates = []
            seen_months = set()
            for item in data.get("items", []):
                if item.get("category") == "roshchodesh":
                    dt = datetime.fromisoformat(item["date"]).astimezone(pytz.timezone("Asia/Jerusalem")).date()
                    month_key = dt.strftime("%Y-%m")
                    if month_key not in seen_months:
                        rosh_dates.append(dt)
                        seen_months.add(month_key)
            return sorted(rosh_dates)
        except Exception as e:
            print(f"❌ Erreur lors de la récupération des Rosh Chodesh: {e}")
            return []

    def get_mevarchim_friday(self, rosh_date):
        """Trouve le vendredi avant Rosh Chodesh."""
        if rosh_date.weekday() == 4:
            return rosh_date - timedelta(days=7)
        elif rosh_date.weekday() == 5:
            return rosh_date - timedelta(days=8)
        else:
            delta = (rosh_date.weekday() - 4 + 7) % 7
            return rosh_date - timedelta(days=delta)

    def identify_shabbat_mevarchim(self, shabbat_df, rosh_dates):
        """Marque les vendredis strictement avant chaque Rosh Chodesh."""
        shabbat_df = shabbat_df.copy()
        shabbat_df["day"] = pd.to_datetime(shabbat_df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        mevarchim_set = set()
        for rd in rosh_dates:
            mevarchim_friday = self.get_mevarchim_friday(rd)
            if mevarchim_friday < rd and mevarchim_friday in shabbat_df["day"].values:
                mevarchim_set.add(mevarchim_friday)
        shabbat_df["שבת מברכין"] = shabbat_df["day"].isin(mevarchim_set)
        return shabbat_df

    def update_excel_with_mevarchim_column(self, excel_path: Path):
        """
        Met à jour l’onglet 'שבתות השנה' avec 'שבת מברכין'.
        Si le fichier n’existe pas, il est créé avec les données intégrées.
        """
        if not excel_path.exists():
            print("Fichier Excel non trouvé, création avec les données intégrées")
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        else:
            df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date

        min_date = df["day"].min()
        max_date = df["day"].max()
        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date + timedelta(days=7))
        df = self.identify_shabbat_mevarchim(df, rosh_dates)

        # Sauvegarde dans Excel
        with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name="שבתות השנה", index=False)
        print("✅ Colonne 'שבת מברכין' mise à jour dans Excel.")

    def get_shabbat_times_from_excel_file(self, current_date):
        """
        Récupère les horaires depuis Excel ou utilise les données internes.
        """
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        if excel_path.exists():
            try:
                df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
                df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
                today_date = current_date.date()
                df = df[df["day"] >= today_date].sort_values(by="day")
                if df.empty:
                    return None
                row = df.iloc[0]
                shabbat_date = datetime.combine(row["day"], datetime.min.time())
                try:
                    candle_time = datetime.strptime(str(row["כנסית שבת"]), "%H:%M").time()
                except Exception as e:
                    print("❌ Erreur lors de la lecture de l'heure 'כנסית שבת':", e)
                    return None
                try:
                    havdalah_time = datetime.strptime(str(row["צאת שבת"]), "%H:%M").time()
                except Exception as e:
                    print("❌ Erreur lors de la lecture de l'heure 'צאת שבת':", e)
                    return None
                shabbat_start = datetime.combine(row["day"], candle_time)
                shabbat_end = datetime.combine(row["day"], havdalah_time)
                is_mevarchim_excel = row.get("שבת מברכין", "") == "Oui"
                return [{
                    "date": shabbat_date,
                    "start": shabbat_start,
                    "end": shabbat_end,
                    "parasha": row.get("פרשה", ""),
                    "parasha_hebrew": row.get("פרשה_עברית", row.get("פרשה", "")),
                    "candle_lighting": row["כנסית שבת"],
                    "is_mevarchim": is_mevarchim_excel
                }]
            except Exception as e:
                print(f"❌ Erreur lors de la lecture du fichier Excel: {e}")
                return None
        else:
            print("Fichier Excel non trouvé, utilisation des données intégrées")
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
            roshchodesh_start = df["day"].min()
            roshchodesh_end = df["day"].max()
            rosh_dates = self.fetch_roshchodesh_dates(roshchodesh_start, roshchodesh_end)
            df = self.identify_shabbat_mevarchim(df, rosh_dates)
            row = df.iloc[0]
            shabbat_date = datetime.combine(row["day"], datetime.min.time())
            try:
                candle_time = datetime.strptime(str(row["כנסית שבת"]), "%H:%M").time()
            except Exception as e:
                print("❌ Erreur lors de la lecture de l'heure 'כנסית שבת':", e)
                return None
            try:
                havdalah_time = datetime.strptime(str(row["צאת שבת"]), "%H:%M").time()
            except Exception as e:
                print("❌ Erreur lors de la lecture de l'heure 'צאת שבת':", e)
                return None
            shabbat_start = datetime.combine(row["day"], candle_time)
            shabbat_end = datetime.combine(row["day"], havdalah_time)
            is_mevarchim_excel = row.get("שבת מברכין", "") == "Oui"
            return [{
                "date": shabbat_date,
                "start": shabbat_start,
                "end": shabbat_end,
                "parasha": row.get("פרשה", ""),
                "parasha_hebrew": row.get("פרשה_עברית", row.get("פרשה", "")),
                "candle_lighting": row["כנסית שבת"],
                "is_mevarchim": is_mevarchim_excel
            }]

    def get_hebcal_times(self, start_date, end_date):
        """Récupère les horaires via l'API Hebcal."""
        tz = pytz.timezone("Asia/Jerusalem")
        base_url = "https://www.hebcal.com/shabbat "
        params = {
            "cfg": "json",
            "geonameid": "293397",
            "b": "18",
            "M": "on",
            "start": start_date.strftime("%Y-%m-%d"),
            "end": end_date.strftime("%Y-%m-%d"),
            "lg": "he"
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
                        parasha_hebrew = parasha_items[0].get("hebrew", "").strip()
                        min_date = min(df["day"]) if "df" in locals() else start_date
                        max_date = max(df["day"]) if "df" in locals() else end_date
                        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date + timedelta(days=7))
                        shabbat_df = pd.DataFrame(self.yearly_shabbat_data)
                        shabbat_df["day"] = pd.to_datetime(shabbat_df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
                        shabbat_df = self.identify_shabbat_mevarchim(shabbat_df, rosh_dates)
                        is_mevarchim = shabbat_df[shabbat_df["day"] == start_time.date()]["שבת מברכין"].iloc[0]
                        shabbat_times.append({
                            "date": start_time.date(),
                            "start": start_time,
                            "end": end_time,
                            "parasha": parasha,
                            "parasha_hebrew": parasha_hebrew,
                            "candle_lighting": start_time.strftime("%H:%M"),
                            "is_mevarchim": is_mevarchim
                        })
            return shabbat_times
        except requests.RequestException as e:
            print(f"❌ Erreur lors de la récupération des données: {e}")
            return []

    def round_to_nearest_five(self, minutes):
        """Arrondit les minutes à la baisse au multiple de 5."""
        return (minutes // 5) * 5

    def format_time(self, minutes):
        """Transforme un total de minutes en HH:MM."""
        if minutes is None or minutes < 0:
            return ""
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"

    def reverse_hebrew_text(self, text):
        """Retourne le texte hébreu pour affichage."""
        return text

    def calculate_times(self, shabbat_start, shabbat_end):
        """
        Calcule les horaires du Chabbat.
        """
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute

        times = {
            "mincha_kabbalat": start_minutes,
            "shir_hashirim": self.round_to_nearest_five(start_minutes - 10),
            "shacharit": self.round_to_nearest_five(7 * 60 + 45),
            "mincha_gdola": self.round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": self.round_to_nearest_five(14 * 60),
            "shiur_nashim": 16 * 60,
            "arvit_hol": None,
            "arvit_motsach": None,
            "mincha_2": None,
            "shiur_rav": None,
            "parashat_hashavua": None,
            "mincha_hol": None
        }

        if start_minutes - times["shir_hashirim"] < 10:
            times["shir_hashirim"] = start_minutes - 10

        times["mincha_2"] = self.round_to_nearest_five(end_minutes - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        times["parashat_hashavua"] = self.round_to_nearest_five(times["shiur_rav"] - 45)

        # Arvit Hol (milieu de semaine, tsé hakochavim)
        sunday_date = shabbat_start.date() + timedelta(days=2)
        s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
        sunday_dusk = s_sunday.get("dusk", None)
        if sunday_dusk:
            sunday_dusk = sunday_dusk.strftime("%H:%M")
        else:
            sunday_dusk = None

        thursday_date = sunday_date + timedelta(days=4)
        s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
        thursday_dusk = s_thursday.get("dusk", None)
        if thursday_dusk:
            thursday_dusk = thursday_dusk.strftime("%H:%M")
        else:
            thursday_dusk = None

        def to_minutes(t):
            if t is None:
                return None
            try:
                h, m = map(int, t.split(":"))
                return h * 60 + m
            except ValueError:
                return None

        if sunday_dusk and thursday_dusk:
            tsé_dimanche = to_minutes(sunday_dusk)
            tsé_jeudi = to_minutes(thursday_dusk)
            if tsé_dimanche is None or tsé_jeudi is None:
                times["arvit_hol"] = 0
            else:
                moyenne = (tsé_dimanche + tsé_jeudi) // 2
                tsé_min = min(tsé_dimanche, tsé_jeudi)
                new_arvit_time = self.round_to_nearest_five(moyenne)
                if tsé_min - new_arvit_time > 3:
                    new_arvit_time = ((tsé_min - 3 + 4) // 5) * 5
                times["arvit_hol"] = new_arvit_time
        else:
            times["arvit_hol"] = 0

        # arvit_motsach : fin du Chabbat
        times["arvit_motsach"] = self.round_to_nearest_five(end_minutes - 5)

        # Min'ha Hol (coucher du soleil - 17min)
        sunday_sunset = s_sunday.get("sunset", None)
        if sunday_sunset:
            sunday_sunset = sunday_sunset.strftime("%H:%M")
        else:
            sunday_sunset = None

        thursday_sunset = s_thursday.get("sunset", None)
        if thursday_sunset:
            thursday_sunset = thursday_sunset.strftime("%H:%M")
        else:
            thursday_sunset = None

        if sunday_sunset and thursday_sunset:
            base = min(to_minutes(sunday_sunset), to_minutes(thursday_sunset)) - 17
            times["mincha_hol"] = self.round_to_nearest_five(base)
        else:
            times["mincha_hol"] = None

        return times

    def get_next_shabbat_time(self, current_shabbat_date):
        """Recherche la date et l'heure du prochain Chabbat."""
        try:
            current_date = current_shabbat_date.date() if isinstance(current_shabbat_date, datetime) else current_shabbat_date
            change_time_date = datetime(2025, 3, 27).date()
            last_shabbat = None
            for shabbat in self.yearly_shabbat_data:
                shabbat_date = datetime.strptime(shabbat["day"], "%Y-%m-%d %H:%M:%S").date()
                if shabbat_date > current_date:
                    if shabbat_date > change_time_date and last_shabbat:
                        shabbat = last_shabbat
                    shabbat_entry_time = shabbat["כנסית שבת"] if "כנסית שבת" in shabbat else shabbat["כנסתית שבת"]
                    hours, minutes = map(int, shabbat_entry_time.split(":"))
                    total_minutes = hours * 60 + minutes
                    mincha_weekday = self.round_to_nearest_five(total_minutes)
                    return shabbat_date.strftime("%d/%m/%Y"), self.format_time(mincha_weekday)
                last_shabbat = shabbat
            return None, None
        except Exception as e:
            print(f"❌ Erreur lors de la récupération du Chabbat suivant: {e}")
            return None, None

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting, shabbat_date, is_mevarchim=False):
        """Crée l’image des horaires du Chabbat avec le design défini."""
        try:
            print("Ouverture du template...")
            print(f"Chemin du template : {self.template_path}")
            print(f"Chemin de la police : {self.font_path}")
            print(f"Chemin de la police Arial Bold : {self.arial_bold_path}")

            if is_mevarchim:
                mevarchim_template = self.template_path.parent / "template_rosh_hodesh.jpg"
                if mevarchim_template.exists():
                    self.template_path = mevarchim_template
                else:
                    print("❌ Template spécial Rosh Chodesh introuvable → template standard utilisé")

            with Image.open(self.template_path) as img:
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
                    (time_x, 790, 'arvit_motsach'),   # Affichage explicite arvit de fin de shabbat (noir)
                ]
                for x, y, key in time_positions:
                    color = "black"
                    draw.text((x, y), self.format_time(times[key]), fill=color, font=font)

                # Affiche l’heure de fin du Chabbat ("מוצאי שבת Kodch") à 830px
                end_time_str = shabbat_end.strftime("%H:%M")
                draw.text((time_x, 830), end_time_str, fill="black", font=font)

                # Affiche le nom de la parasha en bleu à partir de la variable parasha_hebrew
                reversed_parasha = self.reverse_hebrew_text(parasha_hebrew)
                draw.text((300, 280), reversed_parasha, fill="blue", font=self._arial_bold_font, anchor="mm")

                # Affiche l’heure de 'כניסת שבת' en haut (440px)
                draw.text((time_x, 440), candle_lighting, fill="black", font=font)

                # Affiche la mincha_hol en vert à 950px
                draw.text((time_x, 950), self.format_time(times["mincha_hol"]), fill="green", font=font)

                # Affiche arvit_hol (milieu de semaine) en vert à 990px
                draw.text((time_x, 990), self.format_time(times["arvit_hol"]), fill="green", font=font)

                # Génération du nom de fichier safe
                safe_parasha = self.sanitize_filename(parasha)
                output_filename = f"horaires_{safe_parasha}.jpeg"
                output_path = self.output_dir / output_filename
                print(f"Chemin de sortie de l’image : {output_path}")
                img.save(str(output_path))
                print("Image sauvegardée avec succès")

                # Mise à jour du fichier latest-schedule.jpg
                latest_path = self.output_dir / "latest-schedule.jpg"
                if latest_path.exists():
                    latest_path.unlink()
                shutil.copy(str(output_path), str(latest_path))
                print(f"Copie vers le fichier le plus récent : {latest_path}")
                return output_path
        except Exception as e:
            print(f"❌ Erreur lors de la création de l’image: {e}")
            return None

    def update_excel(self, shabbat_data, times):
        """
        Met à jour le fichier Excel ou le crée s'il n'existe pas.
        """
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        next_shabbat_date, next_shabbat_time = self.get_next_shabbat_time(shabbat_data["date"])

        row = {
            "תאריך": shabbat_data["date"].strftime("%d/%m/%Y"),
            "פרשה": shabbat_data["parasha"],
            "API_parasha_hebrew": shabbat_data.get("parasha_hebrew", ""),
            "שיר השירים": self.format_time(times["shir_hashirim"]),
            "כניסת שבת": shabbat_data["candle_lighting"],
            "מנחה": self.format_time(times["mincha_kabbalat"]),
            "שחרית": self.format_time(times["shacharit"]),
            "מנחה אחרי צהריים": self.format_time(times["mincha_gdola"]),
            "תהילים": self.format_time(times["tehilim"]),
            "שיעור לנשים": self.format_time(times["shiur_nashim"]),
            "שיעור פרשה": self.format_time(times["parashat_hashavua"]),
            "שיעור עם הרב": self.format_time(times["shiur_rav"]),
            "מנחה 2": self.format_time(times["mincha_2"]),
            "ערבית מוצאי שבת": self.format_time(times["arvit_motsach"]),
            "ערבית חול": self.format_time(times["arvit_hol"]),
            "מנחה חול": self.format_time(times["mincha_hol"]),
            "מוצאי שבת Kodch": shabbat_data["end"].strftime("%H:%M"),
            " Sabbath suivant (Date)": next_shabbat_date if next_shabbat_date else "N/A",
            " Sabbath suivant (Heure)": next_shabbat_time if next_shabbat_time else "N/A",
            "שבת מברכין": "Oui" if shabbat_data.get("is_mevarchim", False) else "Non"
        }

        try:
            yearly_df = pd.DataFrame(self.yearly_shabbat_data)
            def compute_times(row):
                row_date = datetime.strptime(row["day"], "%Y-%m-%d %H:%M:%S").date()
                sunday_date = row_date + timedelta(days=2)
                s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
                sunday_sunset = s_sunday.get("sunset", None).strftime("%H:%M") if s_sunday.get("sunset") else ""
                thursday_date = sunday_date + timedelta(days=4)
                s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
                thursday_sunset = s_thursday.get("sunset", None).strftime("%H:%M") if s_thursday.get("sunset") else ""
                def to_minutes(t):
                    if t is None:
                        return None
                    try:
                        h, m = map(int, t.split(":"))
                        return h * 60 + m
                    except ValueError:
                        return None
                if sunday_sunset and thursday_sunset:
                    base = min(to_minutes(sunday_sunset), to_minutes(thursday_sunset)) - 17
                    if base < 0:
                        base = 0
                    mincha_hol = self.format_time(self.round_to_nearest_five(base))
                else:
                    mincha_hol = ""
                # Calcul arvit_hol (tsé hakochavim, milieu de semaine)
                sunday_dusk = s_sunday.get("dusk", None).strftime("%H:%M") if s_sunday.get("dusk") else ""
                thursday_dusk = s_thursday.get("dusk", None).strftime("%H:%M") if s_thursday.get("dusk") else ""
                def arvit_hol_time():
                    if sunday_dusk and thursday_dusk:
                        h1, m1 = map(int, sunday_dusk.split(":"))
                        h2, m2 = map(int, thursday_dusk.split(":"))
                        t1 = h1 * 60 + m1
                        t2 = h2 * 60 + m2
                        moyenne = (t1 + t2) // 2
                        tsé_min = min(t1, t2)
                        arvit_time = self.round_to_nearest_five(moyenne)
                        if tsé_min - arvit_time > 3:
                            arvit_time = ((tsé_min - 3 + 4) // 5) * 5
                        return self.format_time(arvit_time)
                    return ""
                return pd.Series({
                    "שקיעה Dimanche": sunday_sunset,
                    "שקיעה Jeudi": thursday_sunset,
                    "צאת הכוכבים Dimanche": sunday_dusk,
                    "צאת הכוכבים Jeudi": thursday_dusk,
                    "מנחה חול": mincha_hol,
                    "ערבית חול": arvit_hol_time()
                })
            times_df = yearly_df.apply(compute_times, axis=1)
            yearly_df = pd.concat([yearly_df, times_df], axis=1)

            if excel_path.exists():
                with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=writer.sheets["Sheet1"].max_row if "Sheet1" in writer.sheets else 0)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            else:
                with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name="Sheet1", index=False)
                    yearly_df.to_excel(writer, sheet_name="שבתות השנה", index=False)
            print(f"✅ Excel mis à jour: {excel_path}")
        except Exception as e:
            print(f"❌ Erreur lors de la mise à jour de l’Excel: {e}")

    def generate(self):
        """
        Génère les horaires et met à jour Excel.
        """
        current_date = datetime.now()
        end_date = current_date + timedelta(days=14)
        shabbat_times = self.get_hebcal_times(current_date, end_date)

        if not shabbat_times:
            print("Aucun horaire trouvé via l’API pour cette semaine.")
            print("Tentative de récupération depuis Excel...")
            excel_result = self.get_shabbat_times_from_excel_file(current_date)
            if not excel_result:
                print("❌ Aucun horaire trouvé pour cette semaine")
                return
            shabbat = excel_result[0]
        else:
            shabbat = shabbat_times[0]

        api_hebrew = shabbat.get("parasha_hebrew", "").strip()
        if not api_hebrew or api_hebrew == shabbat.get("parasha", "").strip():
            excel_result = self.get_shabbat_times_from_excel_file(current_date)
            if excel_result and excel_result[0].get("parasha_hebrew", "").strip():
                shabbat["parasha_hebrew"] = excel_result[0].get("parasha_hebrew", "").strip()
                print("Nom de parasha en hébreu récupéré depuis Excel pour vérification.")
            else:
                print("Aucune version hébraïque trouvée ; on conserve la valeur API.")
        times = self.calculate_times(shabbat['start'], shabbat['end'])
        image_path = self.create_image(
            times,
            shabbat["parasha"],
            shabbat["parasha_hebrew"],
            shabbat["end"],
            shabbat["candle_lighting"],
            shabbat["date"],
            is_mevarchim=shabbat.get("is_mevarchim", False)
        )
        if not image_path:
            print("❌ Échec de la génération de l’image")
        self.update_excel(shabbat, times)

def main():
    """
    Point d'entrée du script.
    """
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
        generator.update_excel_with_mevarchim_column(generator.output_dir / "horaires_shabbat.xlsx")
        generator.generate()
    except Exception as e:
        print(f"❌ Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")

if __name__ == "__main__":
    main()