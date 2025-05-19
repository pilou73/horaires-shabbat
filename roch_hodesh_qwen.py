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

        # Configuration de la localisation pour Ramat Gan, Israël
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

        # Données intégrées pour l'onglet "שבתות השנה"
        self.yearly_shabbat_data = [
            {'day': '2024-12-06 00:00:00', 'פרשה': 'ויצא', 'כנסתית שבת': '16:17', 'צאת שבת': '17:16'},
            {'day': '2024-12-13 00:00:00', 'פרשה': 'וישלח', 'כנסתית שבת': '16:19', 'צאת שבת': '17:17'},
            {'day': '2024-12-20 00:00:00', 'פרשה': 'וישב', 'כנסתית שבת': '16:22', 'צאת שבת': '17:20'},
            {'day': '2024-12-27 00:00:00', 'פרשה': 'מקץ', 'כנסתית שבת': '16:25', 'צאת שבת': '17:24'},
            # ... (autres données)
        ]

    def sanitize_filename(self, value: str) -> str:
        """Transforme une chaîne en slug ASCII-safe pour les noms de fichiers."""
        nfkd = unicodedata.normalize('NFKD', value)
        ascii_str = nfkd.encode('ascii', 'ignore').decode('ascii')
        ascii_str = re.sub(r'[^\w\s-]', '', ascii_str).strip()
        return re.sub(r'\s+', '_', ascii_str)

    def determine_season(self):
        today = datetime.now()
        year = today.year
        start_summer = datetime(year, 3, 29)
        end_summer = datetime(year, 10, 26)
        return "summer" if start_summer <= today <= end_summer else "winter"

    def fetch_roshchodesh_dates(self, start_date, end_date):
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
        if rosh_date.weekday() == 4:
            return rosh_date - timedelta(days=7)
        elif rosh_date.weekday() == 5:
            return rosh_date - timedelta(days=8)
        else:
            delta = (rosh_date.weekday() - 4 + 7) % 7
            return rosh_date - timedelta(days=delta)

    def identify_shabbat_mevarchim(self, shabbat_df, rosh_dates):
        shabbat_df = shabbat_df.copy()
        shabbat_df["day"] = pd.to_datetime(shabbat_df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        mevarchim_set = set()
        for rd in rosh_dates:
            mevarchim_friday = self.get_mevarchim_friday(rd)
            if mevarchim_friday < rd and mevarchim_friday in shabbat_df["day"].values:
                mevarchim_set.add(mevarchim_friday)
        shabbat_df["שבת מברכין"] = shabbat_df["day"].isin(mevarchim_set)
        return shabbat_df

    def get_shabbat_times_from_excel_file(self, current_date):
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        if excel_path.exists():
            try:
                df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
                df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
                today_date = current_date.date()
                df_filtered = df[df["day"] >= today_date].sort_values(by="day")
                if df_filtered.empty:
                    print("❌ Aucune donnée dans Excel pour une date supérieure à aujourd'hui")
                    return None
                row = df_filtered.iloc[0]
                shabbat_date = datetime.combine(row["day"], datetime.min.time())
                try:
                    candle_time = datetime.strptime(str(row["כנסתית שבת"]), "%H:%M").time()
                except Exception as e:
                    print("❌ Erreur lors de la lecture de l'heure 'כנסתית שבת':", e)
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
                    "candle_lighting": row["כנסתית שבת"],
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
            df_filtered = df[df["day"] >= today_date].sort_values(by="day")
            if df_filtered.empty:
                df_filtered = df.sort_values(by="day")
            if df_filtered.empty:
                print("❌ Aucune donnée disponible dans les données internes")
                return None
            row = df_filtered.iloc[0]
            shabbat_date = datetime.combine(row["day"], datetime.min.time())
            try:
                candle_time = datetime.strptime(str(row["כנסתית שבת"]), "%H:%M").time()
            except Exception as e:
                print("❌ Erreur lors de la lecture de l'heure 'כנסתית שבת':", e)
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
                "candle_lighting": row["כנסתית שבת"],
                "is_mevarchim": is_mevarchim_excel
            }]

    def get_hebcal_times(self, start_date, end_date):
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
                        min_date = start_date
                        max_date = end_date
                        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date)
                        shabbat_df = pd.DataFrame(self.yearly_shabbat_data)
                        shabbat_df["day"] = pd.to_datetime(shabbat_df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
                        shabbat_df = self.identify_shabbat_mevarchim(shabbat_df, rosh_dates)
                        is_mevarchim = False
                        match = shabbat_df[shabbat_df["day"] == start_time.date()]
                        if not match.empty:
                            is_mevarchim = match["שבת מברכין"].iloc[0]
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
        return (minutes // 5) * 5

    def format_time(self, minutes):
        if minutes is None or minutes < 0:
            return ""
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"

    def reverse_hebrew_text(self, text):
        return text

    def calculate_times(self, shabbat_start, shabbat_end):
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute

        times = {
            "mincha_kabbalat": start_minutes,
            "shir_hashirim": self.round_to_nearest_five(start_minutes - 10),
            "shacharit": self.round_to_nearest_five(7 * 60 + 45),
            "mincha_gdola": self.round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": self.round_to_nearest_five(14 * 60),
            "shiur_nashim": 16 * 60,
            "mincha_2": None,
            "shiur_rav": None,
            "parashat_hashavua": None,
            "mincha_hol": None,
            "arvit_midweek": None,
            "arvit_motsach": None
        }

        if start_minutes - times["shir_hashirim"] < 10:
            times["shir_hashirim"] = start_minutes - 10

        times["mincha_2"] = self.round_to_nearest_five(end_minutes - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        times["parashat_hashavua"] = self.round_to_nearest_five(times["shiur_rav"] - 45)

        sunday_date = shabbat_start.date() + timedelta(days=2)
        s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
        sunday_dusk = s_sunday.get("dusk", None)
        thursday_date = sunday_date + timedelta(days=4)
        s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
        thursday_dusk = s_thursday.get("dusk", None)

        def to_minutes(t):
            if t is None:
                return None
            try:
                h, m = map(int, t.split(":"))
                return h * 60 + m
            except ValueError:
                return None

        tsé_dimanche = to_minutes(sunday_dusk.strftime("%H:%M")) if sunday_dusk else None
        tsé_jeudi = to_minutes(thursday_dusk.strftime("%H:%M")) if thursday_dusk else None
        if tsé_dimanche and tsé_jeudi:
            moyenne = (tsé_dimanche + tsé_jeudi) // 2
            tsé_min = min(tsé_dimanche, tsé_jeudi)
            new_arvit_time = self.round_to_nearest_five(moyenne)
            if tsé_min - new_arvit_time > 3:
                new_arvit_time = ((tsé_min - 3 + 4) // 5) * 5
            times["arvit_midweek"] = new_arvit_time
        else:
            times["arvit_midweek"] = 0

        sunday_sunset = s_sunday.get("sunset", None)
        thursday_sunset = s_thursday.get("sunset", None)
        if sunday_sunset and thursday_sunset:
            base = min(to_minutes(sunday_sunset.strftime("%H:%M")), to_minutes(thursday_sunset.strftime("%H:%M"))) - 17
            if base < 0:
                base = 0
            times["mincha_hol"] = self.round_to_nearest_five(base)
        else:
            times["mincha_hol"] = 0

        times["arvit_motsach"] = self.round_to_nearest_five(end_minutes - 5)

        return times

    def get_next_shabbat_time(self, current_shabbat_date):
        try:
            if isinstance(current_shabbat_date, datetime):
                current_date = current_shabbat_date.date()
            else:
                current_date = current_shabbat_date
            change_time_date = datetime(2025, 3, 27).date()
            last_shabbat = None
            for shabbat in self.yearly_shabbat_data:
                shabbat_date = datetime.strptime(shabbat["day"], "%Y-%m-%d %H:%M:%S").date()
                if shabbat_date > current_date:
                    if shabbat_date > change_time_date and last_shabbat:
                        shabbat = last_shabbat
                    shabbat_entry_time = shabbat["כנסתית שבת"]
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
        try:
            print("Ouverture du template...")
            print(f"Chemin du template : {self.template_path}")
            print(f"Chemin de la police : {self.font_path}")
            print(f"Chemin de la police Arial Bold : {self.arial_bold_path}")

            with Image.open(self.template_path) as img:
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(str(self.font_path), 30)
                time_x = 120

                time_positions_black = [
                    (time_x, 400, 'shir_hashirim'),
                    (time_x, 475, 'mincha_kabbalat'),
                    (time_x, 510, 'shacharit'),
                    (time_x, 550, 'mincha_gdola'),
                    (time_x, 590, 'tehilim'),
                    (time_x, 630, 'shiur_nashim'),
                    (time_x, 670, 'parashat_hashavua'),
                    (time_x, 710, 'shiur_rav'),
                    (time_x, 750, 'mincha_2'),
                    (time_x, 830, 'arvit_motsach'),
                ]

                time_positions_green = [
                    (time_x, 790, 'arvit_midweek'),
                    (time_x, 950, 'mincha_hol'),
                ]

                for x, y, key in time_positions_black:
                    if key in times:
                        draw.text((x, y), self.format_time(times[key]), fill="black", font=font)
                    else:
                        draw.text((x, y), "", fill="black", font=font)

                for x, y, key in time_positions_green:
                    draw.text((x, y), self.format_time(times.get(key, "")), fill="green", font=font)

                end_time_str = shabbat_end.strftime("%H:%M")
                draw.text((time_x, 910), end_time_str, fill="black", font=font)

                reversed_parasha = self.reverse_hebrew_text(parasha_hebrew)
                draw.text((300, 280), reversed_parasha, fill="blue", font=self._arial_bold_font, anchor="mm")

                draw.text((time_x, 440), candle_lighting, fill="black", font=font)

                safe_parasha = self.sanitize_filename(parasha)
                output_filename = f"horaires_{safe_parasha}.jpeg"
                output_path = self.output_dir / output_filename
                print(f"Chemin de sortie de l'image : {output_path}")
                img.save(str(output_path))
                print("Image sauvegardée avec succès")

                latest_path = self.output_dir / "latest-schedule.jpg"
                if latest_path.exists():
                    latest_path.unlink()
                shutil.copy(str(output_path), str(latest_path))
                print(f"Copie vers le fichier le plus récent : {latest_path}")
                return output_path
        except Exception as e:
            print(f"❌ Erreur lors de la création de l'image: {e}")
            return None

    def update_excel_with_mevarchim_column(self, excel_path: Path):
        if not excel_path.exists():
            print("Fichier Excel non trouvé, création avec les données intégrées")
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        else:
            df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date

        min_date = df["day"].min()
        max_date = df["day"].max()
        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date)
        df = self.identify_shabbat_mevarchim(df, rosh_dates)

        with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name="שבתות השנה", index=False)
        print("✅ Colonne 'שבת מברכין' mise à jour dans Excel.")

    def update_excel(self, shabbat_data, times):
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
            "מוצאי שבת Kodch": shabbat_data["end"].strftime("%H:%M"),
            " Sabbath suivant (Date)": next_shabbat_date if next_shabbat_date else "N/A",
            " Sabbath suivant (Heure)": next_shabbat_time if next_shabbat_time else "N/A",
            "שבת מברכין": "Oui" if shabbat_data.get("is_mevarchim", False) else "Non",
            "mincha_hol": self.format_time(times["mincha_hol"]),
            "arvit_midweek": self.format_time(times["arvit_midweek"]),
            "arvit_motsach": self.format_time(times["arvit_motsach"])
        }
        try:
            yearly_df = pd.DataFrame(self.yearly_shabbat_data)
            yearly_df["day"] = pd.to_datetime(yearly_df["day"], format="%Y-%m-%d %H:%M:%S").dt.date

            def compute_times(row):
                row_date = datetime.strptime(row["day"], "%Y-%m-%d %H:%M:%S").date()
                sunday_date = row_date + timedelta(days=2)
                s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
                sunday_sunset = s_sunday.get("sunset", None)
                if sunday_sunset:
                    sunday_sunset = sunday_sunset.strftime("%H:%M")
                else:
                    sunday_sunset = ""
                thursday_date = sunday_date + timedelta(days=4)
                s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
                thursday_sunset = s_thursday.get("sunset", None)
                if thursday_sunset:
                    thursday_sunset = thursday_sunset.strftime("%H:%M")
                else:
                    thursday_sunset = ""

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
                    mincha_midweek = self.format_time(self.round_to_nearest_five(base))
                else:
                    mincha_midweek = ""
                return pd.Series({
                    "שקיעה Dimanche": sunday_sunset,
                    "שקיעה Jeudi": thursday_sunset,
                    "צאת הכוכבים Dimanche": s_sunday.get("dusk", None).strftime("%H:%M") if s_sunday.get("dusk") else "",
                    "צאת הכוכבים Jeudi": s_thursday.get("dusk", None).strftime("%H:%M") if s_thursday.get("dusk") else "",
                    "מנחה ביניים": mincha_midweek
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
            print(f"❌ Erreur lors de la mise à jour de l'Excel: {e}")

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
        generator.update_excel_with_mevarchim_column(generator.output_dir / "horaires_shabbat.xlsx")
        generator.generate()
    except Exception as e:
        print(f"❌ Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")

if __name__ == "__main__":
    main()