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

def reverse_hebrew_text(text):
    return text[::-1]

def round_to_nearest_five(minutes):
    return (minutes // 5) * 5 if minutes is not None else None

def round_to_next_five(minutes):
    return ((minutes + 4) // 5) * 5 if minutes is not None else None

def format_time_hhmm(minutes):
    if minutes is None or minutes == "":
        return ""
    if isinstance(minutes, str) and ":" in minutes:
        return minutes
    if minutes < 0:
        return ""
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"

class ShabbatScheduleGenerator:
    def __init__(self, template_path, font_path, arial_bold_path, output_dir):
        self.template_path = Path(template_path)
        self.font_path = Path(font_path)
        self.arial_bold_path = Path(arial_bold_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Template introuvable: {self.template_path}")
        if not self.font_path.exists():
            raise FileNotFoundError(f"Police introuvable: {self.font_path}")
        if not self.arial_bold_path.exists():
            raise FileNotFoundError(f"Police Arial Bold introuvable: {self.arial_bold_path}")

        self._font = ImageFont.truetype(str(self.font_path), 30)
        self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)

        self.season = self.determine_season()
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

        # Données exactes fournies
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
    def determine_season(self):
        today = datetime.now()
        year = today.year
        start_summer = datetime(year, 3, 29)
        end_summer = datetime(year, 10, 26)
        return "summer" if start_summer <= today <= end_summer else "winter"

    def sanitize_filename(self, value: str) -> str:
        nfkd = unicodedata.normalize('NFKD', value)
        ascii_str = nfkd.encode('ascii', 'ignore').decode('ascii')
        ascii_str = re.sub(r'[^\w\s-]', '', ascii_str).strip()
        return re.sub(r'\s+', '_', ascii_str)

    def fetch_roshchodesh_dates(self, start_date, end_date):
        url = "https://www.hebcal.com/hebcal"
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
            seen_dates = set()
            for item in data.get("items", []):
                if item.get("category") == "roshchodesh":
                    dt = datetime.fromisoformat(item["date"]).astimezone(pytz.timezone("Asia/Jerusalem")).date()
                    if dt not in seen_dates:
                        rosh_dates.append(dt)
                        seen_dates.add(dt)
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
        if "day" not in shabbat_df.columns and "תאריך" in shabbat_df.columns:
            shabbat_df["day"] = pd.to_datetime(shabbat_df["תאריך"], format="%d/%m/%Y").dt.date
        shabbat_df["day"] = pd.to_datetime(shabbat_df["day"], format="%Y-%m-%d %H:%M:%S", errors='coerce').dt.date.fillna(shabbat_df["day"])
        mevarchim_set = set()
        for rd in rosh_dates:
            mevarchim_friday = self.get_mevarchim_friday(rd)
            if mevarchim_friday < rd and mevarchim_friday in shabbat_df["day"].values:
                mevarchim_set.add(mevarchim_friday)
        shabbat_df["שבת מברכין"] = shabbat_df["day"].isin(mevarchim_set)
        return shabbat_df

    def fill_molad_column(self, df):
        # Moladot 5785 exacts (Jérusalem, UTC+2)
        moladot_5785 = [
            {"molad_date": "2024-10-02 21:57"}, # Tishrei
            {"molad_date": "2024-11-01 08:25"}, # Cheshvan
            {"molad_date": "2024-11-30 18:54"}, # Kislev
            {"molad_date": "2024-12-30 05:08"}, # Tevet
            {"molad_date": "2025-01-28 13:10"}, # Shevat
            {"molad_date": "2025-02-26 20:01"}, # Adar I
            {"molad_date": "2025-03-28 02:43"}, # Adar II
            {"molad_date": "2025-04-26 08:19"}, # Nisan
            {"molad_date": "2025-05-25 13:01"}, # Iyar
            {"molad_date": "2025-06-24 16:59"}, # Sivan
            {"molad_date": "2025-07-24 20:26"}, # Tamuz
            {"molad_date": "2025-08-23 23:21"}, # Av
        ]
        for m in moladot_5785:
            m["datetime"] = datetime.strptime(m["molad_date"], "%Y-%m-%d %H:%M")
        df = df.sort_values("day").reset_index(drop=True)
        df["day"] = pd.to_datetime(df["day"]).dt.date
        df["מולד"] = ""
        for idx, molad in enumerate(moladot_5785):
            dt_start = molad["datetime"].date()
            dt_end = moladot_5785[idx+1]["datetime"].date() if idx+1 < len(moladot_5785) else df["day"].max() + timedelta(days=7)
            mask = (df["day"] >= dt_start) & (df["day"] < dt_end)
            df.loc[mask, "מולד"] = molad["datetime"].strftime("%Y-%m-%d %H:%M")
        return df

    def sunset_for_date(self, dateobj):
        s = sun(self.ramat_gan.observer, date=dateobj, tzinfo=self.ramat_gan.timezone)
        sunset = s.get("sunset")
        if sunset:
            return sunset.strftime("%H:%M")
        return ""

    def sunset_minutes(self, dateobj):
        s = sun(self.ramat_gan.observer, date=dateobj, tzinfo=self.ramat_gan.timezone)
        sunset = s.get("sunset")
        if sunset:
            return sunset.hour * 60 + sunset.minute
        return None

    def birkat_halevana_text(self, shabbat_date, molad_str):
        # Affiche la période de Birkat Halevana sur l'image si pertinent
        if not molad_str:
            return ""
        molad_dt = datetime.strptime(molad_str, "%Y-%m-%d %H:%M")
        birkat_start = molad_dt + timedelta(days=3, hours=8)  # 3 jours, 8h après le molad
        birkat_end = molad_dt + timedelta(days=14, hours=18)  # 14 jours, 18h après le molad
        shabbat_end_dt = shabbat_date + timedelta(hours=25)   # fin de shabbat + 1h (approximatif)
        if birkat_start.date() <= shabbat_end_dt.date() <= birkat_end.date():
            text = f"בין {birkat_start.strftime('%d/%m/%Y %H:%M')} ל-{birkat_end.strftime('%d/%m/%Y %H:%M')}"
            return "מברכים את הלבנה: " + text
        return ""
    def update_excel_with_mevarchim_column(self, excel_path: Path):
        # Crée ou MAJ שבתות השנה avec toutes les colonnes demandées
        if not excel_path.exists():
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        else:
            df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
            if "day" not in df.columns and "תאריך" in df.columns:
                df["day"] = pd.to_datetime(df["תאריך"], format="%d/%m/%Y").dt.date
            else:
                df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S", errors='coerce').dt.date.fillna(df["day"])
        min_date = df["day"].min()
        max_date = df["day"].max()
        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date + timedelta(days=7))
        df = self.identify_shabbat_mevarchim(df, rosh_dates)
        df = self.fill_molad_column(df)

        # Calcul שקיעה dimanche et jeudi, minha/arvit pour chaque ligne
        sunday_sunset_list = []
        thursday_sunset_list = []
        minha_ben_list = []
        arvit_ben_list = []
        for row in df.itertuples():
            sunday = row.day + timedelta(days=(6-row.day.weekday()+1)%7 + 1)  # dimanche après le shabbat
            thursday = row.day + timedelta(days=(6-row.day.weekday()+5)%7 + 5) # jeudi après le shabbat
            sunset_sunday = self.sunset_for_date(sunday)
            sunset_thursday = self.sunset_for_date(thursday)
            sunday_min = self.sunset_minutes(sunday)
            thursday_min = self.sunset_minutes(thursday)
            if sunday_min is None or thursday_min is None:
                minha_ben = ""
                arvit_ben = ""
            else:
                minha_ben = format_time_hhmm(round_to_nearest_five(min(sunday_min, thursday_min) - 18))
                arvit_ben = format_time_hhmm(round_to_next_five(max(sunday_min, thursday_min) + 20))
            sunday_sunset_list.append(sunset_sunday)
            thursday_sunset_list.append(sunset_thursday)
            minha_ben_list.append(minha_ben)
            arvit_ben_list.append(arvit_ben)
        df["שקיעה Dimanche"] = sunday_sunset_list
        df["שקיעה Jeudi"] = thursday_sunset_list
        df["מנחה ביניים"] = minha_ben_list
        df["ערבית ביניים"] = arvit_ben_list

        # Formattage final des colonnes demandées
        df["תאריך"] = pd.to_datetime(df["day"]).dt.strftime("%d/%m/%Y")
        columns_final = ["day", "תאריך", "פרשה", "כנסית שבת", "צאת שבת",
                         "שקיעה Dimanche", "שקיעה Jeudi",
                         "מנחה ביניים", "ערבית ביניים",
                         "שבת מברכין", "מולד"]
        df_final = df[columns_final]

        with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="w") as writer:
            df_final.to_excel(writer, sheet_name="שבתות השנה", index=False)
        print("✅ Colonnes שבת מברכין, מולד, שקיעות et horaires intermédiaires OK dans Excel.")

    def get_shabbat_times_from_excel_file(self, current_date):
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        if excel_path.exists():
            try:
                df = pd.read_excel(excel_path, sheet_name="שבתות השנה")
                if "day" not in df.columns and "תאריך" in df.columns:
                    df["day"] = pd.to_datetime(df["תאריך"], format="%d/%m/%Y").dt.date
                else:
                    df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S", errors='coerce').dt.date.fillna(df["day"])
                today_date = current_date.date()
                df = df[df["day"] >= today_date].sort_values(by="day")
                if df.empty:
                    return None
                row = df.iloc[0]
                shabbat_date = datetime.combine(row["day"], datetime.min.time())
                candle_time = datetime.strptime(str(row["כנסית שבת"]), "%H:%M").time()
                havdalah_time = datetime.strptime(str(row["צאת שבת"]), "%H:%M").time()
                shabbat_start = datetime.combine(row["day"], candle_time)
                shabbat_end = datetime.combine(row["day"], havdalah_time)
                return [{
                    "date": shabbat_date,
                    "start": shabbat_start,
                    "end": shabbat_end,
                    "parasha": row.get("פרשה", ""),
                    "parasha_hebrew": row.get("פרשה_עברית", row.get("פרשה", "")),
                    "candle_lighting": row["כנסית שבת"],
                    "is_mevarchim": row.get("שבת מברכין", False),
                    "molad": row.get("מולד", ""),
                    "sunday_sunset": row.get("שקיעה Dimanche", ""),
                    "thursday_sunset": row.get("שקיעה Jeudi", ""),
                    "mincha_ben": row.get("מנחה ביניים", ""),
                    "arvit_ben": row.get("ערבית ביניים", "")
                }]
            except Exception as e:
                print(f"❌ Erreur lors de la lecture du fichier Excel: {e}")
                return None
        else:
            print("Fichier Excel non trouvé, utilisation des données intégrées")
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
            row = df.iloc[0]
            shabbat_date = datetime.combine(row["day"], datetime.min.time())
            candle_time = datetime.strptime(str(row["כנסית שבת"]), "%H:%M").time()
            havdalah_time = datetime.strptime(str(row["צאת שבת"]), "%H:%M").time()
            shabbat_start = datetime.combine(row["day"], candle_time)
            shabbat_end = datetime.combine(row["day"], havdalah_time)
            return [{
                "date": shabbat_date,
                "start": shabbat_start,
                "end": shabbat_end,
                "parasha": row.get("פרשה", ""),
                "parasha_hebrew": row.get("פרשה_עברית", row.get("פרשה", "")),
                "candle_lighting": row["כנסית שבת"],
                "is_mevarchim": False,
                "molad": "",
                "sunday_sunset": "",
                "thursday_sunset": "",
                "mincha_ben": "",
                "arvit_ben": ""
            }]
    def create_image(self, times, parasha, parasha_hebrew,
                     shabbat_end, candle_lighting, shabbat_date, is_mevarchim=False,
                     mincha_ben=None, arvit_ben=None, birkat_halevana_text=None):
        try:
            template = self.template_path
            if is_mevarchim:
                rc_template = self.template_path.parent / "template_rosh_hodesh.jpg"
                if rc_template.exists():
                    template = rc_template
            with Image.open(template) as img:
                img_w, img_h = img.size
                draw = ImageDraw.Draw(img)
                font = self._font
                bold = self._arial_bold_font
                time_x = 120

                # Affichage des horaires principaux sur l'image
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
                    (time_x, 790, 'arvit_motsach'),
                ]
                for x, y, key in time_positions:
                    value = times.get(key)
                    display = ""
                    if key == 'tehilim':
                        display = f"{self.format_time(times['tehilim_ete'])}/{self.format_time(times['tehilim_hiver'])}"
                    else:
                        display = self.format_time(value)
                    draw.text((x, y), display, fill="black", font=font)

                # Candle lighting and end of shabbat
                draw.text((time_x, 440), candle_lighting, fill="black", font=font)
                draw.text((time_x, 830), shabbat_end.strftime("%H:%M"), fill="black", font=font)
                # Horaires minha/arvit hol (intercalaires)
                if mincha_ben:
                    draw.text((time_x, 950), f"מנחה חול: {mincha_ben}", fill="green", font=font)
                if arvit_ben:
                    draw.text((time_x, 990), f"ערבית חול: {arvit_ben}", fill="green", font=font)
                # Nom de la parasha en haut
                draw.text((300, 280), parasha_hebrew, fill="blue", font=bold, anchor="mm")

                # Birkat halevana
                if birkat_halevana_text:
                    draw.text((img_w // 2, img_h - 120), birkat_halevana_text, fill="purple", font=font, anchor="mm")

                # Nom de l'image
                safe_parasha = self.sanitize_filename(parasha)
                output_filename = f"horaires_{safe_parasha}.jpeg"
                output_path = self.output_dir / output_filename
                img.save(str(output_path))
                latest = self.output_dir / "latest-schedule.jpg"
                if latest.exists():
                    latest.unlink()
                shutil.copy(str(output_path), str(latest))
                return output_path
        except Exception as e:
            print(f"❌ Erreur lors de la création de l'image: {e}")
            return None

    def update_excel(self, shabbat_data, times):
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        # Onglet Sheet1: toutes les infos détaillées
        row = {
            "day": shabbat_data["date"],
            "תאריך": shabbat_data["date"].strftime("%d/%m/%Y"),
            "פרשה": shabbat_data["parasha"],
            "API_parasha_hebrew": shabbat_data.get("parasha_hebrew", ""),
            "שיר השירים": self.format_time(times.get("shir_hashirim")),
            "כניסת שבת": shabbat_data["candle_lighting"],
            "מנחה": self.format_time(times.get("mincha_kabbalat")),
            "שחרית": self.format_time(times.get("shacharit")),
            "מנחה גדולה": self.format_time(times.get("mincha_gdola")),
            "תהילים קיץ": self.format_time(times.get("tehilim_ete")),
            "תהילים חורף": self.format_time(times.get("tehilim_hiver")),
            "שיעור לנשים": self.format_time(times.get("shiur_nashim")),
            "שיעור פרשה": self.format_time(times.get("parashat_hashavua")),
            "שיעור עם הרב": self.format_time(times.get("shiur_rav")),
            "מנחה 2": self.format_time(times.get("mincha_2")),
            "ערבית מוצאי שבת": self.format_time(times.get("arvit_motsach")),
            "ערבית חול": shabbat_data.get("arvit_ben", ""),
            "מנחה חול": shabbat_data.get("mincha_ben", ""),
            "מוצאי שבת קודש": shabbat_data["end"].strftime("%H:%M"),
            "שבת מברכין": "Oui" if shabbat_data.get("is_mevarchim", False) else "Non",
            "מולד": shabbat_data.get("molad", ""),
            "שקיעה Dimanche": shabbat_data.get("sunday_sunset", ""),
            "שקיעה Jeudi": shabbat_data.get("thursday_sunset", "")
        }
        try:
            df_sheet1 = pd.DataFrame([row])
            with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="a" if excel_path.exists() else "w") as writer:
                df_sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
            print(f"✅ Onglet Sheet1 mis à jour dans Excel: {excel_path}")
        except Exception as e:
            print(f"❌ Erreur lors de la mise à jour de l’onglet Sheet1: {e}")

    def format_time(self, minutes):
        return format_time_hhmm(minutes)

    def calculate_times(self, shabbat_start, shabbat_end):
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute
        tehilim_ete = round_to_nearest_five(17 * 60)
        tehilim_hiver = round_to_nearest_five(14 * 60)
        tehilim = tehilim_ete if self.season == "summer" else tehilim_hiver

        times = {
            "mincha_kabbalat": start_minutes,
            "shir_hashirim": round_to_nearest_five(start_minutes - 10),
            "shacharit": round_to_nearest_five(7 * 60 + 45),
            "mincha_gdola": round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": tehilim,
            "tehilim_ete": tehilim_ete,
            "tehilim_hiver": tehilim_hiver,
            "shiur_nashim": 16 * 60 + 15,
            "arvit_hol": None,
            "arvit_motsach": round_to_nearest_five(end_minutes - 9),
            "mincha_2": round_to_nearest_five(end_minutes - 90),
            "shiur_rav": round_to_nearest_five(end_minutes - 135),
            "parashat_hashavua": round_to_nearest_five(end_minutes - 180),
            "mincha_hol": None
        }
        return times

    def generate(self):
        current_date = datetime.now()
        shabbat_times = self.get_shabbat_times_from_excel_file(current_date)
        if not shabbat_times:
            print("❌ Aucun horaire trouvé pour cette semaine")
            return
        shabbat = shabbat_times[0]
        times = self.calculate_times(shabbat['start'], shabbat['end'])

        birkat_text = self.birkat_halevana_text(shabbat['date'], shabbat.get("molad", ""))

        image_path = self.create_image(
            times,
            shabbat["parasha"],
            shabbat.get("parasha_hebrew", shabbat["parasha"]),
            shabbat["end"],
            shabbat["candle_lighting"],
            shabbat["date"],
            is_mevarchim=shabbat.get("is_mevarchim", False),
            mincha_ben=shabbat.get("mincha_ben", ""),
            arvit_ben=shabbat.get("arvit_ben", ""),
            birkat_halevana_text=birkat_text
        )
        if not image_path:
            print("❌ Échec de la génération de l’image")

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
        font_path     = base_path / "resources" / "mriamc_0.ttf"
        arial_bold    = base_path / "resources" / "ARIALBD_0.TTF"
        output_dir    = base_path / "output"
        generator = ShabbatScheduleGenerator(
            template_path, font_path, arial_bold, output_dir
        )
        generator.update_excel_with_mevarchim_column(generator.output_dir / "horaires_shabbat.xlsx")
        generator.generate()
    except Exception as e:
        print(f"❌ Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")

if __name__ == "__main__":
    main()