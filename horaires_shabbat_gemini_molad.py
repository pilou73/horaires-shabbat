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
from zmanim.hebrew_calendar.jewish_calendar import JewishCalendar

# ---- UTILS MOLAD & HEBREU ----
HEBREW_MONTHS = {
    1: 'ניסן', 2: 'אייר', 3: 'סיון', 4: 'תמוז',
    5: 'אב', 6: 'אלול', 7: 'תשרי', 8: 'חשוון',
    9: 'כסלו', 10: 'טבת', 11: 'שבט', 12: 'אדר',
    13: 'אדר ב׳'
}
HEBREW_DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת']

def get_jewish_month_name_hebrew(jm, jy):
    if jm == 12 and JewishCalendar.is_jewish_leap_year(jy):
        return 'אדר א׳'
    if jm == 13:
        return 'אדר ב׳'
    return HEBREW_MONTHS.get(jm, 'חודש לא ידוע')

def find_previous_rosh_chodesh(date_):
    current_date = date_
    for i in range(30):  # Recherche max sur 30 jours en arrière
        cal = JewishCalendar(current_date)
        if cal.jewish_day == 1:
            return current_date
        current_date -= timedelta(days=1)
    raise Exception("Aucun Rosh Khodesh trouvé dans les 30 jours précédents")

def reverse_hebrew_text(text):
    return text[::-1]

def get_weekday_name_hebrew(dt):
    return HEBREW_DAYS[(dt.weekday() + 1) % 7]

def get_next_month_molad(shabbat_date):
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
    hebrew_part = f"מולד: יום {weekday_he} בשעה "
    molad_str = hebrew_part + f"{hour}:{str(minute).zfill(2)} + {chalakim}"
    return molad_str

def get_rosh_chodesh_days_for_next_month(shabbat_date):
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
        raise Exception("Impossible de déterminer le nombre de jours dans le mois hébraïque.")

    jc_next = JewishCalendar()
    jc_next.set_jewish_date(next_jyear, next_jmonth, 1)
    rc1_gdate = jc_next.gregorian_date

    rosh_chodesh_days = []
    if days_in_current_month == 30:
        jc_current.set_jewish_date(current_jyear, current_jmonth, 30)
        rc0_gdate = jc_current.gregorian_date
        rosh_chodesh_days.append((rc0_gdate, current_jmonth, current_jyear, 30))
    rosh_chodesh_days.append((rc1_gdate, next_jmonth, next_jyear, 1))
    return rosh_chodesh_days

def calculate_last_kiddush_levana_date(gregorian_date):
    jc = JewishCalendar(datetime.combine(gregorian_date, datetime.min.time()))
    molad_obj = jc.molad()
    molad_date = datetime.combine(gregorian_date, datetime.min.time())
    molad_dt = molad_date + timedelta(
        hours=molad_obj.molad_hours,
        minutes=molad_obj.molad_minutes,
        seconds=molad_obj.molad_chalakim * 10 / 18
    )
    latest_time = molad_dt + timedelta(days=12, hours=18)
    return molad_dt, latest_time

def find_next_rosh_chodesh(date_):
    current_date = date_
    while True:
        cal = JewishCalendar(current_date)
        if cal.jewish_day == 1:
            return current_date
        current_date += timedelta(days=1)
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

        self.yearly_shabbat_data = [
            # ... (ta liste des shabbat de l'année, comme dans la partie précédente)
        ]

        # Moladot 5785 exacts (Jérusalem, UTC+2)
        self.moladot_5785 = [
            {"molad_date": "2024-10-03 03:21"},
            {"molad_date": "2024-11-01 16:05"},
            {"molad_date": "2024-12-01 04:49"},
            {"molad_date": "2024-12-30 17:33"},
            {"molad_date": "2025-01-29 06:17"},
            {"molad_date": "2025-02-28 19:02"},
            {"molad_date": "2025-03-29 07:46"},
            {"molad_date": "2025-04-28 20:30"},
            {"molad_date": "2025-05-27 09:14"},
            {"molad_date": "2025-06-26 21:58"},
            {"molad_date": "2025-07-25 10:42"},
            {"molad_date": "2025-08-24 23:26"},
        ]
        for m in self.moladot_5785:
            m["datetime"] = datetime.strptime(m["molad_date"], "%Y-%m-%d %H:%M")

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
        df = df.sort_values("day").reset_index(drop=True)
        df["day"] = pd.to_datetime(df["day"]).dt.date
        df["מולד"] = ""
        for idx, molad in enumerate(self.moladot_5785):
            dt_start = molad["datetime"].date()
            dt_end = self.moladot_5785[idx+1]["datetime"].date() if idx+1 < len(self.moladot_5785) else df["day"].max() + timedelta(days=7)
            mask = (df["day"] >= dt_start) & (df["day"] < dt_end)
            df.loc[mask, "מולד"] = molad["datetime"].strftime("%Y-%m-%d %H:%M")
        return df
    def update_excel_with_mevarchim_column(self, excel_path: Path):
        if not excel_path.exists():
            print("Fichier Excel non trouvé, création avec les données intégrées")
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
        with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name="שבתות השנה", index=False)
        print("✅ Colonnes שבת מברכין et מולד mises à jour dans Excel.")

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
                    "is_mevarchim": row.get("שבת מברכין", False) == True or row.get("שבת מברכין", "") == "Oui",
                    "molad": row.get("מולד", ""),
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
            df = self.fill_molad_column(df)
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
                "is_mevarchim": row.get("שבת מברכין", False) == True or row.get("שבת מברכין", "") == "Oui",
                "molad": row.get("מולד", ""),
            }]
    def round_to_nearest_five(self, minutes):
        return (minutes // 5) * 5 if minutes is not None else None

    def round_to_next_five(self, minutes):
        return ((minutes + 4) // 5) * 5 if minutes is not None else None

    def format_time(self, minutes):
        if minutes is None or minutes == "":
            return ""
        if isinstance(minutes, str) and ":" in minutes:
            return minutes
        if minutes < 0:
            return ""
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"

    def calculate_times(self, shabbat_start, shabbat_end):
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute
        tehilim_ete = self.round_to_nearest_five(17 * 60)
        tehilim_hiver = self.round_to_nearest_five(14 * 60)
        tehilim = tehilim_ete if self.season == "summer" else tehilim_hiver

        times = {
            "mincha_kabbalat": start_minutes,
            "shir_hashirim": self.round_to_nearest_five(start_minutes - 10),
            "shacharit": self.round_to_nearest_five(7 * 60 + 45),
            "mincha_gdola": self.round_to_nearest_five(12 * 60 + (30 if self.season == "winter" else 60)),
            "tehilim": tehilim,
            "tehilim_ete": tehilim_ete,
            "tehilim_hiver": tehilim_hiver,
            "shiur_nashim": 16 * 60 + 15,
            "arvit_hol": None,
            "arvit_motsach": None,
            "mincha_2": None,
            "shiur_rav": None,
            "parashat_hashavua": None,
            "mincha_hol": None
        }

        times["mincha_2"] = self.round_to_nearest_five(end_minutes - 90)
        times["shiur_rav"] = self.round_to_nearest_five(times["mincha_2"] - 45)
        times["parashat_hashavua"] = self.round_to_nearest_five(times["shiur_rav"] - 45)

        # Calcul minha/arvit חול
        sunday_date = shabbat_start.date() + timedelta(days=2)
        s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
        sunday_sunset = s_sunday["sunset"].strftime("%H:%M")
        thursday_date = sunday_date + timedelta(days=4)
        s_thu = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
        thursday_sunset = s_thu["sunset"].strftime("%H:%M")

        def to_minutes(t):
            h, m = map(int, t.split(":"))
            return h * 60 + m

        if sunday_sunset and thursday_sunset:
            sunday_sunset_min = to_minutes(sunday_sunset)
            thursday_sunset_min = to_minutes(thursday_sunset)
            min_sunset = min(sunday_sunset_min, thursday_sunset_min)
            max_sunset = max(sunday_sunset_min, thursday_sunset_min)
            times["mincha_hol"] = self.round_to_nearest_five(min_sunset - 18)
            times["arvit_hol"] = self.round_to_next_five(max_sunset + 20)
        else:
            times["mincha_hol"] = ""
            times["arvit_hol"] = ""

        times["arvit_motsach"] = self.round_to_nearest_five(end_minutes - 9)

        return times

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting, shabbat_date, is_mevarchim=False):
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

                # Affichage des horaires principaux
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
                        if key == 'tehilim':
                            if self.season == "summer":
                                formatted_time = f"{self.format_time(times['tehilim_ete'])}/{self.format_time(times['tehilim_hiver'])}"
                                draw.text((x - 50, y), formatted_time, fill="black", font=font)
                            else:
                                draw.text((x, y), self.format_time(times['tehilim']), fill="black", font=font)
                        else:
                            draw.text((x, y), self.format_time(times[key]), fill="black", font=font)

                draw.text((time_x, 440), candle_lighting, fill="black", font=font)
                draw.text((time_x, 830), shabbat_end.strftime("%H:%M"), fill="black", font=font)
                draw.text((time_x, 950), self.format_time(times.get('mincha_hol')), fill="green", font=font)
                draw.text((time_x, 990), self.format_time(times.get('arvit_hol')), fill="green", font=font)
                draw.text((300, 280), parasha_hebrew, fill="blue", font=bold, anchor="mm")

                # שבת מברכין et infos rosh hodesh / molad
                if is_mevarchim:
                    molad_str = get_next_month_molad(shabbat_date)
                    draw.text((200, img_h - 300), molad_str, fill="blue", font=font)
                    rc_days = get_rosh_chodesh_days_for_next_month(shabbat_date)
                    for i, (gdate, m, y, d) in enumerate(rc_days):
                        day_name_he = get_weekday_name_hebrew(gdate)
                        month_name = get_jewish_month_name_hebrew(m, y)
                        rc_line = f"ראש חודש: יום {day_name_he} {gdate.strftime('%d/%m/%Y')} {month_name} ({d})"
                        draw.text((200, img_h - 260 + 40 * i), rc_line, fill="blue", font=font)
                else:
                    try:
                        previous_rosh = find_previous_rosh_chodesh(shabbat_date)
                        molad_dt, latest_kiddush_levana = calculate_last_kiddush_levana_date(previous_rosh)
                        start_kiddush_levana = molad_dt + timedelta(days=6)

                        shabbat_date_only = shabbat_date.date()

                        if shabbat_date_only < start_kiddush_levana.date():
                            msg_start = f"זמן התחלה לאמירת ברכת הלבנה: {start_kiddush_levana.strftime('%d/%m/%Y')}"
                            msg_end = f"תאריך אחרון לאמירת ברכת הלבנה: {latest_kiddush_levana.strftime('%d/%m/%Y')}"
                            draw.text((100, img_h - 300), msg_start, fill="blue", font=font)
                            draw.text((100, img_h - 260), msg_end, fill="blue", font=font)
                        elif start_kiddush_levana.date() <= shabbat_date_only <= latest_kiddush_levana.date():
                            msg_end = f"תאריך אחרון לאמירת ברכת הלבנה: {latest_kiddush_levana.strftime('%d/%m/%Y')}"
                            draw.text((100, img_h - 260), msg_end, fill="blue", font=font)
                        else:
                            msg_ended = "התקופה של ברכת הלבנה הסתיימה."
                            draw.text((100, img_h - 260), msg_ended, fill="red", font=font)
                    except Exception as e:
                        print(f"❌ Erreur lors de l'affichage de la Birkat Halevana : {e}")

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
        row = {
            "day": shabbat_data["date"],
            "תאריך": shabbat_data["date"].strftime("%d/%m/%Y"),
            "פרשה": shabbat_data["parasha"],
            "API_parasha_hebrew": shabbat_data.get("parasha_hebrew", ""),
            "שיר השירים": self.format_time(times["shir_hashirim"]),
            "כניסת שבת": shabbat_data["candle_lighting"],
            "מנחה": self.format_time(times["mincha_kabbalat"]),
            "שחרית": self.format_time(times["shacharit"]),
            "מנחה גדולה": self.format_time(times["mincha_gdola"]),
            "תהילים קיץ": self.format_time(times["tehilim_ete"]),
            "תהילים חורף": self.format_time(times["tehilim_hiver"]),
            "שיעור לנשים": self.format_time(times["shiur_nashim"]),
            "שיעור פרשה": self.format_time(times["parashat_hashavua"]),
            "שיעור עם הרב": self.format_time(times["shiur_rav"]),
            "מנחה 2": self.format_time(times["mincha_2"]),
            "ערבית מוצאי שבת": self.format_time(times["arvit_motsach"]),
            "ערבית חול": self.format_time(times["arvit_hol"]),
            "מנחה חול": self.format_time(times["mincha_hol"]),
            "מוצאי שבת קודש": shabbat_data["end"].strftime("%H:%M"),
            "שבת מברכין": "Oui" if shabbat_data.get("is_mevarchim", False) else "Non",
            "מולד": shabbat_data.get("molad", "")
        }
        try:
            df_sheet1 = pd.DataFrame([row])
            with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="a" if excel_path.exists() else "w") as writer:
                df_sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
            print(f"✅ Onglet Sheet1 mis à jour dans Excel: {excel_path}")
        except Exception as e:
            print(f"❌ Erreur lors de la mise à jour de l’onglet Sheet1: {e}")

    def generate(self):
        current_date = datetime.now()
        shabbat_times = self.get_shabbat_times_from_excel_file(current_date)
        if not shabbat_times:
            print("❌ Aucun horaire trouvé pour cette semaine")
            return
        shabbat = shabbat_times[0]
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