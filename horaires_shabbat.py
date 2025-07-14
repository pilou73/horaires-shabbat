import os
import sys
from pathlib import Path
from datetime import datetime, timedelta, date
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

ICS_PATH = Path(__file__).parent / "resources" / "jeunes.ics"

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
    for i in range(30):
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
    hebrew_part = f" יום {weekday_he} בשעה "
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

def parse_tekufa_ics(filepath):
    tekufot = []
    current_event = {}
    in_event = False
    with open(filepath, encoding='utf8') as f:
        for line in f:
            line = line.strip()
            if line == "BEGIN:VEVENT":
                in_event = True
                current_event = {}
            elif line == "END:VEVENT":
                if "DTSTART" in current_event and "SUMMARY" in current_event:
                    dtstr = current_event["DTSTART"]
                    if dtstr.startswith("TZID=Asia/Jerusalem:"):
                        dtstr = dtstr.replace("TZID=Asia/Jerusalem:", "")
                    dt = datetime.strptime(dtstr, "%Y%m%dT%H%M%S")
                    summary = current_event["SUMMARY"]
                    tekufot.append((dt, summary))
                in_event = False
            elif in_event:
                if line.startswith("DTSTART"):
                    current_event["DTSTART"] = line.split(":", 1)[1]
                elif line.startswith("SUMMARY"):
                    current_event["SUMMARY"] = line.split(":", 1)[1]
    return tekufot

def parse_fast_ics(ics_path):
    if not ics_path.exists():
        print(f"⚠️ Fichier ICS non trouvé : {ics_path}")
        return []
    fasts = []
    with open(ics_path, encoding='utf-8') as f:
        lines = f.readlines()
    current = {}
    in_event = False
    for line in lines:
        line = line.strip()
        if line == "BEGIN:VEVENT":
            current = {}
            in_event = True
        elif line == "END:VEVENT":
            if "DTSTART" in current and "SUMMARY" in current:
                dt = datetime.strptime(current["DTSTART"], "%Y%m%dT%H%M%S")
                fasts.append({
                    "date": dt.date(),
                    "summary": current["SUMMARY"],
                    "start": dt,
                    "end": datetime.strptime(current["DTEND"], "%Y%m%dT%H%M%S") if "DTEND" in current else None
                })
            in_event = False
        elif in_event:
            if line.startswith("DTSTART"):
                current["DTSTART"] = line.split(":", 1)[1]
            elif line.startswith("DTEND"):
                current["DTEND"] = line.split(":", 1)[1]
            elif line.startswith("SUMMARY"):
                current["SUMMARY"] = line.split(":", 1)[1]
    return fasts

def get_fast_from_ics_for_week(shabbat_date, ics_path=ICS_PATH):
    all_fasts = parse_fast_ics(ics_path)
    week_start = (shabbat_date + timedelta(days=1)).date()  # dimanche
    week_end = (shabbat_date + timedelta(days=6)).date()    # vendredi
    return [
        {
            "date": f["date"],
            "nom": f["summary"],
            "start": f["start"].strftime("%H:%M"),
            "end": f["end"].strftime("%H:%M") if f["end"] else ""
        }
        for f in all_fasts
        if week_start <= f["date"] <= week_end
    ]


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
            {'day': '2025-07-11 00:00:00', 'פרשה': 'בלק', 'כנסית שבת': '19:31', 'צאת שבת': '20:31'},
            {'day': '2025-07-18 00:00:00', 'פרשה': 'פינחס', 'כנסית שבת': '19:28', 'צאת שבת': '20:27'},
            {'day': '2025-07-25 00:00:00', 'פרשה': 'מטות-מסעי', 'כנסית שבת': '19:24', 'צאת שבת': '20:23'},
            {'day': '2025-08-01 00:00:00', 'פרשה': 'דברים', 'כנסית שבת': '19:19', 'צאת שבת': '20:17'},
            {'day': '2025-08-08 00:00:00', 'פרשה': 'ואתחנן', 'כנסית שבת': '19:13', 'צאת שבת': '20:10'},
            {'day': '2025-08-15 00:00:00', 'פרשה': 'עקב', 'כנסית שבת': '19:06', 'צאת שבת': '20:02'},
            {'day': '2025-08-22 00:00:00', 'פרשה': 'ראה', 'כנסית שבת': '18:59', 'צאת שבת': '19:54'},
            {'day': '2025-08-29 00:00:00', 'פרשה': 'שופטים', 'כנסית שבת': '18:50', 'צאת שבת': '19:45'},
{'day': '2025-09-05 18:41:00', 'פרשה': 'כי־תצא', 'כנסית שבת': '18:41', 'צאת שבת': '19:35'},
{'day': '2025-09-12 18:32:00', 'פרשה': 'כי־תבוא', 'כנסית שבת': '18:32', 'צאת שבת': '19:26'},
{'day': '2025-09-19 18:23:00', 'פרשה': 'נצבים', 'כנסית שבת': '18:23', 'צאת שבת': '19:16'},
{'day': '2025-09-26 18:14:00', 'פרשה': 'וילך', 'כנסית שבת': '18:14', 'צאת שבת': '19:07'},
{'day': '2025-10-03 18:05:00', 'פרשה': 'האזינו', 'כנסית שבת': '18:05', 'צאת שבת': '18:58'},
{'day': '2025-10-17 17:48:00', 'פרשה': 'בראשית', 'כנסית שבת': '17:48', 'צאת שבת': '18:42'},
{'day': '2025-10-24 17:40:00', 'פרשה': 'נח', 'כנסית שבת': '17:40', 'צאת שבת': '18:34'},
{'day': '2025-10-31 16:33:00', 'פרשה': 'לך־לך', 'כנסית שבת': '16:33', 'צאת שבת': '17:28'},
{'day': '2025-11-07 16:27:00', 'פרשה': 'וירא', 'כנסית שבת': '16:27', 'צאת שבת': '17:23'},
{'day': '2025-11-14 16:23:00', 'פרשה': 'חיי שרה', 'כנסית שבת': '16:23', 'צאת שבת': '17:19'},
{'day': '2025-11-21 16:20:00', 'פרשה': 'תולדות', 'כנסית שבת': '16:20', 'צאת שבת': '17:17'},
{'day': '2025-11-28 16:18:00', 'פרשה': 'ויצא', 'כנסית שבת': '16:18', 'צאת שבת': '17:15'},
{'day': '2025-12-05 16:18:00', 'פרשה': 'וישלח', 'כנסית שבת': '16:18', 'צאת שבת': '17:16'},
{'day': '2025-12-12 16:19:00', 'פרשה': 'וישב', 'כנסית שבת': '16:19', 'צאת שבת': '17:17'},
{'day': '2025-12-19 16:21:00', 'פרשה': 'מקץ', 'כנסית שבת': '16:21', 'צאת שבת': '17:20'},
{'day': '2025-12-26 16:25:00', 'פרשה': 'ויגש', 'כנסית שבת': '16:25', 'צאת שבת': '17:24'},
{'day': '2026-01-02 16:30:00', 'פרשה': 'ויחי', 'כנסית שבת': '16:30', 'צאת שבת': '17:28'},
{'day': '2026-01-09 16:35:00', 'פרשה': 'שמות', 'כנסית שבת': '16:35', 'צאת שבת': '17:34'},
{'day': '2026-01-16 16:41:00', 'פרשה': 'וארא', 'כנסית שבת': '16:41', 'צאת שבת': '17:40'},
{'day': '2026-01-23 16:48:00', 'פרשה': 'בא', 'כנסית שבת': '16:48', 'צאת שבת': '17:46'},
{'day': '2026-01-30 16:54:00', 'פרשה': 'בשלח', 'כנסית שבת': '16:54', 'צאת שבת': '17:52'},
{'day': '2026-02-06 17:01:00', 'פרשה': 'יתרו', 'כנסית שבת': '17:01', 'צאת שבת': '17:58'},
{'day': '2026-02-13 17:07:00', 'פרשה': 'משפטים', 'כנסית שבת': '17:07', 'צאת שבת': '18:03'},
{'day': '2026-02-20 17:13:00', 'פרשה': 'תרומה', 'כנסית שבת': '17:13', 'צאת שבת': '18:09'},
{'day': '2026-02-27 17:18:00', 'פרשה': 'תצוה', 'כנסית שבת': '17:18', 'צאת שבת': '18:14'},
{'day': '2026-03-06 17:24:00', 'פרשה': 'כי תשא', 'כנסית שבת': '17:24', 'צאת שבת': '18:19'},
{'day': '2026-03-13 17:29:00', 'פרשה': 'ויקהל-פקודי', 'כנסית שבת': '17:29', 'צאת שבת': '18:24'},
{'day': '2026-03-20 17:34:00', 'פרשה': 'ויקרא', 'כנסית שבת': '17:34', 'צאת שבת': '18:29'},
{'day': '2026-03-27 18:39:00', 'פרשה': 'צו', 'כנסית שבת': '18:39', 'צאת שבת': '19:34'},
{'day': '2026-04-10 18:48:00', 'פרשה': 'שמיני', 'כנסית שבת': '18:48', 'צאת שבת': '19:45'},
{'day': '2026-04-17 18:53:00', 'פרשה': 'תזריע-מצרע', 'כנסית שבת': '18:53', 'צאת שבת': '19:50'},
{'day': '2026-04-24 18:58:00', 'פרשה': 'אחרי מות-קדשים', 'כנסית שבת': '18:58', 'צאת שבת': '19:55'},
{'day': '2026-05-01 19:03:00', 'פרשה': 'אמור', 'כנסית שבת': '19:03', 'צאת שבת': '20:01'},
{'day': '2026-05-08 19:08:00', 'פרשה': 'בהר-בחקתי', 'כנסית שבת': '19:08', 'צאת שבת': '20:07'},
{'day': '2026-05-15 19:13:00', 'פרשה': 'במדבר', 'כנסית שבת': '19:13', 'צאת שבת': '20:12'},
{'day': '2026-05-22 19:18:00', 'פרשה': 'נשא', 'כנסית שבת': '19:18', 'צאת שבת': '20:18'},
{'day': '2026-05-29 19:22:00', 'פרשה': 'בהעלתך', 'כנסית שבת': '19:22', 'צאת שבת': '20:23'},
{'day': '2026-06-05 19:26:00', 'פרשה': 'שלח־לך', 'כנסית שבת': '19:26', 'צאת שבת': '20:27'},
{'day': '2026-06-12 19:29:00', 'פרשה': 'קורח', 'כנסית שבת': '19:29', 'צאת שבת': '20:30'},
{'day': '2026-06-19 19:31:00', 'פרשה': 'חקת', 'כנסית שבת': '19:31', 'צאת שבת': '20:33'},
{'day': '2026-06-26 19:33:00', 'פרשה': 'בלק', 'כנסית שבת': '19:33', 'צאת שבת': '20:34'},
{'day': '2026-07-03 19:33:00', 'פרשה': 'פינחס', 'כנסית שבת': '19:33', 'צאת שבת': '20:33'},
{'day': '2026-07-10 19:31:00', 'פרשה': 'מטות-מסעי', 'כנסית שבת': '19:31', 'צאת שבת': '20:31'},
{'day': '2026-07-17 19:29:00', 'פרשה': 'דברים', 'כנסית שבת': '19:29', 'צאת שבת': '20:28'},
{'day': '2026-07-24 19:25:00', 'פרשה': 'ואתחנן', 'כנסית שבת': '19:25', 'צאת שבת': '20:24'},
{'day': '2026-07-31 19:20:00', 'פרשה': 'עקב', 'כנסית שבת': '19:20', 'צאת שבת': '20:18'},
{'day': '2026-08-07 19:15:00', 'פרשה': 'ראה', 'כנסית שבת': '19:15', 'צאת שבת': '20:11'},
{'day': '2026-08-14 19:08:00', 'פרשה': 'שופטים', 'כנסית שבת': '19:08', 'צאת שבת': '20:04'},
{'day': '2026-08-21 19:00:00', 'פרשה': 'כי־תצא', 'כנסית שבת': '19:00', 'צאת שבת': '19:55'},
{'day': '2026-08-28 18:52:00', 'פרשה': 'כי־תבוא', 'כנסית שבת': '18:52', 'צאת שבת': '19:46'},
{'day': '2026-09-04 18:43:00', 'פרשה': 'נצבים-וילך', 'כנסית שבת': '18:43', 'צאת שבת': '19:37'},
{'day': '2026-09-18 18:25:00', 'פרשה': 'האזינו', 'כנסית שבת': '18:25', 'צאת שבת': '19:18'},
{'day': '2026-10-09 17:57:00', 'פרשה': 'בראשית', 'כנסית שבת': '17:57', 'צאת שבת': '18:51'},
{'day': '2026-10-16 17:49:00', 'פרשה': 'נח', 'כנסית שבת': '17:49', 'צאת שבת': '18:43'},
{'day': '2026-10-23 17:41:00', 'פרשה': 'לך־לך', 'כנסית שבת': '17:41', 'צאת שבת': '18:36'},
{'day': '2026-10-30 16:34:00', 'פרשה': 'וירא', 'כנסית שבת': '16:34', 'צאת שבת': '17:29'},
{'day': '2026-11-06 16:28:00', 'פרשה': 'חיי שרה', 'כנסית שבת': '16:28', 'צאת שבת': '17:24'},
{'day': '2026-11-13 16:24:00', 'פרשה': 'תולדות', 'כנסית שבת': '16:24', 'צאת שבת': '17:20'},
{'day': '2026-11-20 16:20:00', 'פרשה': 'ויצא', 'כנסית שבת': '16:20', 'צאת שבת': '17:17'},
{'day': '2026-11-27 16:18:00', 'פרשה': 'וישלח', 'כנסית שבת': '16:18', 'צאת שבת': '17:16'},
{'day': '2026-12-04 16:17:00', 'פרשה': 'וישב', 'כנסית שבת': '16:17', 'צאת שבת': '17:16'},
{'day': '2026-12-11 16:18:00', 'פרשה': 'מקץ', 'כנסית שבת': '16:18', 'צאת שבת': '17:17'},
{'day': '2026-12-18 16:21:00', 'פרשה': 'ויגש', 'כנסית שבת': '16:21', 'צאת שבת': '17:19'},
{'day': '2026-12-25 16:24:00', 'פרשה': 'ויחי', 'כנסית שבת': '16:24', 'צאת שבת': '17:23'},
{'day': '2027-01-01 16:29:00', 'פרשה': 'שמות', 'כנסית שבת': '16:29', 'צאת שבת': '17:28'},
{'day': '2027-01-08 16:34:00', 'פרשה': 'וארא', 'כנסית שבת': '16:34', 'צאת שבת': '17:33'},
{'day': '2027-01-15 16:40:00', 'פרשה': 'בא', 'כנסית שבת': '16:40', 'צאת שבת': '17:39'},
{'day': '2027-01-22 16:47:00', 'פרשה': 'בשלח', 'כנסית שבת': '16:47', 'צאת שבת': '17:45'},
{'day': '2027-01-29 16:53:00', 'פרשה': 'יתרו', 'כנסית שבת': '16:53', 'צאת שבת': '17:51'},
{'day': '2027-02-05 16:59:00', 'פרשה': 'משפטים', 'כנסית שבת': '16:59', 'צאת שבת': '17:56'},
{'day': '2027-02-12 17:06:00', 'פרשה': 'תרומה', 'כנסית שבת': '17:06', 'צאת שבת': '18:02'},
{'day': '2027-02-19 17:12:00', 'פרשה': 'תצוה', 'כנסית שבת': '17:12', 'צאת שבת': '18:08'},
{'day': '2027-02-26 17:17:00', 'פרשה': 'כי תשא', 'כנסית שבת': '17:17', 'צאת שבת': '18:13'},
{'day': '2027-03-05 17:23:00', 'פרשה': 'ויקהל', 'כנסית שבת': '17:23', 'צאת שבת': '18:18'},
{'day': '2027-03-12 17:28:00', 'פרשה': 'פקודי', 'כנסית שבת': '17:28', 'צאת שבת': '18:23'},
{'day': '2027-03-19 17:33:00', 'פרשה': 'ויקרא', 'כנסית שבת': '17:33', 'צאת שבת': '18:28'},
{'day': '2027-03-26 18:38:00', 'פרשה': 'צו', 'כנסית שבת': '18:38', 'צאת שבת': '19:33'},
{'day': '2027-04-02 18:43:00', 'פרשה': 'שמיני', 'כנסית שבת': '18:43', 'צאת שבת': '19:38'},
{'day': '2027-04-09 18:47:00', 'פרשה': 'תזריע', 'כנסית שבת': '18:47', 'צאת שבת': '19:44'},
{'day': '2027-04-16 18:52:00', 'פרשה': 'מצרע', 'כנסית שבת': '18:52', 'צאת שבת': '19:49'},
{'day': '2027-04-30 19:02:00', 'פרשה': 'אחרי מות', 'כנסית שבת': '19:02', 'צאת שבת': '20:00'},
{'day': '2027-05-07 19:07:00', 'פרשה': 'קדשים', 'כנסית שבת': '19:07', 'צאת שבת': '20:06'},
{'day': '2027-05-14 19:12:00', 'פרשה': 'אמור', 'כנסית שבת': '19:12', 'צאת שבת': '20:11'},
{'day': '2027-05-21 19:17:00', 'פרשה': 'בהר', 'כנסית שבת': '19:17', 'צאת שבת': '20:17'},
{'day': '2027-05-28 19:21:00', 'פרשה': 'בחקתי', 'כנסית שבת': '19:21', 'צאת שבת': '20:22'},
{'day': '2027-06-04 19:25:00', 'פרשה': 'במדבר', 'כנסית שבת': '19:25', 'צאת שבת': '20:26'},
{'day': '2027-06-11 19:29:00', 'פרשה': 'נשא', 'כנסית שבת': '19:29', 'צאת שבת': '20:30'},
{'day': '2027-06-18 19:31:00', 'פרשה': 'בהעלתך', 'כנסית שבת': '19:31', 'צאת שבת': '20:32'},
{'day': '2027-06-25 19:32:00', 'פרשה': 'שלח־לך', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
{'day': '2027-07-02 19:33:00', 'פרשה': 'קורח', 'כנסית שבת': '19:33', 'צאת שבת': '20:33'},
{'day': '2027-07-09 19:32:00', 'פרשה': 'חקת', 'כנסית שבת': '19:32', 'צאת שבת': '20:32'},
{'day': '2027-07-16 19:29:00', 'פרשה': 'בלק', 'כנסית שבת': '19:29', 'צאת שבת': '20:29'},
{'day': '2027-07-23 19:26:00', 'פרשה': 'פינחס', 'כנסית שבת': '19:26', 'צאת שבת': '20:25'},
{'day': '2027-07-30 19:21:00', 'פרשה': 'מטות-מסעי', 'כנסית שבת': '19:21', 'צאת שבת': '20:19'},
{'day': '2027-08-06 19:16:00', 'פרשה': 'דברים', 'כנסית שבת': '19:16', 'צאת שבת': '20:12'},
{'day': '2027-08-13 19:09:00', 'פרשה': 'ואתחנן', 'כנסית שבת': '19:09', 'צאת שבת': '20:05'},
{'day': '2027-08-20 19:01:00', 'פרשה': 'עקב', 'כנסית שבת': '19:01', 'צאת שבת': '19:57'},
{'day': '2027-08-27 18:53:00', 'פרשה': 'ראה', 'כנסית שבת': '18:53', 'צאת שבת': '19:48'},
{'day': '2027-09-03 18:45:00', 'פרשה': 'שופטים', 'כנסית שבת': '18:45', 'צאת שבת': '19:39'},
{'day': '2027-09-10 18:36:00', 'פרשה': 'כי־תצא', 'כנסית שבת': '18:36', 'צאת שבת': '19:29'},
{'day': '2027-09-17 18:26:00', 'פרשה': 'כי־תבוא', 'כנסית שבת': '18:26', 'צאת שבת': '19:20'},
{'day': '2027-09-24 18:17:00', 'פרשה': 'נצבים-וילך', 'כנסית שבת': '18:17', 'צאת שבת': '19:10'},
{'day': '2027-10-08 17:59:00', 'פרשה': 'האזינו', 'כנסית שבת': '17:59', 'צאת שבת': '18:52'},
{'day': '2027-10-29 17:36:00', 'פרשה': 'בראשית', 'כנסית שבת': '17:36', 'צאת שבת': '18:30'},
{'day': '2027-11-05 16:29:00', 'פרשה': 'נח', 'כנסית שבת': '16:29', 'צאת שבת': '17:25'},
{'day': '2027-11-12 16:24:00', 'פרשה': 'לך־לך', 'כנסית שבת': '16:24', 'צאת שבת': '17:20'},
{'day': '2027-11-19 16:21:00', 'פרשה': 'וירא', 'כנסית שבת': '16:21', 'צאת שבת': '17:17'},
{'day': '2027-11-26 16:18:00', 'פרשה': 'חיי שרה', 'כנסית שבת': '16:18', 'צאת שבת': '17:16'},
{'day': '2027-12-03 16:17:00', 'פרשה': 'תולדות', 'כנסית שבת': '16:17', 'צאת שבת': '17:15'},
{'day': '2027-12-10 16:18:00', 'פרשה': 'ויצא', 'כנסית שבת': '16:18', 'צאת שבת': '17:16'},
{'day': '2027-12-17 16:20:00', 'פרשה': 'וישלח', 'כנסית שבת': '16:20', 'צאת שבת': '17:19'},
{'day': '2027-12-24 16:23:00', 'פרשה': 'וישב', 'כנסית שבת': '16:23', 'צאת שבת': '17:22'},
{'day': '2027-12-31 16:28:00', 'פרשה': 'מקץ', 'כנסית שבת': '16:28', 'צאת שבת': '17:27'},
{'day': '2028-01-07 16:33:00', 'פרשה': 'ויגש', 'כנסית שבת': '16:33', 'צאת שבת': '17:32'},
{'day': '2028-01-14 16:39:00', 'פרשה': 'ויחי', 'כנסית שבת': '16:39', 'צאת שבת': '17:38'},
{'day': '2028-01-21 16:45:00', 'פרשה': 'שמות', 'כנסית שבת': '16:45', 'צאת שבת': '17:43'},
{'day': '2028-01-28 16:52:00', 'פרשה': 'וארא', 'כנסית שבת': '16:52', 'צאת שבת': '17:49'},
{'day': '2028-02-04 16:58:00', 'פרשה': 'בא', 'כנסית שבת': '16:58', 'צאת שבת': '17:55'},
{'day': '2028-02-11 17:05:00', 'פרשה': 'בשלח', 'כנסית שבת': '17:05', 'צאת שבת': '18:01'},
{'day': '2028-02-18 17:11:00', 'פרשה': 'יתרו', 'כנסית שבת': '17:11', 'צאת שבת': '18:07'},
{'day': '2028-02-25 17:16:00', 'פרשה': 'משפטים', 'כנסית שבת': '17:16', 'צאת שבת': '18:12'},
{'day': '2028-03-03 17:22:00', 'פרשה': 'תרומה', 'כנסית שבת': '17:22', 'צאת שבת': '18:17'},
{'day': '2028-03-10 17:27:00', 'פרשה': 'תצוה', 'כנסית שבת': '17:27', 'צאת שבת': '18:22'},
{'day': '2028-03-17 17:32:00', 'פרשה': 'כי תשא', 'כנסית שבת': '17:32', 'צאת שבת': '18:27'},
{'day': '2028-03-24 18:37:00', 'פרשה': 'ויקהל-פקודי', 'כנסית שבת': '18:37', 'צאת שבת': '19:32'},
{'day': '2028-03-31 18:42:00', 'פרשה': 'ויקרא', 'כנסית שבת': '18:42', 'צאת שבת': '19:37'},
{'day': '2028-04-07 18:46:00', 'פרשה': 'צו', 'כנסית שבת': '18:46', 'צאת שבת': '19:43'},
{'day': '2028-04-21 18:56:00', 'פרשה': 'שמיני', 'כנסית שבת': '18:56', 'צאת שבת': '19:53'},
{'day': '2028-04-28 19:01:00', 'פרשה': 'תזריע-מצרע', 'כנסית שבת': '19:01', 'צאת שבת': '19:59'},
{'day': '2028-05-05 19:06:00', 'פרשה': 'אחרי מות-קדשים', 'כנסית שבת': '19:06', 'צאת שבת': '20:05'},
{'day': '2028-05-12 19:11:00', 'פרשה': 'אמור', 'כנסית שבת': '19:11', 'צאת שבת': '20:10'},
{'day': '2028-05-19 19:16:00', 'פרשה': 'בהר-בחקתי', 'כנסית שבת': '19:16', 'צאת שבת': '20:16'},
{'day': '2028-05-26 19:20:00', 'פרשה': 'במדבר', 'כנסית שבת': '19:20', 'צאת שבת': '20:21'},
{'day': '2028-06-02 19:25:00', 'פרשה': 'נשא', 'כנסית שבת': '19:25', 'צאת שבת': '20:26'},
{'day': '2028-06-09 19:28:00', 'פרשה': 'בהעלתך', 'כנסית שבת': '19:28', 'צאת שבת': '20:29'},
{'day': '2028-06-16 19:31:00', 'פרשה': 'שלח־לך', 'כנסית שבת': '19:31', 'צאת שבת': '20:32'},
{'day': '2028-06-23 19:32:00', 'פרשה': 'קורח', 'כנסית שבת': '19:32', 'צאת שבת': '20:33'},
{'day': '2028-06-30 19:33:00', 'פרשה': 'חקת', 'כנסית שבת': '19:33', 'צאת שבת': '20:33'},
{'day': '2028-07-07 19:32:00', 'פרשה': 'בלק', 'כנסית שבת': '19:32', 'צאת שבת': '20:32'},
{'day': '2028-07-14 19:30:00', 'פרשה': 'פינחס', 'כנסית שבת': '19:30', 'צאת שבת': '20:29'},
{'day': '2028-07-21 19:27:00', 'פרשה': 'מטות-מסעי', 'כנסית שבת': '19:27', 'צאת שבת': '20:25'},
{'day': '2028-07-28 19:22:00', 'פרשה': 'דברים', 'כנסית שבת': '19:22', 'צאת שבת': '20:20'},
{'day': '2028-08-04 19:17:00', 'פרשה': 'ואתחנן', 'כנסית שבת': '19:17', 'צאת שבת': '20:14'},
{'day': '2028-08-11 19:10:00', 'פרשה': 'עקב', 'כנסית שבת': '19:10', 'צאת שבת': '20:06'},
{'day': '2028-08-18 19:03:00', 'פרשה': 'ראה', 'כנסית שבת': '19:03', 'צאת שבת': '19:58'},
{'day': '2028-08-25 18:55:00', 'פרשה': 'שופטים', 'כנסית שבת': '18:55', 'צאת שבת': '19:50'},
{'day': '2028-09-01 18:46:00', 'פרשה': 'כי־תצא', 'כנסית שבת': '18:46', 'צאת שבת': '19:40'},
{'day': '2028-09-08 18:37:00', 'פרשה': 'כי־תבוא', 'כנסית שבת': '18:37', 'צאת שבת': '19:31'},
{'day': '2028-09-15 18:28:00', 'פרשה': 'נצבים-וילך', 'כנסית שבת': '18:28', 'צאת שבת': '19:21'},
{'day': '2028-09-22 18:19:00', 'פרשה': 'האזינו', 'כנסית שבת': '18:19', 'צאת שבת': '19:12'},
        ]

        self.tekufa_list = []
        tekufa_ics_path = self.template_path.parent / "tekufa_2025_2035.ics"
        if tekufa_ics_path.exists():
            self.tekufa_list = parse_tekufa_ics(tekufa_ics_path)
        else:
            print(f"⚠️ Fichier tekufa_2025_2035.ics non trouvé à {tekufa_ics_path}")

    def sanitize_filename(self, value: str) -> str:
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

    def get_tekufa_for_shabbat(self, shabbat_date):
        week_start = datetime.combine(shabbat_date, datetime.min.time())
        week_end = week_start + timedelta(days=6, hours=23, minutes=59)
        for dt, summary in self.tekufa_list:
            if week_start <= dt <= week_end:
                return dt, summary
        return None

    def get_tekufa_for_next_week(self, shabbat_date):
        week_start = datetime.combine(shabbat_date, datetime.min.time())
        week_end = week_start + timedelta(days=6, hours=23, minutes=59)
        for dt, summary in self.tekufa_list:
            if week_start <= dt <= week_end:
                return dt, summary
        return None

    def update_excel_with_mevarchim_column(self, excel_path: Path):
        import openpyxl
        # Charger tous les onglets existants
        if excel_path.exists():
            xls = pd.ExcelFile(excel_path)
            sheets = {name: xls.parse(name) for name in xls.sheet_names}
            if "שבתות השנה" in sheets:
                df = sheets["שבתות השנה"]
            else:
                df = pd.DataFrame(self.yearly_shabbat_data)
                df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date
        else:
            sheets = {}
            df = pd.DataFrame(self.yearly_shabbat_data)
            df["day"] = pd.to_datetime(df["day"], format="%Y-%m-%d %H:%M:%S").dt.date

        min_date = df["day"].min()
        max_date = df["day"].max()
        rosh_dates = self.fetch_roshchodesh_dates(min_date, max_date + timedelta(days=7))
        df = self.identify_shabbat_mevarchim(df, rosh_dates)

        # CALCUL des colonnes horaires intermédiaires
        def compute_times(row):
            row_date = row["day"]
            if isinstance(row_date, pd.Timestamp):
                row_date = row_date.date()
            sunday_date = row_date + timedelta(days=2)
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
                minha_midweek = self.format_time(self.round_to_nearest_five(min_sunset - 18))
                arvit_midweek = self.format_time(self.round_to_next_five(max_sunset + 20))
            else:
                minha_midweek = ""
                arvit_midweek = ""

            return pd.Series({
                "שקיעה Dimanche": sunday_sunset,
                "שקיעה Jeudi": thursday_sunset,
                "מנחה ביניים": minha_midweek,
                "ערבית ביניים": arvit_midweek
            })

        times_df = df.apply(compute_times, axis=1)

        # Supprimer les anciennes colonnes si présentes
        cols_to_remove = ["שקיעה Dimanche", "שקיעה Jeudi", "מנחה ביניים", "ערבית ביניים"]
        for col in cols_to_remove:
            if col in df.columns:
                df.drop(col, axis=1, inplace=True)

        df = pd.concat([df, times_df], axis=1)

        # Ajout de la colonne molad uniquement pour שבת מברכין
        molad_col = []
        for i, row in df.iterrows():
            is_mevarchim = row.get("שבת מברכין") in [True, "כן"]
            if is_mevarchim:
                try:
                    molad_str = get_next_month_molad(row["day"])
                except Exception:
                    molad_str = ""
            else:
                molad_str = ""
            molad_col.append(molad_str)
        df["molad"] = molad_col

        # Ajout de la colonne tekoufa
        tekufa_col = []
        for shabbat in df["day"]:
            tekufa = self.get_tekufa_for_shabbat(shabbat)
            if tekufa:
                dt, summary = tekufa
                tekufa_col.append(dt.strftime("%Y-%m-%d %H:%M"))
            else:
                tekufa_col.append("")
        df["tekoufa"] = tekufa_col

        # Déplacer tekoufa à la fin
        if 'tekoufa' in df.columns:
            tekoufa = df.pop('tekoufa')
            df['tekoufa'] = tekoufa

        # Mise à jour de la feuille
        sheets["שבתות השנה"] = df

        # Écriture de tous les onglets
        with pd.ExcelWriter(str(excel_path), engine="openpyxl", mode="w") as writer:
            for name, sheet_df in sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)

        print("✅ Colonnes mises à jour : שבת מברכין, horaires ביניים, molad (sélectif) et tekoufa (fin).")


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
                is_mevarchim_excel = row.get("שבת מברכין", False) == True or row.get("שבת מברכין", "") == "Oui"
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
            candle_time = datetime.strptime(str(row["כנסית שבת"]), "%H:%M").time()
            havdalah_time = datetime.strptime(str(row["צאת שבת"]), "%H:%M").time()
            shabbat_start = datetime.combine(row["day"], candle_time)
            shabbat_end = datetime.combine(row["day"], havdalah_time)
            is_mevarchim_excel = row.get("שבת מברכין", False) == True or row.get("שבת מברכין", "") == "Oui"
            return [{
                "date": shabbat_date,
                "start": shabbat_start,
                "end": shabbat_end,
                "parasha": row.get("פרשה", ""),
                "parasha_hebrew": row.get("פרשה_עברית", rowget("פרשה", "")),
                "candle_lighting": row["כנסית שבת"],
                "is_mevarchim": is_mevarchim_excel
            }]

    def round_to_nearest_five(self, minutes):
        return (minutes // 5) * 5

    def round_to_next_five(self, minutes):
        return ((minutes + 4) // 5) * 5 if minutes is not None else None

    def format_time(self, minutes):
        if minutes is None or minutes < 0:
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
            "shiur_nashim": 16 * 60 +15,
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

        sunday_date = shabbat_start.date() + timedelta(days=2)
        s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
        sunday_sunset = s_sunday.get("sunset", None)
        thursday_date = sunday_date + timedelta(days=4)
        s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
        thursday_sunset = s_thursday.get("sunset", None)

        def to_minutes(t):
            if t is None:
                return None
            try:
                h, m = map(int, t.split(":"))
                return h * 60 + m
            except ValueError:
                return None

        if sunday_sunset and thursday_sunset:
            sunday_sunset_str = sunday_sunset.strftime("%H:%M")
            thursday_sunset_str = thursday_sunset.strftime("%H:%M")
            sunday_sunset_min = to_minutes(sunday_sunset_str)
            thursday_sunset_min = to_minutes(thursday_sunset_str)
            min_sunset = min(sunday_sunset_min, thursday_sunset_min)
            minha_hol_minutes = min_sunset - 18
            times["mincha_hol"] = self.round_to_nearest_five(minha_hol_minutes)
            max_sunset = max(sunday_sunset_min, thursday_sunset_min)
            arvit_hol_minutes = max_sunset + 20
            times["arvit_hol"] = self.round_to_next_five(arvit_hol_minutes)
        else:
            times["mincha_hol"] = None
            times["arvit_hol"] = None

        times["arvit_motsach"] = self.round_to_nearest_five(end_minutes - 9)

        return times

    def create_image(self, times, parasha, parasha_hebrew,
                     shabbat_end, candle_lighting, shabbat_date, is_mevarchim=False):
        try:
            template = self.template_path
            if is_mevarchim:
                rc_template = self.template_path.parent / "template_rosh_hodesh.jpg"
                if rc_template.exists():
                    template = rc_template

            # Définir le chemin d'accès aux icônes (UNE SEULE FOIS)
            icon_path = self.template_path.parent / "resources"

            with Image.open(template) as img:
                try:
                    img_w, img_h = img.size # Définir img_w et img_h ici
                    draw = ImageDraw.Draw(img)
                    font = self._font
                    bold = self._arial_bold_font
                    time_x = 120

                    # Charger les icônes (UNE SEULE FOIS, GÉRER LES EXCEPTIONS)
                    first_moon_icon = None
                    full_moon_icon = None
                    eau_icon = None
                    eau2_icon = None

                    try:
                        first_moon_icon = Image.open(icon_path / "first_moon.png").convert("RGBA").resize((48, 48), Image.LANCZOS)
                        full_moon_icon = Image.open(icon_path / "full_moon.png").convert("RGBA").resize((48, 48), Image.LANCZOS)
                        eau_icon = Image.open(icon_path / "eau.png").resize((64, 64), Image.LANCZOS)
                    except FileNotFoundError as e:
                        print(f"❌ Erreur: Une ou plusieurs icônes sont introuvables: {e}")

                    # Affichage des horaires
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
                    reversed_parasha = reverse_hebrew_text(parasha_hebrew)
                    draw.text((300, 280), parasha_hebrew, fill="blue", font=bold, anchor="mm")

                    if is_mevarchim:
                        molad_str = get_next_month_molad(shabbat_date)
                        draw.text((200, img_h - 300), molad_str, fill="blue", font=font)
                        rc_days = get_rosh_chodesh_days_for_next_month(shabbat_date)
                        rosh_lines = []
                        for gdate, m, y, d in rc_days:
                            day_name_he = get_weekday_name_hebrew(gdate)
                            month_name = get_jewish_month_name_hebrew(m, y)
                            rosh_lines.append(
                                f"ראש חודש: יום {day_name_he} {gdate.strftime('%d/%m/%Y')} {month_name} ({d})"
                            )
                        for i, rc_line in enumerate(rosh_lines):
                            draw.text(
                                (200, img_h - 260 + 40 * i),
                                rc_line,
                                fill="blue",
                                font=font
                            )
                    if not is_mevarchim:
                        try:
                            previous_rosh = find_previous_rosh_chodesh(shabbat_date)
                            molad_dt, latest_kiddush_levana = calculate_last_kiddush_levana_date(previous_rosh)
                            start_kiddush_levana = molad_dt + timedelta(days=6)
                            shabbat_date_only = shabbat_date.date()
                            y_start = img_h - 320
                            y_end = img_h - 260

                            if shabbat_date_only < start_kiddush_levana.date():
                                msg_start = f"תחילת אמירה ברכת הלבנה ביום: {start_kiddush_levana.strftime('%d/%m/%Y')}"
                                msg_end = f"תאריך אחרון לאמירת ברכת הלבנה: {latest_kiddush_levana.strftime('%d/%m/%Y')}"

                                # Affichage de l'icône first_moon
                                if first_moon_icon:
                                    img.paste(first_moon_icon, (55, y_start - 10), first_moon_icon)
                                draw.text((55 + 48 + 10, y_start), msg_start, fill="purple", font=font)

                                # Affichage de l'icône full_moon
                                if full_moon_icon:
                                    img.paste(full_moon_icon, (55, y_end - 10), full_moon_icon)
                                draw.text((55 + 48 + 10, y_end), msg_end, fill="purple", font=font)

                            elif start_kiddush_levana.date() <= shabbat_date_only <= latest_kiddush_levana.date():
                                msg_end = f"ערב אחרון לאמירת ברכת הלבנה: {latest_kiddush_levana.strftime('%d/%m/%Y')}"
                                if full_moon_icon:
                                    img.paste(full_moon_icon, (55, y_end - 10), full_moon_icon)
                                draw.text((55 + 48 + 10, y_end), msg_end, fill="purple", font=font)
                            else:
                                msg_ended = "התקופה של ברכת הלבנה הסתיימה."
                                draw.text((100, y_end), msg_ended, fill="red", font=font)
                        except Exception as e:
                            print(f"❌ Erreur lors de l'affichage de la Birkat Halevana : {e}")

                    # 🔽 Affichage des jeunes (à partir de jeunes.ics)
                    fasts = get_fast_from_ics_for_week(shabbat_date)
                    if fasts:
                        fast_y = img_h - 140
                        for fast in fasts:
                            hebrew_name = fast["nom"]
                            line = f"צום {hebrew_name}: {fast['start']} - {fast['end']}"
                            if eau2_icon:
                                img.paste(eau2_icon, (55, fast_y - 10), eau2_icon)
                            draw.text((55 + 64 + 10, fast_y), line, fill="blue", font=font)
                            fast_y += 50


                    # --- Tekoufa à venir : affichage en bleu ---
                    tekufa_next = self.get_tekufa_for_next_week(shabbat_date)
                    if tekufa_next:
                        dt, summary = tekufa_next
                        match = re.search(r'Tekufat\s+(\w+)\s+(\d{4})', summary)
                        hebrew_month = ""
                        if match:
                            name_map = {
                                'Nisan': 'ניסן', 'Tammuz': 'תמוז', 'Tishri': 'תשרי', 'Tevet': 'טבת'
                            }
                            month_eng = match.group(1)
                            hebrew_month = name_map.get(month_eng, month_eng)
                        tekufa_msg = f"תקופת {hebrew_month} ביום {dt.strftime('%d/%m/%Y')} בשעה {dt.strftime('%H:%M')}"

                        if eau_icon:
                            img.paste(eau_icon, (55, img_h - 200 - 15), eau_icon)
                        draw.text((55 + 48 + 10, img_h - 200), tekufa_msg, fill="blue", font=font)

                    safe_parasha = self.sanitize_filename(parasha)
                    output_filename = f"horaires_{safe_parasha}.jpeg"
                    output_path = self.output_dir / output_filename
                    print(f"✅ Image sauvegardée ici : {output_path}") # Ajout d'une instruction de débogage
                    img.save(str(output_path))
                    latest = self.output_dir / "latest-schedule.jpg"
                    if latest.exists():
                        latest.unlink()
                    shutil.copy(str(output_path), str(latest))
                    return output_path
                except Exception as e:
                    print(f"❌ Erreur lors du traitement de l'image: {e}")
                    return None
        except FileNotFoundError as e:
            print(f"❌ Erreur lors de l'ouverture du template: {e}")
            return None
        except Exception as e:
            print(f"❌ Erreur générale: {e}")
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
            "מוצאי שבת": shabbat_data["end"].strftime("%H:%M"),
            "שבת מברכין": "כן" if shabbat_data.get("is_mevarchim", False) else "לא"
        }

        tekufa_info = self.get_tekufa_for_shabbat(shabbat_data["date"])
        if tekufa_info:
            dt, summary = tekufa_info
            row["tekoufa"] = dt.strftime("%Y-%m-%d %H:%M")
        else:
            row["tekoufa"] = ""

        try:
            # Écrire les données de ce Chabbat dans un onglet 'השבת'
            df_current = pd.DataFrame([row])
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_current.to_excel(writer, sheet_name="השבת", index=False)
            print("✅ Onglet 'THIS SHABBAT' mis à jour avec les données du Chabbat en cours.")
        except Exception as e:
            print(f"❌ Erreur lors de la mise à jour de l’Excel: {e}")


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