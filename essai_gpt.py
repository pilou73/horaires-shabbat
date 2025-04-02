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
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Vérification des fichiers requis
        for file, desc in [(self.template_path, "Template"),
                           (self.font_path, "Police"),
                           (self.arial_bold_path, "Police Arial Bold")]:
            if not file.exists():
                raise FileNotFoundError(f"{desc} introuvable: {file}")
                
        try:
            self._test_font = ImageFont.truetype(str(self.font_path), 30)
            self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)
        except Exception as e:
            raise Exception(f"Erreur de chargement de la police: {e}")
            
        # Localisation pour Ramat Gan, Israël (pour les calculs astronomiques)
        self.ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)
        
        # Données annuelles statiques (exemple partiel pour 5785)
        self.yearly_shabbat_data = [
            {'day': '2024-12-06 00:00:00', 'פרשה': 'Vayetzei', 'כנסית שבת': '16:17', 'צאת שבת': '17:16'},
            {'day': '2024-12-13 00:00:00', 'פרשה': 'Vayishlach', 'כנסית שבת': '16:19', 'צאת שבת': '17:17'},
            {'day': '2024-12-20 00:00:00', 'פרשה': 'Vayeshev', 'כנסית שבת': '16:22', 'צאת שבת': '17:20'},
            # ... autres données
        ]
    
    @staticmethod
    def round_to_nearest_five(minutes):
        return (minutes // 5) * 5

    @staticmethod
    def format_time(minutes):
        hours = minutes // 60
        mins = minutes % 60
        return f"{hours:02d}:{mins:02d}"
    
    @staticmethod
    def to_minutes(time_val):
        # Si time_val est un objet datetime, le convertir en chaîne au format "HH:MM"
        if isinstance(time_val, datetime):
            time_val = time_val.strftime("%H:%M")
        h, m = map(int, time_val.split(':'))
        return h * 60 + m

    def is_summer(self, date_obj):
        """
        Détermine dynamiquement si la date est en période d'été.
        L'heure d'été commence le dernier vendredi de mars
        et l'heure d'hiver commence le dernier dimanche d'octobre.
        """
        year = date_obj.year
        
        # Calcul du dernier vendredi de mars
        last_day_march = datetime(year, 3, 31)
        days_to_friday = (last_day_march.weekday() - 4) % 7  # vendredi=4
        last_friday_march = last_day_march - timedelta(days=days_to_friday)
        
        # Calcul du dernier dimanche d'octobre
        last_day_oct = datetime(year, 10, 31)
        days_to_sunday = (last_day_oct.weekday() - 6) % 7  # dimanche=6
        last_sunday_oct = last_day_oct - timedelta(days=days_to_sunday)
        
        # Debug: affichage des bornes de l'heure d'été
        print(f"Début d'été (dernier vendredi de mars): {last_friday_march.date()}")
        print(f"Fin d'été (dernier dimanche d'octobre): {last_sunday_oct.date()}")
        
        return last_friday_march.date() <= date_obj.date() <= last_sunday_oct.date()

    def get_hebcal_times(self, start_date, end_date):
        tz = pytz.timezone('Asia/Jerusalem')
        base_url = "https://www.hebcal.com/shabbat"
        params = {
            'cfg': 'json',
            'geonameid': '293397',
            'b': '18',
            'M': 'on',
            'start': start_date.strftime('%Y-%m-%d'),
            'end': end_date.strftime('%Y-%m-%d')
        }
        try:
            response = requests.get(base_url, params=params)
            response.raise_for_status()
            data = response.json()
            shabbat_times = []
            for item in data['items']:
                if item['category'] == 'candles':
                    start_time = datetime.fromisoformat(item['date']).astimezone(tz)
                    havdalah_items = [i for i in data['items'] if i['category'] == 'havdalah']
                    parasha_items = [i for i in data['items'] if i['category'] == 'parashat']
                    if havdalah_items and parasha_items:
                        end_time = datetime.fromisoformat(havdalah_items[0]['date']).astimezone(tz)
                        parasha = parasha_items[0]['title'].replace('Parashat ', '')
                        parasha_hebrew = parasha_items[0].get('hebrew', parasha)
                        shabbat_times.append({
                            'date': start_time.date(),
                            'start': start_time,
                            'end': end_time,
                            'parasha': parasha,
                            'parasha_hebrew': parasha_hebrew,
                            'candle_lighting': start_time.strftime('%H:%M')
                        })
            print("Horaires récupérés via Hebcal:", shabbat_times)
            return shabbat_times
        except requests.RequestException as e:
            print(f"Erreur lors de la récupération des données: {e}")
            return []
    
    def calculate_times(self, shabbat):
        """
        Calcule l'ensemble des horaires selon tes règles.
        Des messages de vérification sont affichés pour suivre les calculs.
        """
        candle_minutes = self.to_minutes(shabbat['candle_lighting'])
        tzait_minutes = self.to_minutes(shabbat['end']) if shabbat['end'] else None
        if tzait_minutes is None:
            raise ValueError("L'heure de צאת שבת n'est pas définie.")
        print(f"candle_minutes: {candle_minutes}, tzait_minutes: {tzait_minutes}")
        
        # (2) Minha de la veille = heure d'entrée du Shabbat
        minha_eve = candle_minutes
        print(f"Minha veille (entrée): {minha_eve}")
        
        # (3) Shir HaShirim = au moins 9 minutes avant minha veille
        shir_hashirim = self.round_to_nearest_five(max(minha_eve - 9, 0))
        print(f"Shir HaShirim: {shir_hashirim}")
        
        # (4) Shacharit fixe à 7:45
        shacharit = 7 * 60 + 45
        
        # (5) Mincha Gdola (Minha 1) : fixe selon saison
        if self.is_summer(shabbat['start']):
            mincha_gdola = 13 * 60      # 13:00 en été
        else:
            mincha_gdola = 12 * 60 + 30  # 12:30 en hiver
        print(f"Mincha Gdola: {mincha_gdola}")
        
        # (6) Tehilim : en hiver = mincha_gdola + 45, en été = affichage double (17:00/...)
        tehilim_val = self.round_to_nearest_five(mincha_gdola + 45)
        if self.is_summer(shabbat['start']):
            tehilim = f"17:00/{self.format_time(tehilim_val)}"
        else:
            tehilim = self.format_time(tehilim_val)
        print(f"Tehilim: {tehilim}")
        
        # (7) Cours Parashat Hashavua = 3h avant tzait
        course_parasha = self.round_to_nearest_five(max(tzait_minutes - 180, 0))
        print(f"Cours Parasha: {course_parasha}")
        
        # (8) Arvit de fin de Shabbat : en été 10 minutes avant tzait, en hiver 5 minutes avant
        if self.is_summer(shabbat['start']):
            arvit_end = self.round_to_nearest_five(max(tzait_minutes - 10, 0))
        else:
            arvit_end = self.round_to_nearest_five(max(tzait_minutes - 5, 0))
        print(f"Arvit de fin de Shabbat: {arvit_end}")
        
        # (9) Minha (deuxième) = 90 minutes avant arvit_end
        minha_two = self.round_to_nearest_five(max(arvit_end - 90, 0))
        print(f"Minha deux: {minha_two}")
        
        # (10) Cours du Rav = 45 minutes avant minha_two
        course_rav = self.round_to_nearest_five(max(minha_two - 45, 0))
        print(f"Cours du Rav: {course_rav}")
        
        # (11) Cours pour femmes = fixe à 16:00
        course_women = 16 * 60
        
        # (12) En semaine, Minha de milieu de semaine
        sunday_date = shabbat['start'].date() + timedelta(days=2)
        thursday_date = sunday_date + timedelta(days=4)
        s_sunday = sun(self.ramat_gan.observer, date=sunday_date, tzinfo=self.ramat_gan.timezone)
        s_thursday = sun(self.ramat_gan.observer, date=thursday_date, tzinfo=self.ramat_gan.timezone)
        sunday_sunset = s_sunday['sunset'].strftime("%H:%M")
        thursday_sunset = s_thursday['sunset'].strftime("%H:%M")
        base_midweek = min(self.to_minutes(sunday_sunset), self.to_minutes(thursday_sunset)) - 20
        minha_midweek = self.round_to_nearest_five(max(base_midweek, 0))
        print(f"Minha milieu de semaine: {minha_midweek}")
        
        # (13) En semaine, Arvit se base sur l'heure de צאת הכוכבים (dusk) de dimanche et jeudi,
        # choisissant une heure (multiple de 5) STRICTEMENT avant ces heures.
        twilight_sunday = self.to_minutes(s_sunday['dusk'].strftime("%H:%M"))
        twilight_thursday = self.to_minutes(s_thursday['dusk'].strftime("%H:%M"))
        twilight_minutes = min(twilight_sunday, twilight_thursday)
        arvit_midweek = self.round_to_nearest_five(max(twilight_minutes - 5, 0))
        print(f"Arvit milieu de semaine: {arvit_midweek}")
        
        times = {
            'minha_eve': self.format_time(minha_eve),
            'shir_hashirim': self.format_time(shir_hashirim),
            'shacharit': self.format_time(shacharit),
            'mincha_gdola': self.format_time(mincha_gdola),
            'tehilim': tehilim,
            'course_parasha': self.format_time(course_parasha),
            'arvit_end': self.format_time(arvit_end),
            'minha_two': self.format_time(minha_two),
            'course_rav': self.format_time(course_rav),
            'course_women': self.format_time(course_women),
            'minha_midweek': self.format_time(minha_midweek),
            'arvit_midweek': self.format_time(arvit_midweek)
        }
        print("Horaires calculés:", times)
        return times

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting, shabbat_date):
        try:
            with Image.open(self.template_path) as img:
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(str(self.font_path), 30)
                time_x = 120

                # Positions d'affichage (à ajuster selon le template)
                positions = [
                    (time_x, 400, ('shir_hashirim', "Shir HaShirim")),
                    (time_x, 475, ('minha_eve', "כניסת שבת")),
                    (time_x, 510, ('shacharit', "שחרית")),
                    (time_x, 550, ('mincha_gdola', "מנחה גדולה")),
                    (time_x, 590, ('course_parasha', "קורס פרשת השבוע")),
                    (time_x, 630, ('tehilim', "תהילים")),
                    (time_x, 670, ('course_rav', "קורס עם הרב")),
                    (time_x, 710, ('course_women', "קורס לנשים")),
                    (time_x, 790, ('arvit_end', 'ערבית מוצ"ש'))
                ]
                for x, y, (key, label) in positions:
                    time_str = times.get(key, "")
                    draw.text((x, y), f"{label}: {time_str}", fill="black", font=font)
                
                # Horaires en semaine affichés en bas
                week_schedule = f"מנחה ביניים: {times['minha_midweek']}   ערבית (שבוע): {times['arvit_midweek']}"
                draw.text((time_x, 950), week_schedule, fill="green", font=font)
                
                draw.text((300, 280), parasha_hebrew, fill="blue", font=self._arial_bold_font, anchor="mm")
                
                output_path = self.output_dir / f"horaires_{parasha}.jpeg"
                img.save(str(output_path))
                print(f"Image sauvegardée: {output_path}")
                return output_path
        except Exception as e:
            print(f"Erreur lors de la création de l'image: {e}")
            return None
    
    def update_excel(self, shabbat, times):
        excel_path = self.output_dir / "horaires_shabbat.xlsx"
        row = {
            'תאריך': shabbat['date'].strftime('%d/%m/%Y'),
            'פרשה': shabbat['parasha'],
            'כניסת שבת': shabbat['candle_lighting'],
            'צאת שבת': shabbat['end'],
            'שיר השירים': times['shir_hashirim'],
            'שחרית': times['shacharit'],
            'מנחה גדולה': times['mincha_gdola'],
            'תהילים': times['tehilim'],
            'קורס פרשת השבוע': times['course_parasha'],
            'קורס עם הרב': times['course_rav'],
            'מנחה שנייה': times['minha_two'],
            'ערבית מוצ"ש': times['arvit_end'],
            'מנחה ביניים (שבוע)': times['minha_midweek'],
            'ערבית (שבוע)': times['arvit_midweek'],
            'קורס לנשים': times['course_women']
        }
        yearly_df = pd.DataFrame(self.yearly_shabbat_data)
        row_df = pd.DataFrame([row])
        try:
            if excel_path.exists():
                with pd.ExcelWriter(str(excel_path), engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    row_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False,
                                    startrow=writer.sheets['Sheet1'].max_row)
                    yearly_df.to_excel(writer, sheet_name='שבתות השנה', index=False)
            else:
                with pd.ExcelWriter(str(excel_path), engine='openpyxl') as writer:
                    row_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    yearly_df.to_excel(writer, sheet_name='שבתות השנה', index=False)
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
        times = self.calculate_times(shabbat)
        image_path = self.create_image(times, shabbat['parasha'], shabbat['parasha_hebrew'], shabbat['end'], shabbat['candle_lighting'], shabbat['start'])
        if not image_path:
            print("Échec de la génération de l'image")
        self.update_excel(shabbat, times)

def main():
    try:
        if getattr(sys, 'frozen', False):
            base_path = Path(sys.executable).parent
        elif '__file__' in globals():
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
