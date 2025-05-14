import os
import sys
from pathlib import Path
from datetime import datetime, timedelta
import math
import requests
import pytz
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

class ShabbatScheduleGenerator:
    def __init__(self, template_path, font_path, arial_bold_path, output_dir):
        self.template_path = Path(template_path)
        self.font_path = Path(font_path)
        self.arial_bold_path = Path(arial_bold_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)  # Crée le dossier output s'il n'existe pas
        
        # Vérifie que les fichiers nécessaires existent
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template introuvable: {self.template_path}")
        if not self.font_path.exists():
            raise FileNotFoundError(f"Police introuvable: {self.font_path}")
        if not self.arial_bold_path.exists():
            raise FileNotFoundError(f"Police Arial Bold introuvable: {self.arial_bold_path}")
            
        try:
            # Charge les polices
            self._test_font = ImageFont.truetype(str(self.font_path), 30)
            self._arial_bold_font = ImageFont.truetype(str(self.arial_bold_path), 40)
        except Exception as e:
            raise Exception(f"Erreur de chargement de la police: {e}")

        # Données de l'onglet שבתות השנה intégrées dans le code
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

    def get_hebcal_times(self, start_date, end_date):
        tz = pytz.timezone('Asia/Jerusalem')
        base_url = "https://www.hebcal.com/shabbat"
        
        params = {
            'cfg': 'json',
            'geonameid': '293397',  # Ramat Gan, Israel
            'b': '18',  # Heure de fin de Shabbat (Havdalah)
            'M': 'on',  # Inclure les Parashot
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
                        parasha_hebrew = parasha_items[0].get('hebrew', parasha)  # Récupérer le nom en hébreu
                        
                        shabbat_times.append({
                            'date': start_time.date(),
                            'start': start_time,
                            'end': end_time,
                            'parasha': parasha,
                            'parasha_hebrew': parasha_hebrew,
                            'candle_lighting': start_time.strftime('%H:%M')  # Ajout de l'heure d'allumage des bougies
                        })
            
            return shabbat_times

        except requests.RequestException as e:
            print(f"Erreur lors de la récupération des données: {e}")
            return []

    def calculate_times(self, shabbat_start, shabbat_end):
        start_minutes = shabbat_start.hour * 60 + shabbat_start.minute
        end_minutes = shabbat_end.hour * 60 + shabbat_end.minute
        
        times = {
            'mincha_kabbalat': start_minutes,  # Pas d'arrondi pour מנחה et קבלת שבת
            'shir_hashirim': self.round_to_nearest_five(start_minutes - 15),
            'shacharit': self.round_to_nearest_five(7 * 60 + 45),
            'mincha_gdola': self.round_to_nearest_five(12 * 60 + 45),
            'parashat_hashavua': self.round_to_nearest_five(end_minutes - (3 * 60)),
            'tehilim': self.round_to_nearest_five(15 * 60),
            'nashim': self.round_to_nearest_five(15 * 60),
            'shiur_rav': self.round_to_nearest_five(end_minutes - (2 * 60 + 15))
        }
        
        times['mincha_2'] = self.round_to_nearest_five(times['shiur_rav'] + 45)
        times['arvit'] = self.round_to_nearest_five(end_minutes - 5)
        
        return times

    def format_time(self, minutes):
        hours = minutes // 60
        mins = minutes % 60
        return f"{hours:02d}:{mins:02d}"

    #def reverse_hebrew_text(self, text):
        # Inverser le texte en hébreu pour un affichage correct
        #return text[::-1]

    def round_to_nearest_five(self, minutes):
        # Arrondir les minutes au multiple de 5 le plus proche
        return round(minutes / 5) * 5

    def get_next_shabbat_time(self, current_shabbat_date):
        try:
            # Convertir la date actuelle en format datetime pour la comparaison
            current_date = datetime(current_shabbat_date.year, current_shabbat_date.month, current_shabbat_date.day)
            
            # Trouver le Shabbat suivant dans les données intégrées
            for shabbat in self.yearly_shabbat_data:
                shabbat_date = datetime.strptime(shabbat['day'], '%Y-%m-%d %H:%M:%S')
                if shabbat_date > current_date:
                    # Récupérer l'heure d'entrée du Shabbat
                    shabbat_entry_time = shabbat['כנסית שבת'].split()[0]
                    
                    # Convertir l'heure en minutes pour l'arrondir
                    hours, minutes = map(int, shabbat_entry_time.split(':'))
                    total_minutes = hours * 60 + minutes
                    rounded_minutes = self.round_to_nearest_five(total_minutes)
                    
                    # Convertir les minutes arrondies en heure
                    rounded_hours = rounded_minutes // 60
                    rounded_mins = rounded_minutes % 60
                    rounded_time = f"{rounded_hours:02d}:{rounded_mins:02d}"
                    
                    # Retourner la date et l'heure arrondie
                    return shabbat_date.strftime('%d/%m/%Y'), rounded_time
            return None, None
        except Exception as e:
            print(f"Erreur lors de la récupération du Shabbat suivant: {e}")
            return None, None

    def create_image(self, times, parasha, parasha_hebrew, shabbat_end, candle_lighting):
        try:
            print("Ouverture du template...")
            print(f"Chemin du template : {self.template_path}")
            print(f"Chemin de la police : {self.font_path}")
            print(f"Chemin de la police Arial Bold : {self.arial_bold_path}")
            
            with Image.open(self.template_path) as img:
                print("Template ouvert avec succès.")
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(str(self.font_path), 30)

                # Coordonnées ajustées pour les horaires
                time_x = 120  # Position X des horaires
                time_positions = [
                    (time_x, 400, 'shir_hashirim'),
                    (time_x, 475, 'mincha_kabbalat'),
                    (time_x, 510, 'shacharit'),
                    (time_x, 550, 'mincha_gdola'),
                    (time_x, 590, 'parashat_hashavua'),
                    (time_x, 630, 'tehilim'),
                    (time_x, 670, 'shiur_rav'),
                    (time_x, 710, 'nashim'),
                    (time_x, 750, 'mincha_2'),
                    (time_x, 790, 'arvit')
                ]

                # Écrire les horaires sans les noms des activités
                for x, y, time_key in time_positions:
                    formatted_time = self.format_time(times[time_key])
                    draw.text((x, y), formatted_time, fill="black", font=font)

                # Ajouter l'heure de fin de Shabbat en face de "מוצאי שבת קודש"
                end_time_str = shabbat_end.strftime('%H:%M')
                draw.text((time_x, 830), end_time_str, fill="black", font=font)

                # Ajouter le nom de la Parasha en hébreu dans le carré en haut à gauche
                #parasha_hebrew_reversed = self.reverse_hebrew_text(parasha_hebrew)  # Inverser le texte           
                draw.text((300, 280), parasha_hebrew, fill="blue", font=self._arial_bold_font, anchor="mm")

                # Ajouter l'heure de כניסת שבת en haut
                draw.text((time_x, 440), candle_lighting, fill="black", font=font)

                # Récupérer l'heure de מנחה (en bas de la page, sous ראשון-חמישי)
                mincha_time = times['mincha_kabbalat']
                mincha_str = self.format_time(mincha_time)

                # Ajouter l'heure de שבת הבאה en face de מנחה (en bas de la page)
                next_shabbat_time = self.get_next_shabbat_time(shabbat_end.date())
                if next_shabbat_time[1]:  # Si l'heure est disponible
                    draw.text((time_x, 950), next_shabbat_time[1], fill="green", font=font)  # Coordonnées ajustées

                # Calculer l'heure de ערבית
                arvit_time = self.round_to_nearest_five(mincha_time + 45)
                arvit_str = self.format_time(arvit_time)

                # Ajouter l'heure de ערבית en face du mot ערבית
                draw.text((time_x, 990), arvit_str, fill="green", font=font)  # Coordonnées ajustées

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
        
        # Récupérer la date et l'heure du Shabbat suivant
        next_shabbat_date, next_shabbat_time = self.get_next_shabbat_time(shabbat_data['date'])
        
        row = {
            'תאריך': shabbat_data['date'].strftime('%d/%m/%Y'),
            'פרשה': shabbat_data['parasha'],
            'שיר השירים': self.format_time(times['shir_hashirim']),
            'כנסית שבת': shabbat_data['candle_lighting'],  # Colonne כנסית שבת déplacée ici
            'מנחה': self.format_time(times['mincha_kabbalat']),
            'שחרית': self.format_time(times['shacharit']),
            'מנחה גדולה': self.format_time(times['mincha_gdola']),
            'שיעור לנשים': self.format_time(times['nashim']),
            'תהילים לילדים': self.format_time(times['tehilim']),
            'שיעור פרשת השבוע': self.format_time(times['parashat_hashavua']),
            'שיעור עם הרב': self.format_time(times['shiur_rav']),
            'מנחה 2': self.format_time(times['mincha_2']),
            'ערבית מוצ"ש': self.format_time(times['arvit']),
            'מוצאי שבת קודש': shabbat_data['end'].strftime('%H:%M'),
            'שבת הבאה (Date)': next_shabbat_date if next_shabbat_date else "N/A",  # Date du Shabbat suivant
            'שבת הבאה (Heure)': next_shabbat_time if next_shabbat_time else "N/A"  # Heure du Shabbat suivant
        }

        try:
            # Charger ou créer le fichier Excel
            if excel_path.exists():
                with pd.ExcelWriter(str(excel_path), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    # Ajouter la nouvelle ligne à l'onglet Sheet1
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=writer.sheets['Sheet1'].max_row)
                    
                    # Recréer l'onglet שבתות השנה à partir des données intégrées
                    yearly_df = pd.DataFrame(self.yearly_shabbat_data)
                    yearly_df.to_excel(writer, sheet_name='שבתות השנה', index=False)
            else:
                # Créer un nouveau fichier Excel avec les deux onglets
                with pd.ExcelWriter(str(excel_path), engine='openpyxl') as writer:
                    df = pd.DataFrame([row])
                    df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    yearly_df = pd.DataFrame(self.yearly_shabbat_data)
                    yearly_df.to_excel(writer, sheet_name='שבתות השנה', index=False)
            
            print(f"Excel mis à jour: {excel_path}")
        except Exception as e:
            print(f"Erreur lors de la mise à jour de l'Excel: {e}")

    def generate(self):
        current_date = datetime.now()
        end_date = current_date + timedelta(days=14)  # Récupérer les données pour 2 semaines
        
        shabbat_times = self.get_hebcal_times(current_date, end_date)
        if not shabbat_times:
            print("Aucun horaire trouvé pour cette semaine")
            return

        shabbat = shabbat_times[0]
        times = self.calculate_times(shabbat['start'], shabbat['end'])
        
        image_path = self.create_image(times, shabbat['parasha'], shabbat['parasha_hebrew'], shabbat['end'], shabbat['candle_lighting'])
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
