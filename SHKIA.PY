# Si nécessaire, décommentez la ligne suivante pour installer les dépendances en local
# !pip install convertdate astral pandas openpyxl

import datetime
import pandas as pd
from astral import LocationInfo
from astral.sun import sun
import convertdate.hebrew

# Paramètres pour l'année hébraïque תשפ"ה (5785)
hebrew_year = 5785

# Détermination de la date grégorienne du Rosh Hashanah (Tishri 1, mois 7) de l'année hébraïque
start_date_tuple = convertdate.hebrew.to_gregorian(hebrew_year, 7, 1)
start_date = datetime.date(*start_date_tuple)

# Détermination de la date grégorienne du Rosh Hashanah de l'année suivante
next_start_date_tuple = convertdate.hebrew.to_gregorian(hebrew_year + 1, 7, 1)
next_start_date = datetime.date(*next_start_date_tuple)

# La fin de l'année hébraïque est le jour précédent le prochain Rosh Hashanah
end_date = next_start_date - datetime.timedelta(days=1)

# Préparation de la localisation pour Ramat Gan, Israël
ramat_gan = LocationInfo("Ramat Gan", "Israel", "Asia/Jerusalem", 32.0680, 34.8248)

# Liste pour stocker les résultats
data = []

# Parcours de la période
current_date = start_date
while current_date <= end_date:
    # En Python, weekday() retourne 6 pour dimanche (lundi=0)
    if current_date.weekday() == 6:  # Si c'est un dimanche
        # Calcul de l'heure de coucher de soleil pour le dimanche
        s_sunday = sun(ramat_gan.observer, date=current_date, tzinfo=ramat_gan.timezone)
        sunset_sunday = s_sunday['sunset'].strftime("%H:%M")
        
        # Calcul du jeudi qui suit : dimanche + 4 jours
        thursday_date = current_date + datetime.timedelta(days=4)
        if thursday_date < start_date or thursday_date > end_date:
            sunset_thursday = ""  # On laisse vide si le jeudi n'est pas dans l'année hébraïque
        else:
            s_thursday = sun(ramat_gan.observer, date=thursday_date, tzinfo=ramat_gan.timezone)
            sunset_thursday = s_thursday['sunset'].strftime("%H:%M")
        
        # Ajout des données dans la liste
        data.append({
            "תאריך": current_date.strftime("%Y-%m-%d"),
            "שקיעה Dimanche": sunset_sunday,
            "שקיעה Jeudi": sunset_thursday
        })
    current_date += datetime.timedelta(days=1)

# Création du DataFrame et export vers Excel
df = pd.DataFrame(data)
excel_filename = "hebrew_year_tashpah_sundays_and_following_thursdays_sunset.xlsx"
df.to_excel(excel_filename, index=False)

print(f"Fichier Excel '{excel_filename}' créé avec succès.")
