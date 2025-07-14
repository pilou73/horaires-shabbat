#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from pathlib import Path
import pytz
from icalendar import Calendar, Event

# Référence Tekufat Tishri 5784 : 7 octobre 2023 à 21:39 (heure moyenne Jérusalem)
initial_jm = datetime(2023, 10, 7, 21, 39)

# Intervalle שיטת שמואל : 91 jours, 7 h 30 min
interval = timedelta(days=91, hours=7, minutes=30)

# Noms cycliques des tekufot
names = ["Tekufat Tishri", "Tekufat Tevet", "Tekufat Nisan", "Tekufat Tammuz"]

# Fuseau de Jérusalem
tz_jerusalem = pytz.timezone("Asia/Jerusalem")

# Préparation du Calendar ICS
cal = Calendar()
cal.add('prodid', '-//Tekufot 2025–2035 (שיטת שמואל)//')
cal.add('version', '2.0')

print("Génération des Tekufot 2025–2035 selon שיטת שמואל :")
print("--------------------------------------------------")

index = 0
while True:
    current_jm = initial_jm + index * interval
    year = current_jm.year
    if year > 2035:
        break

    if 2025 <= year <= 2035:
        name = names[index % 4]

        # On interprète current_jm comme heure locale de Jérusalem (évite DST automatique)
        dt_local = tz_jerusalem.localize(current_jm)

        # Ajustement manuel pour Tekufat Tevet (on retire l'heure d'hiver indésirable)
        if name == "Tekufat Tevet":
            dt_local = dt_local - timedelta(hours=1)

        # Affichage console
        print(f"{name} ({year}): {dt_local.strftime('%Y-%m-%d %H:%M')} (Israel)")

        # Création de l’événement ICS
        ev = Event()
        ev.add('summary', name)
        ev.add('dtstart', dt_local)
        ev.add('dtend', dt_local + timedelta(minutes=1))
        cal.add_component(ev)

    index += 1

# Sauvegarde du fichier .ics
output_path = Path("tekufa_2025_2035.ics")
with open(output_path, 'wb') as f:
    f.write(cal.to_ical())

print("✅ Fichier", output_path, "généré avec succès.")
