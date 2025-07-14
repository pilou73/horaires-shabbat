#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from pathlib import Path
import pytz
from icalendar import Calendar, Event, Alarm

# Référence Tekufat Tishri 5784 (heure moyenne de Jérusalem) :
initial_jm = datetime(2023, 10, 7, 21, 39)

# Intervalle selon שיטת שמואל : 91 jours, 7 h 30 min
interval = timedelta(days=91, hours=7, minutes=30)

# Noms cycliques des tekufot
names = ["Tekufat Tishri", "Tekufat Tevet", "Tekufat Nisan", "Tekufat Tammuz"]

# Fuseau de Jérusalem
tz_jerusalem = pytz.timezone("Asia/Jerusalem")

# Préparation du calendrier iCalendar
cal = Calendar()
cal.add('prodid', '-//Tekufot 2025–2035 (שיטת שמואל)//')
cal.add('version', '2.0')

print("Génération des Tekufot 2025–2035 avec rappels 30 min avant :")
print("-------------------------------------------------------------")

index = 0
while True:
    current_jm = initial_jm + index * interval
    year = current_jm.year
    if year > 2035:
        break

    if 2025 <= year <= 2035:
        name = names[index % 4]
        dt_local = tz_jerusalem.localize(current_jm)

        # Corriger Tekufat Tevet si besoin (retenue dans l'heure d'hiver)
        if name == "Tekufat Tevet":
            dt_local -= timedelta(hours=1)

        summary = f"{name} {dt_local.year}"
        print(f"{summary}: {dt_local.strftime('%Y-%m-%d %H:%M')} (Israel)")

        # Créer l'événement
        ev = Event()
        ev.add('summary', summary)
        ev.add('dtstart', dt_local)
        ev.add('dtend', dt_local + timedelta(minutes=1))

        # ➕ Ajouter rappel VALARM 30 minutes avant (DISPLAY)
        alarm = Alarm()
        alarm.add('ACTION', 'DISPLAY')
        alarm.add('DESCRIPTION', f"Tekufa {summary} dans 30 minutes")
        alarm.add('TRIGGER', timedelta(minutes=-30))
        ev.add_component(alarm)

        cal.add_component(ev)

    index += 1

# Sauvegarder le fichier ICS
output = Path("tekufa_2025_2035.ics")
with open(output, 'wb') as f:
    f.write(cal.to_ical())

print(f"✅ Fichier « {output} » généré avec rappels.")
