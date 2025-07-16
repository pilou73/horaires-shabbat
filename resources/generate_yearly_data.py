import requests
from datetime import datetime, timedelta
import re
import json

def fetch_shabbatot(start, end):
    url = "https://www.hebcal.com/hebcal"
    params = {
        "v": 1,
        "cfg": "json",
        "maj": "on",
        "ss": "on",
        "mf": "on",
        "geonameid": "293397",  # Ramat Gan
        "start": start,
        "end": end,
        "lg": "he",
        "s": "on",
        "c": "on",
        "H": "on",
    }
    r = requests.get(url, params=params)
    r.raise_for_status()
    return r.json()

def extract_time_from_title(title):
    m = re.search(r'(\d{1,2}):(\d{2})', title)
    if m:
        return f"{m.group(1).zfill(2)}:{m.group(2)}"
    return ''

def parse_events(data):
    candles_dict = {}
    havdalah_dict = {}
    for item in data["items"]:
        if item["category"] == "candles":
            date = item["date"][:10]
            candles_dict[date] = extract_time_from_title(item["title"])
        elif item["category"] == "havdalah":
            date = item["date"][:10]
            havdalah_dict[date] = extract_time_from_title(item["title"])
    shabbat_list = []
    for item in data["items"]:
        if item["category"] == "parashat":
            shabbat_date = item["date"]
            parasha = item["hebrew"]
            friday_dt = datetime.fromisoformat(shabbat_date) - timedelta(days=1)
            friday_str = friday_dt.strftime("%Y-%m-%d")
            candle = candles_dict.get(friday_str, "")
            havdalah = havdalah_dict.get(shabbat_date, "")
            if candle:
                day_str = f"{friday_str} {candle}:00"
            else:
                day_str = f"{friday_str} 00:00:00"
            shabbat_list.append({
                'day': day_str,
                'פרשה': parasha,
                'כנסית שבת': candle,
                'צאת שבת': havdalah
            })
    return shabbat_list

if __name__ == "__main__":
    start = "2029-09-01"
    end = "2030-09-30"
    data = fetch_shabbatot(start, end)
    shabbat_data = parse_events(data)

    # Écriture TXT (format python dict)
    with open("yearly_shabbat_data_5790.txt", "w", encoding="utf-8") as f:
        for entry in shabbat_data:
            f.write(f"{entry},\n")

    # Écriture JSON
    with open("yearly_shabbat_data_5790.json", "w", encoding="utf-8") as fjson:
        json.dump(shabbat_data, fjson, ensure_ascii=False, indent=2)

    print("✅ yearly_shabbat_data_5790.txt et yearly_shabbat_data_5790.json générés.")