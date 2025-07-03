from datetime import datetime
from zmanim.hebrew_calendar.jewish_calendar import JewishCalendar

HEBREW_MONTHS = {
    1: 'ניסן',
    2: 'אייר',
    3: 'סיון',
    4: 'תמוז',
    5: 'אב',
    6: 'אלול',
    7: 'תשרי',
    8: 'חשוון',
    9: 'כסלו',
    10: 'טבת',
    11: 'שבט',
    12: 'אדר',
    13: 'אדר ב׳',  # שנה מעוברת
}

def get_jewish_month_name_hebrew(jewish_month, jewish_year):
    # שנה מעוברת: אדר א', אדר ב'
    if jewish_month == 12 and JewishCalendar.is_jewish_leap_year(jewish_year):
        return 'אדר א׳'
    if jewish_month == 13:
        return 'אדר ב׳'
    return HEBREW_MONTHS.get(jewish_month, 'חודש לא ידוע')

def int_to_hebrew(num):
    # פשטות: מספרים עד 5999
    hebrew_digits = {
        1: 'א', 2: 'ב', 3: 'ג', 4: 'ד', 5: 'ה', 6: 'ו', 7: 'ז', 8: 'ח', 9: 'ט',
        10: 'י', 20: 'כ', 30: 'ל', 40: 'מ', 50: 'נ', 60: 'ס', 70: 'ע', 80: 'פ', 90: 'צ',
        100: 'ק', 200: 'ר', 300: 'ש', 400: 'ת'
    }
    if num >= 1000:
        thousands = num // 1000
        rest = num % 1000
        return f"{''.join([hebrew_digits.get(thousands, str(thousands))])}׳{int_to_hebrew(rest)}"
    result = ''
    for value in [400, 300, 200, 100, 90, 80, 70, 60, 50, 40, 30, 20, 10]:
        while num >= value:
            result += hebrew_digits[value]
            num -= value
    if num == 15:  # ט״ו
        result += 'ט״ו'
    elif num == 16:  # ט״ז
        result += 'ט״ז'
    else:
        if num > 0:
            result += hebrew_digits[num]
    # גרשיים בסוף
    if len(result) > 1:
        result = result[:-1] + '״' + result[-1]
    elif len(result) == 1:
        result += '׳'
    return result

def calculate_molad_hebrew(gregorian_date):
    try:
        jewish_cal = JewishCalendar(gregorian_date)
        molad_obj = jewish_cal.molad()
        hour = molad_obj.molad_hours
        minute = molad_obj.molad_minutes
        chalakim = molad_obj.molad_chalakim
        month_name = get_jewish_month_name_hebrew(jewish_cal.jewish_month, jewish_cal.jewish_year)
        year_he = int_to_hebrew(jewish_cal.jewish_year)
        molad_str = f"{hour} שעות {minute} דקות {chalakim} חלקים"
        return {
            "gregorian": gregorian_date.strftime("%Y-%m-%d"),
            "hebrew_month": month_name,
            "molad": molad_str,
            "hebrew_year": year_he
        }
    except Exception as e:
        return {"error": f"שגיאה בחישוב מולד: {e}"}

if __name__ == "__main__":
    test_date = datetime(2025, 5, 20)
    result = calculate_molad_hebrew(test_date)
    if "error" in result:
        print(result["error"])
    else:
        print("📅 תאריך לועזי:", result["gregorian"])
        print("🌙 חודש עברי:", result["hebrew_month"])
        print("🕰️  מולד:", result["molad"])
        print("🗓️  שנה עברית:", result["hebrew_year"])