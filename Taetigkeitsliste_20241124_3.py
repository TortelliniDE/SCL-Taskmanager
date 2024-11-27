import calendar
import random
from datetime import date, datetime, timedelta

def create_task_list(year, month, tasks, first_name, last_name, mitarbeiternummer, wochenarbeitszeit, resturlaub):
    # Liste der Wochentage
    weekdays = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
    
    # Liste der Monatsnamen auf Deutsch
    german_months = [
        "Januar", "Februar", "März", "April", "Mai", "Juni", 
        "Juli", "August", "September", "Oktober", "November", "Dezember"
    ]
    
    # Gesetzliche Feiertage in Bayern (Jahr anpassbar)
    public_holidays_bayern = {
        "Neujahr": date(year, 1, 1),
        "Heilige Drei Könige": date(year, 1, 6),
        "Karfreitag": None,  # Dynamisch berechnet
        "Ostermontag": None,  # Dynamisch berechnet
        "Tag der Arbeit": date(year, 5, 1),
        "Christi Himmelfahrt": None,  # Dynamisch berechnet
        "Pfingstmontag": None,  # Dynamisch berechnet
        "Fronleichnam": None,  # Dynamisch berechnet
        "Mariä Himmelfahrt": date(year, 8, 15),
        "Tag der Deutschen Einheit": date(year, 10, 3),
        "Allerheiligen": date(year, 11, 1),
        "1. Weihnachtsfeiertag": date(year, 12, 25),
        "2. Weihnachtsfeiertag": date(year, 12, 26),
    }

    # Funktion zur Berechnung beweglicher Feiertage
    def calculate_movable_feasts():
        from datetime import timedelta
        # Berechnung des Ostersonntags nach der Methode von Gauß
        a = year % 19
        b = year // 100
        c = year % 100
        d = b // 4
        e = b % 4
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i = c // 4
        k = c % 4
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        month = (h + l - 7 * m + 114) // 31
        day = ((h + l - 7 * m + 114) % 31) + 1
        easter_sunday = date(year, month, day)
        
        # Karfreitag, Ostermontag, Christi Himmelfahrt, Pfingstmontag, Fronleichnam
        public_holidays_bayern["Karfreitag"] = easter_sunday - timedelta(days=2)
        public_holidays_bayern["Ostermontag"] = easter_sunday + timedelta(days=1)
        public_holidays_bayern["Christi Himmelfahrt"] = easter_sunday + timedelta(days=39)
        public_holidays_bayern["Pfingstmontag"] = easter_sunday + timedelta(days=50)
        public_holidays_bayern["Fronleichnam"] = easter_sunday + timedelta(days=60)

    # Berechnung der beweglichen Feiertage
    calculate_movable_feasts()
    
    # Anzahl der Tage im Monat
    days_in_month = calendar.monthrange(year, month)[1]

    # Ausgabe der Tätigkeitsliste mit einer visuell hervorgehobenen Überschrift
    print() 
    print(f"\033[1;34m\033[1mTätigkeitsliste für {first_name} {last_name}\033[0m")
    print() 
    
    print(f"Mitarbeiternummer: {mitarbeiternummer}")
    print(f"Wochenarbeitszeit: {wochenarbeitszeit} h/Wo")
    print(f"Resturlaub: {resturlaub} Tage")
    
    # Aktuelles Datum mit "Stand"
    current_date = datetime.now().strftime("%d.%m.%Y")
    print(f"\nStand: {current_date}")
    
    # Ausgabe des Monats und Jahres
    print(f"\n\033[1;34m\033[1mMonat: {german_months[month - 1]} {year}\033[0m\n")
    
    # Tabellenüberschrift
    print(f"{'Datum':<12} {'Dauer':<8} {'Wochentag':<12} {'Task'}")
    print("="*50)

    # Berechnung der täglichen Arbeitszeit und Verteilung auf Wochentage
    work_hours_per_week = wochenarbeitszeit
    total_week_days = 5  # Montag bis Freitag
    daily_hours = []

    # Zufällige Arbeitszeiten für Montag bis Freitag generieren, insgesamt mindestens 40h/Woche
    total_hours_per_week = 0
    for i in range(total_week_days):
        min_hours = 6 * 60  # 6:00 in Minuten
        max_hours = 9 * 60 + 55  # 9:55 in Minuten
        random_minutes = random.randint(min_hours, max_hours)
        daily_hours.append(random_minutes)
        total_hours_per_week += random_minutes
    
    # Wenn die gesamte Arbeitszeit weniger als 40h ist, passen wir die Stunden an
    deficit_hours = work_hours_per_week * 60 - total_hours_per_week
    if deficit_hours > 0:
        # Verteilen der verbleibenden Stunden gleichmäßig
        for i in range(deficit_hours // total_week_days):
            for j in range(total_week_days):
                daily_hours[j] += 1
                deficit_hours -= 1
                if deficit_hours <= 0:
                    break
    
    # Ausgabe der Arbeitszeiten für den Monat
    for day in range(1, days_in_month + 1):
        weekday_index = calendar.weekday(year, month, day)
        weekday = weekdays[weekday_index]
        current_date = date(year, month, day)
        
        # Prüfen, ob das aktuelle Datum ein Feiertag ist
        holiday = next((name for name, holiday_date in public_holidays_bayern.items() if holiday_date == current_date), None)
        
        # Aufgabenlogik
        if holiday:
            task_entry = holiday
            work_duration = "0:00"
        elif weekday_index in [5, 6]:  # Samstag oder Sonntag
            task_entry = "Wochenende"
            work_duration = "0:00"
        else:  # Wochentage (Montag - Freitag)
            task_entry = ', '.join(random.sample(tasks, random.randint(2, min(5, len(tasks)))))  # 2-5 Aufgaben zufällig auswählen
            work_minutes = daily_hours[weekday_index % total_week_days]  # Die Stunden aus daily_hours entnehmen
            hours = work_minutes // 60
            minutes = work_minutes % 60
            work_duration = f"{hours:01}:{minutes:02}"
        
        # Leerzeile vor jedem Montag
        if weekday_index == 0 and day != 1:
            print()
        
        # Zeilenformatierung: Gelb für Feiertage und Wochenenden
        if holiday or weekday_index in [5, 6]:
            print(f"\033[1;33m{day:02d}.{month:02d}. {work_duration:<8} {weekday:<12} {task_entry}\033[0m")
        else:
            print(f"{day:02d}.{month:02d}. {work_duration:<8} {weekday:<12} {task_entry}")


# Eingabe von Vorname, Nachname und Monat
first_name = "Sylvia" #input("Gib den Vornamen des Mitarbeiters ein: ").strip()
last_name = "Prigge" #input("Gib den Nachnamen des Mitarbeiters ein: ").strip()
month = int(input("Gib den Monat ein (1-12): "))

# Aktuelles Jahr
year = datetime.now().year

# Neue Variablen
Mitarbeiternummer = "47110815"
Wochenarbeitszeit = 40
Resturlaub = 7.5

# Neue Liste der Aufgaben
tasks = [
    'Einarbeitung',
    'Angebotseinholungen',
    'Bestellungen',
    'Recherche'
]

create_task_list(year, month, tasks, first_name, last_name, Mitarbeiternummer, Wochenarbeitszeit, Resturlaub)









