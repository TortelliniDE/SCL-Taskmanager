import calendar
import random
from datetime import date, datetime

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
    
    # Arbeitstage im Monat (Montag bis Freitag, keine Feiertage)
    workdays_in_month = [
        day for day in range(1, days_in_month + 1) 
        if calendar.weekday(year, month, day) < 5 and date(year, month, day) not in public_holidays_bayern.values()
    ]

    # Berechnung der Sollarbeitszeit im Monat (160 Stunden = 9600 Minuten)
    total_soll_minutes = 9600

    # Zufällige Verteilung der Arbeitszeit auf die Arbeitstage
    daily_minutes = [random.randint(360, 540) for _ in range(len(workdays_in_month))]
    total_assigned_minutes = sum(daily_minutes)

    # Korrektur, um genau 9600 Minuten zu erreichen
    correction = total_soll_minutes - total_assigned_minutes
    for i in range(len(daily_minutes)):
        daily_minutes[i] += correction // len(daily_minutes)
    daily_minutes[-1] += correction % len(daily_minutes)

    # Ausgabe der Tätigkeitsliste mit einer visuell hervorgehobenen Überschrift
    print()
    print(f"\033[1;34m\033[1mTätigkeitsliste für {first_name} {last_name}\033[0m\n")
    print(f"Mitarbeiternummer: {mitarbeiternummer}")
    print(f"Wochenarbeitszeit: {wochenarbeitszeit} h/Wo")
    print(f"Resturlaub: {resturlaub} Tage")
    
    # Aktuelles Datum mit "Stand"
    current_date = datetime.now().strftime("%d.%m.%Y")
    print(f"Stand: {current_date}")
    
    # Ausgabe des Monats und Jahres mit "Monat:" in blau und fett
    print(f"\n\033[1;34m\033[1mMonat: {german_months[month - 1]} {year}\033[0m\n")
    
    # Tabellenüberschrift
    print(f"{'Datum':<6} {'Dauer':<8} {'Wochentag':<12} {'Task'}")
    print("="*115)

    total_minutes_worked = 0
    workday_idx = 0  # Index für Arbeitstage

    # Generieren der täglichen Aufgaben
    for day in range(1, days_in_month + 1):
        weekday_index = calendar.weekday(year, month, day)
        weekday = weekdays[weekday_index]
        current_date = date(year, month, day)
        
        # Prüfen, ob es ein Feiertag ist
        holiday_name = next((name for name, holiday_date in public_holidays_bayern.items() if holiday_date == current_date), None)

        if holiday_name:  # Feiertag
            task_entry = holiday_name
            work_hours = 0
            work_minutes = 0
        elif weekday_index in [5, 6]:  # Wochenende
            task_entry = "Wochenende"
            work_hours = 0
            work_minutes = 0
        else:  # Arbeitstag
            task_entry = ', '.join(random.sample(tasks, random.randint(2, 5)))
            work_minutes = daily_minutes[workday_idx]
            work_hours = work_minutes // 60
            work_minutes %= 60
            total_minutes_worked += work_hours * 60 + work_minutes
            workday_idx += 1

        work_duration = f"{work_hours}:{work_minutes:02d}"

        # Leerzeile vor jedem Montag
        if weekday_index == 0 and day != 1:
            print()

        # Gelbe Formatierung für Wochenenden und Feiertage
        if holiday_name or weekday_index in [5, 6]:
            print(f"\033[1;33m{day:02d}.{month:02d}. {work_duration:<8} {weekday:<12} {task_entry}\033[0m")
        else:
            print(f"{day:02d}.{month:02d}. {work_duration:<8} {weekday:<12} {task_entry}")
    
    # Ausgabe der Sollarbeitszeit
    print(f"\nSollarbeitszeit im Monat: 160:00 h")
    
    # Gesamtarbeitszeit
    total_hours_worked = total_minutes_worked // 60
    total_minutes_worked %= 60
    print(f"geleistete Arbeitszeit im Monat: {total_hours_worked}:{total_minutes_worked:02d} h")


# Eingabe von Vorname, Nachname und Monat
first_name = "Sylvia"
last_name = "Prigge"
month = int(input("Gib den Monat ein (1-12): "))

# Aktuelles Jahr
year = datetime.now().year

# Variablen
Mitarbeiternummer = "47110815"
Wochenarbeitszeit = 40  
Resturlaub = 7.5  

# Liste der Aufgaben
tasks = [
    'Einarbeitung',
    'Angebotseinholungen',
    'Bestellungen',
    'Recherche',
    'Rechnungsprüfung',
    'SCL-Büro',
    'RG-Prüfung',
    'Projektarbeit',
    'administrative Tätigkeiten',
    'Teko'
]

# Funktion aufrufen
create_task_list(year, month, tasks, first_name, last_name, Mitarbeiternummer, Wochenarbeitszeit, Resturlaub)



