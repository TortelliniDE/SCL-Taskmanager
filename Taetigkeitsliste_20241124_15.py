import os
import calendar
import random
from datetime import date, datetime
import pandas as pd  # Für den Excel-Import
from tkinter import Tk, Listbox, Button, Label, Scrollbar, MULTIPLE, EXTENDED, END, messagebox


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
        
        public_holidays_bayern["Karfreitag"] = easter_sunday - timedelta(days=2)
        public_holidays_bayern["Ostermontag"] = easter_sunday + timedelta(days=1)
        public_holidays_bayern["Christi Himmelfahrt"] = easter_sunday + timedelta(days=39)
        public_holidays_bayern["Pfingstmontag"] = easter_sunday + timedelta(days=50)
        public_holidays_bayern["Fronleichnam"] = easter_sunday + timedelta(days=60)

    calculate_movable_feasts()
    
    days_in_month = calendar.monthrange(year, month)[1]
    workdays_in_month = [
        day for day in range(1, days_in_month + 1) 
        if calendar.weekday(year, month, day) < 5 and date(year, month, day) not in public_holidays_bayern.values()
    ]

    total_soll_minutes = 9600
    daily_minutes = [random.randint(360, 540) for _ in range(len(workdays_in_month))]
    correction = total_soll_minutes - sum(daily_minutes)
    for i in range(len(daily_minutes)):
        daily_minutes[i] += correction // len(daily_minutes)
    daily_minutes[-1] += correction % len(daily_minutes)

    print(f"\n\n\033[1;34m\033[1mTätigkeitsliste für {first_name} {last_name} im Monat {german_months[month - 1]} {year}\033[0m\n")
    print(f"Mitarbeiternummer: {mitarbeiternummer}")
    print(f"Wochenarbeitszeit: {wochenarbeitszeit} h/Wo")
    print(f"Resturlaub: {resturlaub} Tage")
    
    # Aktuelles Datum mit "Stand"
    current_date = datetime.now().strftime("%d.%m.%Y")
    print(f"Stand: {current_date} \n")
    
    print(f"{'Datum':<7} {'Dauer':<8} {'Wochentag':<12} {'Tätigkeiten'}")
    print("="*120)

    total_minutes_worked = 0
    workday_idx = 0

    for day in range(1, days_in_month + 1):
        weekday_index = calendar.weekday(year, month, day)
        weekday = weekdays[weekday_index]
        current_date = date(year, month, day)
        holiday_name = next((name for name, holiday_date in public_holidays_bayern.items() if holiday_date == current_date), None)

        # Leerzeile vor jedem Montag
        if weekday_index == 0 and day != 1:
            print()

        if holiday_name:
            task_entry = holiday_name
            work_hours = 0
            work_minutes = 0
            print(f"\033[1;33m{day:02d}.{month:02d}.  0:00     {weekday:<12} {task_entry}\033[0m")
        elif weekday_index in [5, 6]:
            task_entry = "Wochenende"
            work_hours = 0
            work_minutes = 0
            print(f"\033[1;33m{day:02d}.{month:02d}.  0:00     {weekday:<12} {task_entry}\033[0m")
        else:
            task_entry = ', '.join(random.sample(tasks, random.randint(2, 3)))
            work_minutes = daily_minutes[workday_idx]
            work_hours = work_minutes // 60
            work_minutes %= 60
            total_minutes_worked += work_hours * 60 + work_minutes
            workday_idx += 1
            print(f"{day:02d}.{month:02d}.  {work_hours}:{work_minutes:02d}     {weekday:<12} {task_entry}")
    
    total_hours_worked = total_minutes_worked // 60
    total_minutes_worked %= 60
    print(f"\nSollarbeitszeit im Monat: 160:00 h")
    print(f"Geleistete Arbeitszeit im Monat: {total_hours_worked}:{total_minutes_worked:02d} h")

# Mitarbeiterdaten aus Excel laden
def load_employees_from_excel(file_path):
    df = pd.read_excel(file_path)
    if not {'Vorname', 'Nachname', 'Mitarbeiternummer', 'Resturlaub'}.issubset(df.columns):
        raise ValueError("Die Excel-Datei muss die Spalten 'Vorname', 'Nachname', 'Mitarbeiternummer' und 'Resturlaub' enthalten.")
    return df

# Mitarbeiter auswählen
def choose_employee(employees_df):
    print("\nVerfügbare Mitarbeiter:")
    for idx, row in employees_df.iterrows():
        print(f"{idx + 1}: {row['Vorname']} {row['Nachname']}")
    while True:
        try:
            choice = int(input("\nWähle die Nummer des Mitarbeiters: "))
            if 1 <= choice <= len(employees_df):
                selected_employee = employees_df.iloc[choice - 1]
                return (selected_employee['Vorname'], selected_employee['Nachname'], 
                        selected_employee['Mitarbeiternummer'], selected_employee['Resturlaub'])
            else:
                print("Ungültige Auswahl. Bitte eine gültige Nummer eingeben.")
        except ValueError:
            print("Bitte eine gültige Nummer eingeben.")


# Neue GUI-Funktionen
def load_tasks_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'Task' not in df.columns:
            raise ValueError("Die Excel-Datei muss eine Spalte 'Task' enthalten.")
        return df['Task'].dropna().tolist()
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Laden der Excel-Datei: {e}")
        return []

def save_tasks_to_excel(tasks, file_path):
    try:
        df = pd.DataFrame({'Task': tasks})
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Erfolg", f"Tasks wurden erfolgreich in '{file_path}' gespeichert.")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Speichern der Excel-Datei: {e}")

def update_selected_tasks():
    selected = list(task_listbox.curselection())
    for index in selected:
        task = task_listbox.get(index)
        if task not in input_task_listbox.get(0, END):
            input_task_listbox.insert(END, task)

def clear_input_list():
    input_task_listbox.delete(0, END)

def remove_selected_input_tasks():
    selected = list(input_task_listbox.curselection())
    for index in reversed(selected):
        input_task_listbox.delete(index)

def save_input_list():
    current_date = datetime.now().strftime("%Y.%m.%d")
    tasks = list(input_task_listbox.get(0, END))
    if tasks:
        save_tasks_to_excel(tasks, f"Tasks_{current_date}.xlsx")
    else:
        messagebox.showwarning("Warnung", "Keine Tasks in der Eingabeliste zum Speichern vorhanden.")

# GUI-Setup
root = Tk()
root.title("Tätigkeiten Liste")
root.geometry("800x600")

# Labels
Label(root, text="Verfügbare Tasks (aus Projektliste_SCL.xlsx)").grid(row=0, column=0, padx=10, pady=5)
Label(root, text="Eingabeliste").grid(row=0, column=2, padx=10, pady=5)

# Listbox für verfügbare Tasks
task_listbox = Listbox(root, selectmode=EXTENDED, width=40, height=25)
task_listbox.grid(row=1, column=0, padx=10, pady=5)

task_scrollbar = Scrollbar(root, orient="vertical", command=task_listbox.yview)
task_scrollbar.grid(row=1, column=1, sticky="ns")
task_listbox.config(yscrollcommand=task_scrollbar.set)

Button(root, text="Ausgewählte hinzufügen →", command=update_selected_tasks).grid(row=1, column=1, padx=10)
Button(root, text="Liste Leeren", command=clear_input_list).grid(row=2, column=2, pady=5)
Button(root, text="Ausgewählte entfernen", command=remove_selected_input_tasks).grid(row=3, column=2, pady=5)

input_task_listbox = Listbox(root, selectmode=MULTIPLE, width=40, height=25)
input_task_listbox.grid(row=1, column=2, padx=10, pady=5)

Button(root, text="Speichern", command=save_input_list, width=20).grid(row=4, column=2, pady=20)

task_file = "Projekte_SCL.xlsx"
if os.path.exists(task_file):
    tasks = load_tasks_from_excel(task_file)
    for task in tasks:
        task_listbox.insert(END, task)
else:
    messagebox.showerror("Fehler", f"Die Datei '{task_file}' wurde nicht gefunden.")

root.mainloop()



# Aufgaben aus Excel laden
def load_tasks_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Task'].dropna().tolist()

# Auswahl der Excel-Datei
def choose_excel_file():
    xlsx_files = [file for file in os.listdir() if file.endswith('.xlsx')]
    if not xlsx_files:
        print("Keine .xlsx-Dateien im aktuellen Verzeichnis gefunden.")
        return None
    print("\nVerfügbare Excel-Dateien:")
    for idx, file in enumerate(xlsx_files, 1):
        print(f"{idx}: {file}")
    while True:
        try:
            current_date = datetime.now().strftime("%Y_%m_%d")
            choice = int(input (f"\nWähle die Nr. der Datei Projekte_{current_date}"))
            if 1 <= choice <= len(xlsx_files):
                return xlsx_files[choice - 1]
            else:
                print("Ungültige Auswahl. Bitte eine Nummer aus der Liste wählen.")
        except ValueError:
            print("Bitte eine gültige Nummer eingeben.")

# Hauptskript
employees_file = "Mitarbeiter.xlsx"
if not os.path.exists(employees_file):
    print(f"Die Datei {employees_file} wurde nicht gefunden.")
    exit()

employees_df = load_employees_from_excel(employees_file)
first_name, last_name, Mitarbeiternummer, Resturlaub = choose_employee(employees_df)

month = int(input("\nGib den Monat ein (1-12): "))
tasks_file = choose_excel_file()
if not tasks_file:
    print("Keine Datei ausgewählt. Das Programm wird beendet.")
    exit()

###hier
tasks = load_tasks_from_excel(tasks_file)
create_task_list(datetime.now().year, month, tasks, first_name, last_name, Mitarbeiternummer, 40, Resturlaub)

print(f"\nDie Tätigkeitsliste {month}/{datetime.now().year} für {first_name} {last_name} wurde erstellt.\n")

