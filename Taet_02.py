import os
import pandas as pd
from datetime import datetime
from create_task_list import create_task_list
from gui import setup_gui, load_tasks_from_excel  # GUI-Funktionen importieren

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
            choice = int(input("\nFür welchen Mitarbeiters soll die Tätigkeitsliste erstellt werden?  "))
            if 1 <= choice <= len(employees_df):
                selected_employee = employees_df.iloc[choice - 1]
                return (selected_employee['Vorname'], selected_employee['Nachname'], 
                        selected_employee['Mitarbeiternummer'], selected_employee['Resturlaub'])
            else:
                print("Ungültige Auswahl. Bitte eine gültige Nummer eingeben.")
        except ValueError:
            print("Bitte eine gültige Nummer eingeben.")

# Hauptskript
def main():
    employees_file = "Mitarbeiter.xlsx"
    if not os.path.exists(employees_file):
        print(f"Die Datei {employees_file} wurde nicht gefunden.")
        exit()

    employees_df = load_employees_from_excel(employees_file)
    first_name, last_name, Mitarbeiternummer, Resturlaub = choose_employee(employees_df)

    month = int(input("\nGib den Monat ein für den die Tätigkeitsliste erstellt werden soll (1-12): "))
    year = datetime.now().year

    # GUI-Setup und Task-Auswahl
    task_file = "Projekte_SCL.xlsx"
    if not os.path.exists(task_file):
        print(f"Die Datei '{task_file}' wurde nicht gefunden.")
        exit()

    # Startet die GUI, um Aufgaben auszuwählen
    setup_gui(task_file)

    # Überprüfen, ob Tasks ausgewählt wurden
    tasks_tmp = "Tasks_tmp.xlsx"
    if not os.path.exists(tasks_tmp):
        print("Es wurden keine Tasks ausgewählt. Das Programm wird beendet.")
        exit()

    # Tasks aus der temporären Datei laden
    tasks = load_tasks_from_excel(tasks_tmp)

    # Tätigkeitsliste erstellen
    create_task_list(year, month, tasks, first_name, last_name, Mitarbeiternummer, 40, Resturlaub)

    # Temporäre Datei löschen
    os.remove(tasks_tmp)
    print(f"\nDie Tätigkeitsliste {month}/{year} für {first_name} {last_name} wurde erstellt.\n")

if __name__ == "__main__":
    main()
