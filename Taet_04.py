import os
import pandas as pd
from datetime import datetime
from create_task_list import create_task_list
from gui import setup_gui, load_tasks_from_excel  # GUI-Funktionen importieren

# Flag, ob Mitarbeiter aus Datei geladen werden soll
employer_from_file = False  # Standardwert: False


# Hauptskript
def main():
    if employer_from_file:
        from employee_selector import load_employees_from_excel, choose_employee  # Mitarbeiterfunktionen laden

        employees_file = "Mitarbeiter.xlsx"
        if not os.path.exists(employees_file):
            print(f"Die Datei {employees_file} wurde nicht gefunden.")
            exit()

        employees_df = load_employees_from_excel(employees_file)
        first_name, last_name, Mitarbeiternummer, Resturlaub = choose_employee(employees_df)
    else:
        # Standard-Mitarbeiterdaten
        first_name, last_name = "Sylvia", "Prigge"
        Mitarbeiternummer = "SCL-4711"  # Beispielnummer
        Resturlaub = 7.5  # Beispielwert

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
        print("Es wurden keine Tasks ausgewählt. Das Progamm wird beendet.")
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
