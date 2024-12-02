import os
import pandas as pd
from datetime import datetime
from create_task_list import create_task_list
from gui import setup_gui, load_tasks_from_excel  # GUI-Funktionen importieren
from logging_util import log_messages

# Flag, ob Mitarbeiter aus Datei geladen werden soll
employer_from_file = False  # Standardwert: False


# Speichern der Logs in Excel
log_file = "Logs.xlsx"
log_df = pd.DataFrame({"Logs": log_messages})
log_df.to_excel(log_file, index=False)
print(f"Die Logs wurden in '{log_file}' gespeichert.")

# Liste für Log-Nachrichten
log_messages = []

# Angepasste Print-Funktion
def log_and_print(message):
    print(message)
    log_messages.append(message)

# Hauptskript
def main():
    if employer_from_file:
        from employee_selector import load_employees_from_excel, choose_employee  # Mitarbeiterfunktionen laden

        employees_file = "Mitarbeiter.xlsx"
        if not os.path.exists(employees_file):
            log_and_print(f"Die Datei {employees_file} wurde nicht gefunden.")
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
        log_and_print(f"Die Datei '{task_file}' wurde nicht gefunden.")
        exit()

    # Startet die GUI, um Aufgaben auszuwählen
    setup_gui(task_file)

    # Überprüfen, ob Tasks ausgewählt wurden
    tasks_tmp = "Tasks_tmp.xlsx"
    if not os.path.exists(tasks_tmp):
        log_and_print("Es wurden keine Tasks ausgewählt. Das Progamm wird beendet.")
        exit()

    # Tasks aus der temporären Datei laden
    tasks = load_tasks_from_excel(tasks_tmp)

    # Tätigkeitsliste erstellen
    create_task_list(year, month, tasks, first_name, last_name, Mitarbeiternummer, 40, Resturlaub)

    # Temporäre Datei löschen
    os.remove(tasks_tmp)
    log_and_print(f"\nDie Tätigkeitsliste {month}/{year} für {first_name} {last_name} wurde erstellt.\n")

    # Log-Nachrichten in Excel speichern
    log_file = "Log_Output.xlsx"
    log_df = pd.DataFrame({"Logs": log_messages})
    log_df.to_excel(log_file, index=False)
    log_and_print(f"Die Log-Nachrichten wurden in '{log_file}' gespeichert.")


from logging_util import log_messages
print("\n".join(log_messages))  # Zeigt alle gesammelten Logs an


if __name__ == "__main__":
    main()
