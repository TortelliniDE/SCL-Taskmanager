import pandas as pd
from logging_util import log_and_print  # Log-Funktion importieren

# Mitarbeiterdaten aus Excel laden
def load_employees_from_excel(file_path):
    df = pd.read_excel(file_path)
    if not {'Vorname', 'Nachname', 'Mitarbeiternummer', 'Resturlaub'}.issubset(df.columns):
        raise ValueError("Die Excel-Datei muss die Spalten 'Vorname', 'Nachname', 'Mitarbeiternummer' und 'Resturlaub' enthalten.")
    return df

# Mitarbeiter auswählen
def choose_employee(employees_df):
    log_and_print("\nVerfügbare Mitarbeiter:")
    for idx, row in employees_df.iterrows():
        log_and_print(f"{idx + 1}: {row['Vorname']} {row['Nachname']}")
    while True:
        try:
            choice = int(input("\nFür welchen Mitarbeiters soll die Tätigkeitsliste erstellt werden?  "))
            if 1 <= choice <= len(employees_df):
                selected_employee = employees_df.iloc[choice - 1]
                return (selected_employee['Vorname'], selected_employee['Nachname'], 
                        selected_employee['Mitarbeiternummer'], selected_employee['Resturlaub'])
            else:
                log_and_print("Ungültige Auswahl. Bitte eine gültige Nummer eingeben.")
        except ValueError:
            log_and_print("Bitte eine gültige Nummer eingeben.")

