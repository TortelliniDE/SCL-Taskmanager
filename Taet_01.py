import os
import calendar
import random
from datetime import date, datetime
import pandas as pd  # Für den Excel-Import
from tkinter import Tk, Listbox, Button, Label, Scrollbar, MULTIPLE, EXTENDED, END, messagebox

from create_task_list import create_task_list


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


# Aufgaben aus Excel laden
def load_tasks_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Task'].dropna().tolist()

def choose_excel_file():
    file_name = 'Tasks_tmp.xlsx'
    if os.path.exists(file_name):
        return file_name
    else:
        print(f"Die Datei {file_name} wurde im aktuellen Verzeichnis nicht gefunden.")
        return None


# GUI-Funktionen
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
        messagebox.showinfo("Erfolg", f"Tasks erfolgreich in '{file_path}' gespeichert. Fenster schließen!")
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
        save_tasks_to_excel(tasks, f"Tasks_tmp.xlsx")
    else:
        messagebox.showwarning("Warnung", "Keine Tasks in der Eingabeliste zum Speichern vorhanden.")



######################################################################################################

# GUI-Setup
root = Tk()
root.title("SCL - Tätigkeiten Liste")
root.geometry("860x750")

# Labels
Label(root, text="Verfügbare Tasks (aus Projektliste_SCL.xlsx)").grid(row=0, column=0, padx=10, pady=5)
Label(root, text="Eingabeliste für aktuellen Monat").grid(row=0, column=2, padx=10, pady=5)

# Listbox für verfügbare Tasks
task_listbox = Listbox(root, selectmode=EXTENDED, width=40, height=25)
task_listbox.grid(row=1, column=0, padx=10, pady=5)

task_scrollbar = Scrollbar(root, orient="vertical", command=task_listbox.yview)
task_scrollbar.grid(row=1, column=1, sticky="ns")
task_listbox.config(yscrollcommand=task_scrollbar.set)

Button(root, text="Ausgewählte hinzufügen →", command=update_selected_tasks).grid(row=1, column=1, padx=10)
#Button(root, text="<-- Ausgewählte entfernen", command=remove_selected_input_tasks).grid(row=1, column=1, pady=10)
Button(root, text="Liste Leeren", command=clear_input_list).grid(row=2, column=2, pady=5)
Button(root, text="Ausgewählte entfernen", command=remove_selected_input_tasks).grid(row=3, column=2, pady=5)

input_task_listbox = Listbox(root, selectmode=MULTIPLE, width=40, height=25)
input_task_listbox.grid(row=1, column=2, padx=10, pady=5)

#Button(root, text="Speichern", command=save_input_list, width=20).grid(row=4, column=2, pady=20)
#Button(root, text="Speichern", command=save_input_list, width=20, bg="green", fg="white").grid(row=4, column=2, pady=20)
Button(root, text="Speichern", command=save_input_list, width=20, bg="green", fg="white",
       activebackground="darkgreen", activeforeground="yellow").grid(row=4, column=2, pady=20)


task_file = "Projekte_SCL.xlsx"
if os.path.exists(task_file):
    tasks = load_tasks_from_excel(task_file)
    for task in tasks:
        task_listbox.insert(END, task)
else:
    messagebox.showerror("Fehler", f"Die Datei '{task_file}' wurde nicht gefunden.")

root.mainloop()





# Tasks Input für aktuellen Monat
selected_file = choose_excel_file()
if selected_file:
    print(f"Die Tätigkeiten-Input Datei für diesen Monat ist: {selected_file}")


# Hauptskript
employees_file = "Mitarbeiter.xlsx"
if not os.path.exists(employees_file):
    print(f"Die Datei {employees_file} wurde nicht gefunden.")
    exit()

employees_df = load_employees_from_excel(employees_file)
first_name, last_name, Mitarbeiternummer, Resturlaub = choose_employee(employees_df)

month = int(input("\nGib den Monat ein für den die Tätigkeitsliste erstellt werden soll (1-12): "))
tasks_file = choose_excel_file()
if not tasks_file:
    print("Keine Datei ausgewählt. Das Programm wird beendet.")
    exit()


tasks = load_tasks_from_excel(tasks_file)
create_task_list(datetime.now().year, month, tasks, first_name, last_name, Mitarbeiternummer, 40, Resturlaub)


# Den Namen der temporären Datei angeben
tasks_tmp = "Tasks_tmp.xlsx"

# Überprüfen, ob die temporäre Datei noch existiert und sie dann löschen
if os.path.exists(tasks_tmp):
    os.remove(tasks_tmp)
    #print(f"Die temporäre Datei {tasks_tmp} wurde erfolgreich gelöscht.")
else:
    print(f"Die Datei {tasks_tmp} existiert nicht.")


#TODO: Ausgabe in excel file
print(f"\nDie Tätigkeitsliste {month}/{datetime.now().year} für {first_name} {last_name} wurde erstellt.\n")

