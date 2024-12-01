import os
from datetime import date, datetime
import pandas as pd  # Für den Excel-Import


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
    

# Den Namen der temporären Datei angeben
tasks_tmp = "Tasks_tmp.xlsx"