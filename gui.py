from tkinter import Tk, Listbox, Button, Label, Scrollbar, MULTIPLE, EXTENDED, END, messagebox
import pandas as pd
from logging_util import log_and_print  # Log-Funktion importieren

# GUI-Funktionen
def load_tasks_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'Task' not in df.columns:
            raise ValueError("Die Excel-Datei muss eine Spalte 'Task' enthalten.")
        log_and_print(f"Tasks aus Datei '{file_path}' erfolgreich geladen.")
        return df['Task'].dropna().tolist()
    except Exception as e:
        log_and_print(f"Fehler beim Laden der Excel-Datei '{file_path}': {e}")
        messagebox.showerror("Fehler", f"Fehler beim Laden der Excel-Datei: {e}")
        return []

def save_tasks_to_excel(tasks, file_path, root):
    try:
        df = pd.DataFrame({'Task': tasks})
        df.to_excel(file_path, index=False)
        log_and_print(f"Tasks erfolgreich in '{file_path}' gespeichert.")
        messagebox.showinfo("Erfolg", f"Tasks erfolgreich in '{file_path}' gespeichert. Fenster schließen!", parent=root)
        root.quit()  # Schließt das Hauptfenster
        root.destroy()  # Zerstört das Hauptfenster
    except Exception as e:
        log_and_print(f"Fehler beim Speichern der Excel-Datei '{file_path}': {e}")
        messagebox.showerror("Fehler", f"Fehler beim Speichern der Excel-Datei: {e}")

def update_selected_tasks(task_listbox, input_task_listbox):
    selected = list(task_listbox.curselection())
    for index in selected:
        task = task_listbox.get(index)
        if task not in input_task_listbox.get(0, END):
            input_task_listbox.insert(END, task)
    log_and_print("Ausgewählte Tasks zur Eingabeliste hinzugefügt.")

def add_all_tasks(task_listbox, input_task_listbox):
    # Fügt alle Tasks aus der verfügbaren Liste zur Eingabeliste hinzu
    for index in range(task_listbox.size()):
        task = task_listbox.get(index)
        if task not in input_task_listbox.get(0, END):
            input_task_listbox.insert(END, task)
    log_and_print("Alle Tasks zur Eingabeliste hinzugefügt.")

def clear_input_list(input_task_listbox):
    input_task_listbox.delete(0, END)
    log_and_print("Eingabeliste wurde geleert.")

def remove_selected_input_tasks(input_task_listbox):
    selected = list(input_task_listbox.curselection())
    for index in reversed(selected):
        input_task_listbox.delete(index)
    log_and_print("Ausgewählte Tasks aus der Eingabeliste entfernt.")

def save_input_list(input_task_listbox, root):
    tasks = list(input_task_listbox.get(0, END))
    if tasks:
        log_and_print(f"{len(tasks)} Tasks werden gespeichert.")
        save_tasks_to_excel(tasks, f"Tasks_tmp.xlsx", root)
    else:
        log_and_print("Speichern abgebrochen: Keine Tasks in der Eingabeliste.")
        messagebox.showwarning("Warnung", "Keine Tasks in der Eingabeliste zum Speichern vorhanden.")

def setup_gui(task_file):
    root = Tk()
    root.title("SCL - Tätigkeiten Liste")
    root.geometry("860x750")

    log_and_print("GUI gestartet.")

    # Labels
    Label(root, text="Verfügbare Tasks (aus Projektliste_SCL.xlsx)").grid(row=0, column=0, padx=10, pady=5)
    Label(root, text="Eingabeliste für aktuellen Monat").grid(row=0, column=2, padx=10, pady=5)

    # Listbox für verfügbare Tasks
    task_listbox = Listbox(root, selectmode=EXTENDED, width=40, height=25)
    task_listbox.grid(row=1, column=0, padx=10, pady=5)

    task_scrollbar = Scrollbar(root, orient="vertical", command=task_listbox.yview)
    task_scrollbar.grid(row=1, column=1, sticky="ns")
    task_listbox.config(yscrollcommand=task_scrollbar.set)

    # Buttons
    Button(root, text="Alle auswählen", command=lambda: add_all_tasks(task_listbox, input_task_listbox)).grid(row=1, column=1, pady=10)
    Button(root, text="Ausgewählte hinzufügen →", command=lambda: update_selected_tasks(task_listbox, input_task_listbox)).grid(row=2, column=1, pady=10)
    Button(root, text="Liste Leeren", command=lambda: clear_input_list(input_task_listbox)).grid(row=3, column=2, pady=5)
    Button(root, text="Ausgewählte entfernen", command=lambda: remove_selected_input_tasks(input_task_listbox)).grid(row=4, column=2, pady=5)

    input_task_listbox = Listbox(root, selectmode=MULTIPLE, width=40, height=25)
    input_task_listbox.grid(row=1, column=2, padx=10, pady=5)

    Button(root, text="Speichern", command=lambda: save_input_list(input_task_listbox, root), width=20, bg="green", fg="white",
           activebackground="darkgreen", activeforeground="yellow").grid(row=5, column=2, pady=20)

    # Tasks aus der Datei laden und in die Listbox einfügen
    if task_file:
        tasks = load_tasks_from_excel(task_file)
        for task in tasks:
            task_listbox.insert(END, task)
    else:
        log_and_print(f"Die Datei '{task_file}' wurde nicht gefunden.")
        messagebox.showerror("Fehler", f"Die Datei '{task_file}' wurde nicht gefunden.")

    root.mainloop()
    log_and_print("GUI beendet.")

