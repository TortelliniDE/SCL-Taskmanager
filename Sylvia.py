import pandas as pd
import random
from datetime import datetime, timedelta

# Funktionen

# Importieren aller Aufgaben aus einem Excel-File
def import_tasks(file_path):
    tasks_df = pd.read_excel(file_path)
    return tasks_df['Task'].tolist()

# Aktuellen Monat einlesen und Überschriften erstellen
# Titel "Tätigkeitsliste <eingegebener Monat>/<aktuelles Jahr>"
def get_month():
    month = int(input("Bitte geben Sie den Monat (Zahl 1-12) ein: "))
    year = datetime.now().year
    month_name = datetime.strptime(str(month), "%m").strftime("%B")
    return f"Tätigkeitsliste {month_name}/{year}", month, year

# Aktuelle Monatsübersicht inklusive der Feiertage in Bayern erstellen
def create_month(month, year):
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1)
    else:
        end_date = datetime(year, month + 1, 1)
    
    month_days = [start_date + timedelta(days=i) for i in range((end_date - start_date).days)]
    
    # Feiertage Bayern (Beispiel mit Namen)
    feiertage_bayern = {
        datetime(year, 1, 1): "Neujahr",
        datetime(year, 5, 1): "Tag der Arbeit",
        datetime(year, 10, 3): "Tag der Deutschen Einheit",
        datetime(year, 12, 25): "1. Weihnachtsfeiertag",
        datetime(year, 12, 26): "2. Weihnachtsfeiertag"
        # Weitere Feiertage hinzufügen
    }
    
    return month_days, feiertage_bayern

# Soll-Arbeitsstunden und tatsächlich geleistete Stunden eintragen
def working_hours(month_days, feiertage_bayern):
    working_days = [day for day in month_days if day.weekday() < 5 and day not in feiertage_bayern]
    
    soll_hours = 8 * len(working_days)
    ist_hours = 8 * len(working_days)
    return soll_hours, ist_hours, working_days

# Aufgaben zufällig auf alle Arbeitstage verteilen
def create_tasks(tasks, working_days):
    task_distribution = {}
    for day in working_days:
        num_tasks = random.randint(3, 5)
        task_distribution[day] = random.sample(tasks, num_tasks)
    return task_distribution

# Monatsliste auf dem Bildschirm ausdrucken
def print_tasks(header, month_days, task_distribution, feiertage_bayern, soll_hours, ist_hours):
    mitarbeiternummer = "SCL-4711"
    wochenarbeitszeit = 40
    resturlaub = 7.5
    stand_date = datetime.now().strftime("%d.%m.%Y")

    print(header)
    print("\nMitarbeiternummer: " + mitarbeiternummer)
    print("Wochenarbeitszeit: {} h/Wo".format(wochenarbeitszeit))
    print("Resturlaub: {} Tage".format(resturlaub))
    print("Stand: {}".format(stand_date))
    print("\nDatum   Dauer    Wochentag    Tätigkeiten")
    print("="*120)

    for day in month_days:
        day_str = day.strftime("%d.%m.")
        weekday_str = day.strftime("%A")
        if day in feiertage_bayern:
            print(f"{day_str}  0:00     {weekday_str:<10}  {feiertage_bayern[day]}")
        elif day in task_distribution:
            tasks = ", ".join(task_distribution[day])
            print(f"{day_str}  8:00     {weekday_str:<10}  {tasks}")
        else:
            print(f"{day_str}  0:00     {weekday_str:<10}  Wochenende")

    print("\nSollarbeitszeit im Monat {}/{}: {:02d}:00 h".format(month_days[0].month, month_days[0].year, soll_hours))
    print("Geleistete Arbeitszeit im Monat {}/{}: {:02d}:00 h".format(month_days[0].month, month_days[0].year, ist_hours))
    print("\nDie Tätigkeitsliste {}/{} wurde erstellt.".format(month_days[0].month, month_days[0].year))

# Monatsliste als Excel exportieren
def export_tasks(header, month_days, task_distribution, feiertage_bayern, soll_hours, ist_hours, output_path):
    data = []

    for day in month_days:
        day_str = day.strftime("%d.%m.")
        weekday_str = day.strftime("%A")
        if day in feiertage_bayern:
            data.append((day_str, "0:00", weekday_str, feiertage_bayern[day]))
        elif day in task_distribution:
            tasks = ", ".join(task_distribution[day])
            data.append((day_str, "8:00", weekday_str, tasks))
        else:
            data.append((day_str, "0:00", weekday_str, "Wochenende"))

    df = pd.DataFrame(data, columns=['Datum', 'Dauer', 'Wochentag', 'Tätigkeiten'])
    df.loc[len(df)] = ['Gesamt', '', '', '']
    df.loc[len(df)] = ['Sollarbeitszeit', '{:02d}:00'.format(soll_hours), '', '']
    df.loc[len(df)] = ['Ist-Arbeitszeit', '{:02d}:00'.format(ist_hours), '', '']

    sheet_name = "Tätigkeitsliste"
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    df.to_excel(writer, index=False, sheet_name=sheet_name)
    
    worksheet = writer.sheets[sheet_name]
    worksheet.write_string('G1', header)
    worksheet.write_string('G3', "Mitarbeiternummer: SCL-4711")
    worksheet.write_string('G4', "Wochenarbeitszeit: 40 h/Wo")
    worksheet.write_string('G5', "Resturlaub: 7.5 Tage")
    worksheet.write_string('G6', "Stand: {}".format(datetime.now().strftime("%d.%m.%Y")))

    writer.close()

# Hauptprogramm
def main():
    file_path = 'Projekte_SCL.xlsx'
    output_path = 'output.xlsx'

    tasks = import_tasks(file_path)
    header, month, year = get_month()
    month_days, feiertage_bayern = create_month(month, year)
    soll_hours, ist_hours, working_days = working_hours(month_days, feiertage_bayern)
    task_distribution = create_tasks(tasks, working_days)

    print_tasks(header, month_days, task_distribution, feiertage_bayern, soll_hours, ist_hours)
    export_tasks(header, month_days, task_distribution, feiertage_bayern, soll_hours, ist_hours, output_path)

if __name__ == '__main__':
    main()
