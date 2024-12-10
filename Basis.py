
# Funktionen

# Importieren aller Aufgaben aus einem excel file
def import_tasks:
    pass

# Aktuellen Monat einlesen und Überschriften erstellen
# Titel "Tätigkeitsliste <eingegebener Monat>/<aktuelles Jahr> "
def get_month

# Aktuelle Momatsübersicht inklusive der Feiertage in Bayern erstellen
def create_month

# Soll-Arbeitsstunden nur für die Arbeitstage in der erzeugten Monatsliste mit 8:00 pro Arbeitstag eingeben, alle anderen mit 0:00 füllen
# Sollarbeitszeit (Summe der Stunden aller Arbeitstage) für den aktuellen Monat berechnen und am Ende ausgeben in h
# Tatsächlich geleistete Stunden für alle Arbeitstage für jeden Arbeitstag auch mit 8:00 eintragen (außer Feiertage und Wochenende)
# Tatsächlich geleistete Stunden der Arbeitstage summieren und am Ende der Sollarbeitszeit gegenüberstellen in h
def working_hours

# Aufgaben aus der importierten Liste zufällig auf alle Arbeitstage der vorher erstellten Monatsübersicht verteilen (außer Wochenende und Feiertage)
# es sollen pro Tag min 3 und max 5 tasks verteilt werden
def create_tasks

# Die Überschrift und die Monatsliste incl. aller Stunden und Feiertage und der tasks ausgeben
def print_tasks

# Die Überschrift und die Monatsliste incl. aller Stunden und Feiertage und der tasks als excel file exportieren
def export_tasks:

# Hauptprogramm
def main():


if __name__ == '__main__':
    main()