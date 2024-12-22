Tätigkeitsnachweis eines Monats generieren
###########################################

Mit Doppelklick auf die Datei "Tätigkeitsnachweis.bat" wird eine Tätigkeitsliste erzeugt und in einem excel file "output.xmls" im gleichen Verzeichnis gespeichert.
Dabei wird man zunerst zur Eingabe des Monats aufgefordert (eingeben als 1 für Januar, 2 für Februar, ... 12 für Dezember)
Dann kommt die Frage, ob man in dem Monat Urlaub hatte (j/n)
Falls ja kann man den Tag (z.B. 23) oder mehrere Tage eingeben (z.B. 24, 31) oder den Ulaubsbereich von-bis (z.B. 9-13) eingeben.
Falls nein wird das Programm ohne Urlaubseingabe fortgesetzt.
Das Programm verteilt mindestens 6 und max. 7 Aufgaben aus der SCL-Projekte Excel Datei und verteilt diese per Zufallsgenerator auf die Arbeitstage.
Für jeden Arbeitstag werden 8:00 h eingetragen und die Summe der Soll und Ist Arbeitstage am Ende ausgegeben.
Wochenenden, Urlaube und Feiertage werden mit 0:00 h eingetragen.
