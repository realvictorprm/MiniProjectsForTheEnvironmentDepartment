Hier findet sich das Programm für die Auswertung der Zugriffsstatistiken.


----


## Features:
- Automatisches Aufsummieren der Unterthemen zum jeweiligen Übergeordnetem Thema
- Automatische Erstellung eines Pie-Charts welches die ausgewerteten Daten als Grundlage benutzt
- Flexibel veränderbares Thema-Mapping

## Requirements:
#### Windows:
- Mindestens das .NET Framework V4.5 (Windows 7 minimum)
#### Linux oder Mac:
- Mono Version die mindestens .NET V4.5 unterstützt

## Bedienungs-Anleitung:
Um eine automatische Auswertung auszuführen einfach die Excel-Datei (von der Datenbank, roh und unbearbeitet) auf die *.exe Datei ziehen.
Es erscheint ein Kommandofenster welches lediglich Debug-Informationen ausgibt, falls etwas schiefgelaufen ist.
Sobald "done" erscheint, wurde eine Excel-Datei namens "Generated.xlsl" in demselben Ordner wie die *.exe Datei erstellt. 
 
## Problembehandlung:
#### Die Überschrift des generierten Diagrammes lautet [MISSING], wie ist dies zu vermeiden?
Bitte die Datei nicht umbennen, es wird erwartet, dass das Datum im Dateinamen enthalten ist. Beispiel: `hitliste_selektoren_udo_01_2017.xlsx` . Dabei sucht das Programm immer nach diesem Abschnitt `01_2017`
#### Es kommen komische Fehler.
Am besten den Fehler genau lesen. Dieser ist immer in Englisch und beschreibt in etwa was schief gelaufen ist. Eventuell findet sich hier ein Eintrag für den Fehler.
#### Es wird ein Fehler angezeigt der in etwa lautet: `System.Exception: Invalid JSON starting at character ...`
Es befindet sich ein Syntaxfehler im Mappingfile. Bitte überprüfen ob eventuell Kommas, Klammern oder Anführungsstriche doppelt sind oder fehlen.
#### Es wird ein Fehler angezeigt der in etwa lautet: `System.IO.FileNotFoundException: Die Datei oder Assembly ...` 
Das Programm wurde nicht korrekt entpackt bzw. es fehlend Bibliotheksdateien. Bitte die Ordnerstruktur unverändert in einen leeren Ordner kopieren. Falls dennoch Probleme auftreten bitte den Systemadministrator bzw. mich kontaktieren.
#### Es wird ein Fehler angezeigt der in etwa lautet: `System.Exception: invalid`
Die Mapping Datei ist zwar strukturel korrekt jedoch nicht inhaltlich. Es sind nur Zeichenketten erlaubt, keine Zahlen oder Werte.

## Zusätzliche Informationen:

Das Programm kann auch über die Commandozeile gestartet werden. Es wird automatisch darauf hingewiesen, dass Argumente erforderlich sind für den Aufruf. Um zu erfahren welche Verfügbar sind einfach `/h` als Argument übergeben.

Die Mapping-Datei kann und soll gegebenenfalls angepasst werden. Diese befindet sich im demselben Ordner wie die *.exe Datei und heißt `mappingfile.json`. Die Struktur ist folgende:
 
## Mappingfile Schema:
```json
{
 "Name des übergeordneten Themas. Inhalt immer in Anführungszeichen! Mehrere Themen mit Komma trennen.": [
    "Name des untergeordneten Themas.  Inhalt immer in Anführungszeichen! Mehrere Themen mit Komma trennen. "
  ]
}
```
Die untergeordneten Themen sind in einem Array bzw. Liste gespeichert und müssen deswegen immer von eckigen Klammern umgeben sein.
