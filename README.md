# JsonExcel
Dieses AddIn dient der Darstellung *eines* JSON-Objektes in Excel.
Dabei werden die Name/Wert Paare in einer gemeinsamen Spalte angezeigt, während die Json-Struktur in den Spalten davor angezeigt werden.

![JsonExcel Ribbon](https://raw.githubusercontent.com/diogenes25/JsonExcel_AddIn/master/Doc/images/JsonExcelRibbon.PNG)


1. [Arbeitsweise](##arbeitsweise)
2. [Erstellen von JSON-Objekten](##Erstellen_von_JSON-Objekten)

## Arbeitsweise
Im Gegensatz zu den bisherigen Excel-AddIns konzentriert sich dieses AddIn mehr auf die Json-Struktur, statt auf wiederkehrende Daten.
Die Struktur des Json-Objektes wird beibehalten indem jedes Name/Wert Paar in einer Zeile inkl. des gesamten Objekt-Pfad enthält angezeigt wird.

### Beispiel JSON-Object
```json
{
	"family" : {
		"father" : {
			"name" : "Homer",
			"profession" : "Operator",
			"Interests": {
				"beer" : "Duff Beer",
				"food" : "Donut"
			}
		},
		"mother" : {
			"name" : "Marge"
		}
	}
}
```
### Darstellung in Excel

|   |   A    |   B    |      C    |    D       | E |   F     |
|---|--------|--------|-----------|------------|:-:|---------|
| 1 |[family]|[father]|           |[name]      | : |Homer    |
| 2 |[family]|[father]|           |[profession]| : |Operator |
| 3 |[family]|[father]|[interests]|[beer]      | : |Duff Beer|
| 3 |[family]|[father]|[interests]|[food]      | : |Donut    |
| 4 |[family]|[mother]|           |[name]      | : |Marge    |

![Excel vs. Json view](https://raw.githubusercontent.com/diogenes25/JsonExcel_AddIn/master/Doc/images/ExcelAndVim.PNG)

## Erstellen von JSON-Objekten
Jede Zeile muss ein Name:Wert - Paar enthalten.
Der Doppelpunkt (':') muss direkt vor der Zelle mit dem Wert eingetragen werden.
Der erste Wert vor dem ':' ist der Name des Wertes. Leere Zellen vor dem Doppelpunkt werden ignoriert.
Sämtliche Zeilen vor der "Name"-Zelle definieren die Struktur des Json-Objektes.
Eckige Klammer in den "Struktur"-Zellen dienen nur der Darstellung und werden beim Exportieren ignoriert.


![Json erstellen](https://raw.githubusercontent.com/diogenes25/JsonExcel_AddIn/master/Doc/JsonExcel/Video/MP4/JsonExcelExport2.gif)

## Hintergrund
Excel ist zur Bearbeitung von JSON-Dateien ein nur bedingt geeignetes Programm.
Während JSON strukturierte Daten darstellen kann, sind mit Excel „nur“ zweidimensionale Matrizen möglich.
Entsprechend werden bei den meisten JSON-Dateien die mit Excel bearbeitet werden auch nur zwei Dimensionen betrachtet und die Bearbeitung der JSON-Daten konzentriert sich ausschliesslich auf die reinen Werte und nicht auf die JSON-Objekt-Struktur.

### Beispiel der Funktionsweise anderer AddIns:

|   |   A    |   B    |      C    |
|---|:------:|:------:|:---------:|
| 1 |Value of Cell A1|Value of Cell B1|Value of Cell C1|
| 2 |Value of Cell A2|Value of Cell B2|Value of Cell C2|
| 3 |Value of Cell A3|Value of Cell B3|Value of Cell C3|

### Ausgabe als JSON.

```json
[
	{
		"col_One" : "Value of Cell A1",
		"col_Two" : "Value of Cell B1",
		"col_Three" : "Value of Cell C1",
	},
	{
		"col_One" : "Value of Cell A2",
		"col_Two" : "Value of Cell B2",
		"col_Three" : "Value of Cell C2",
	},
	{
		"col_One" : "Value of Cell A3",
		"col_Two" : "Value of Cell B3",
		"col_Three" : "Value of Cell C3",
	}
]
``` 
