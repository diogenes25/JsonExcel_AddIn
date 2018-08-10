# JsonExcel


## Bearbeiten von JSON-Dateien inkl. dire Json-Struktur. 

![Json erstellen](Doc\JsonExcel\Video\WebM\JsonExcel.webm)

|   |   A    |   B    |      C    |    D       | E |   F     |
|---|--------|--------|-----------|------------|:-:|---------|
| 1 |[family]|[father]|           |[name]      | : |Homer    |
| 2 |[family]|[father]|           |[profession]| : |Operator |
| 3 |[family]|[father]|[interests]|[beer]      | : |Duff Beer|
| 3 |[family]|[father]|[interests]|[food]      | : |Donut    |
| 4 |[family]|[mother]|           |[name]      | : |Marge    |

```json
{
	"family" : {
		"father" : {
			"name" : "Homer",
			"profession" : "Operator in nuclear power station",
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
