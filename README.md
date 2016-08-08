## comtrade-adjust

### German
Dieses vbscript-Makro beseitigt folgende Fehler in von PSS Sincal bzw. PSS NETOMAC erzeugten COMTRADE-Dateien:

*  Die von Sincal erzeugten Signalnamen für die Plottdefinition sind zu lang für NETOMAC, aus irgendwelchen Gründen wirft NETOMAC in der .cfg dann die Einheit für den Strom weg. Diese wird durch "kA" ergänzt
*  Das Datumsformat im Zeitstempel ist falsch, Monat und Tag sind vertauscht. Einige Programme (z.B. Kocos Artes) können deswegen die Datei nicht öffnen.

#### Einbindung in Sincal
Das Makro darf NICHT auf einem Netzlaufwerk liegen, ansonsten kann Sincal es nicht ausführen. Über Extras -> Optionen -> Makros kann der Pfad festgelegt werden, in denen Sincal nach vbscript-Makros suchen soll. Anschließend kann das Makro über Extras -> Makro -> Makros… -> comtrade-adjust.vbs ausgeführt werden.

### English
todo
