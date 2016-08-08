Dim SincalApp
Dim SincalDoc
Dim path

' Sincal-Projekt initialisieren
Set SincalApp = WScript.CreateObject("SIASincal.Application")
Set SincalDoc = SincalApp.ActiveDocument

' Pfad auslesen und vervollständigen
path = left(SincalDoc.DataSource, len(SincalDoc.DataSource) - 12) & "NETO\COMTRADE\NETWORK.CFG"

' NETWORK.CFG einlesen
Dim Dateisystem, Textdatei, Text
Set Dateisystem = CreateObject("Scripting.FileSystemObject")
Set Textdatei = Dateisystem.OpenTextFile(path)
Text = Textdatei.ReadAll
Textdatei.Close

'Datei löschen
Dateisystem.DeleteFile path

' erste Zeile überspringen
position = instr(Text, vbCrLf)


' Anzahl Signale lesen
position = instr(position+1, Text, ",")
anz_sig = Cint(mid(Text, position-3, 3))

' 
akt_sig = 0
do
	' nächste Zeile
	position = instr(position + 1, Text, vbCrLf)

	i = 0
	
	' gehe zum fünften Komma
	do
		position = instr(position + 1, Text, ",")
		i = i+1
	loop while i < 4

	' Wenn Zeichen dahinter ein Komma ist, setze "kA" ein
	if mid(Text, position+1, 1) = "," then
			
			' String wieder zusammenbauen
			Text = left(Text, position) & "kA" & right(Text, len(Text) -position)
		end if

		akt_sig = akt_sig + 1
		
' wiederhole für jedes Signal
loop while akt_sig < anz_sig

i = 0
do while i < 2
	' Suche Position von '/' als Zeichen für Zeitstempel
	position = instr(position + 1, Text, "/")
	tag = mid(Text, position + 1, 2)
	monat = mid(Text, position - 2, 2)
	datum = mid(Text, position-2, 10)

	Text = left(Text, position - 3) & tag & "/" & monat & right(Text, len(Text) - position - 2)
	' nächste Zeile
	position = instr(position + 1, Text, vbCrLf)
	i = i+1
loop

' neue Datei schreiben
Set Textdatei = Dateisystem.CreateTextFile(path)
Textdatei.Write Text
Textdatei.Close
