'......................................................................................
'... rsyncBackup.vbs 1.04 .................. Autor: Karsten Violka kav@ctmagazin.de ...
'... c't 9/06 .........................................................................
'......................................................................................
'
'--------------------------------------------------------------------------------------
' Bekannte Probleme:
'   -- rsync kopiert keine geöffneten Dateien
'   -- rsync kopiert nur Pfade bis zu einer Länge von 260 Zeichen.
'   -- rsync kopiert keine NTFS-Spezialitäten (Junctions, Streams, Sparse Files)

' Skript mit niedriger Priorität starten: 
' 	start /min /belownormal cscript.exe rsyncBackup.vbs
'--------------------------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------------------------
'----- Konfiguration ------------------------------------------------------------------
'--------------------------------------------------------------------------------------

' Pfad zur Datei mit der Liste der Quellverzeichnisse
const BACKUPFOLDER = "D:\Backup\bin\BackupFolder.txt"

' Kommaseparierte Liste mit Dateinamen. *-Wildcard ist möglich
const EXLUDE_FILES = "Cache,parent.lock,Temp*"

' Das Zielverzeichnis sollte sich auf einem mit NTFS formatierten Laufwerk befinden
'const DESTINATION="e:\rsyncbackup"
const DESTINATION="D:\Backup\Data"

' Pfad für die Log-Dateien. KEIN abschliessender Backslash!!!
const LOGPATH="D:\Backup\Log"

' Pfad für das rsync Verzeichnis. KEIN abschliessender Backslash!!!
'const HARDLINKPATH="$USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Hardlink\Hardlink.appref-ms"
const HARDLINKPATH="D:\Backup\bin\Hardlink.exe"

' Anzahl der aufbewahrten Backups:
const STAGE0_HOURLY = 1
const STAGE1_DAILY = 7
const STAGE2_WEEKLY = 4

const STAGE1_DAILY_FOLDER =  "\1_täglich"
const STAGE2_WEEKLY_FOLDER=  "\2_wöchentlich"
const STAGE3_MONTHLY_FOLDER= "\3_monatlich"

' Konstanten für ADO
const adVarChar = 200
const adDate = 7
' Feldnamen fürs RecordSet
Dim rsFieldNames
rsFieldNames = Array("name", "date")

'---- Global verwendete Variablen
Dim fso, wsh
Dim sourceFolders
Dim excludeFiles
Dim logFile, logFileData, logFileError
dim strHardlinkPath
Dim strSourceFolder, recentBackupFolder, strDateFolder, strDestinationFolder
Dim strCmd, cmdResult

'--------------------------------------------------------------------------------------
'----- Hauptroutine -------------------------------------------------------------
'--------------------------------------------------------------------------------------

'logFile = wsh.ExpandEnvironmentStrings("%userprofile%") & "\rsyncBackup.log"
logFile = LOGPATH & "\" & getDateFolderName() & ".log"
logFileData = LOGPATH & "\" & getDateFolderName() & "_Data.log"
logFileError = LOGPATH & "\" & getDateFolderName() & "_Error.log"

if instr(HARDLINKPATH,"$USERPROFILE") then
	strHardlinkPath = replace(HARDLINKPATH, "$USERPROFILE", wsh.ExpandEnvironmentStrings("%userprofile%"))
else
	strHardlinkPath = HARDLINKPATH
end if

Set recentBackupFolder = Nothing
set fso = CreateObject("Scripting.FileSystemObject")
set wsh = CreateObject("WScript.Shell")


logAppend(vbCRLf & "-------- Start: " & Now & " --------------------------------------------")

' Quellverzeichnisse vorbereiten
sourceFolders = GetBackupFolder()
checkFolders()
excludeFiles = split(EXLUDE_FILES, ",")

' Zielverzeichnisse vorbereiten
strDateFolder = getDateFolderName()
strDestinationFolder = DESTINATION & "\~" & strDateFolder ' Zielordner zunächst Tilde voranstellen
Set recentBackupFolder = getRecentFolder(DESTINATION)

' Befehlszeile zusammenbauen
strCmd=getCommandline()
logAppend("--- rsync-Befehlszeile:")
logAppend(strCmd)

' Backup starten
cmdResult=wsh.Run(strCmd, 0, true)
if cmdresult = 0 then

	' Zielordner umbenennen und Tilde entfernen
	fso.MoveFolder strDestinationFolder, DESTINATION & "\" & strDateFolder

	'-- Backups rotieren und alte Backups löschen
	rotate getFolderObject(DESTINATION), _
			getFolderObject(DESTINATION & STAGE1_DAILY_FOLDER), STAGE0_HOURLY, "d"
	rotate getFolderObject(DESTINATION & STAGE1_DAILY_FOLDER), _
			getFolderObject(DESTINATION & STAGE2_WEEKLY_FOLDER), STAGE1_DAILY, "ww"
	rotate getFolderObject(DESTINATION & STAGE2_WEEKLY_FOLDER), _
			getFolderObject(DESTINATION & STAGE3_MONTHLY_FOLDER), STAGE2_WEEKLY, "m"
else
	logAppend("--- Ausgabe von rsync:" & vbCrLf & cmdResult)
end if
logAppend("-------- Fertig: " & Now & " --------------------------------------------")


'---------------------------------------------------------------------------------------
'--- Funktionen ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

'--- checkFolders() -------------------------------------------------------------------
' Prüft ob die eingetragenen Pfade plausibel sind.
Function checkFolders()
	
	Dim aSourceFolder
	
	For Each aSourceFolder in sourceFolders
		If Not fso.FolderExists(aSourceFolder) Then
			criticalErrorHandler "checkFolders()", "Quellordner '" & aSourceFolder & "' existiert nicht.", 0, ""
		End If
	Next
	
	If Not fso.DriveExists(fso.getDriveName(DESTINATION)) Then
		criticalErrorHandler "checkFolders()", "Ziellaufwerk " & fso.getDriveName(DESTINATION) & " nicht gefunden", 0, ""
	End If
	
	If Not fso.getDrive(fso.getDriveName(DESTINATION)).FileSystem = "NTFS" Then
		logAppend("--- Warnung: Zielpfad " & DESTINATION & " liegt nicht auf einem NTFS-Laufwerk!")
		logAppend("--- Warnung: rsync erstellt dort keine Hard-Links, sondern vollständige Kopien")
	End If
	
End Function

'--- getRsyncCmd() ----------------------------------------------------------------------
' Baut das hardlink-Kommando zusammen.
Function getCommandline()
	dim cmd, aSourceFolder, aExcludeFile
	
	'cmd = wsh.ExpandEnvironmentStrings("%comspec%") & " /c " & chr(34) & strHardlinkPath & chr(34) & " "
	cmd = chr(34) & strHardlinkPath & chr(34) 
	
	If Not recentBackupFolder Is Nothing Then
		cmd = cmd & " --link-dest """ _
			& recentBackupFolder.Path & """"
	End If
	
'	For Each aExcludeFile in excludeFiles
'		cmd = cmd & " --exclude """ & aExcludeFile & """"
'	Next
	
	cmd = cmd & " --source-folder "
	For Each aSourceFolder in sourceFolders
		cmd = cmd & """" & aSourceFolder & """ "
	Next
	
	cmd = cmd & " --destination """ & strDestinationFolder &  """"
	cmd = cmd & " --logfile """ & logFileData & """"
	cmd = cmd & " --errorfile """ & logFileError & """"
	
	getCommandline = cmd
End Function

'--- getDateFolderName()------------------------------------------------------------
' Generiert einen Ordnernamen mit dem aktuellen Datum und der Uhrzeit.
Function getDateFolderName()
	Dim jetzt
	jetzt = Now()
	getDateFolderName = Year(jetzt) & "-" & addLeadingZero(Month(jetzt))_
		& "-" & addLeadingZero(Day(jetzt))_
		& "_"	& addLeadingZero(Hour(jetzt))_
		& "~" & addLeadingZero(Minute(jetzt))
End Function

'--- addLeadingZero(number) -------------------------------------------------------------
' Fügt bei Zahlen < 10 führende Null ein.
Function addLeadingZero(number)
	If number < 10 Then
		number = "0" & number
	End If 
	addLeadingZero = number
End Function

'--- getFolderObject(path) -------------------------------------------------------------
' Liefert zum übergebenen Pfad-String ein WSH-Objekt vom Typ Folder
' Wenn das Verzeichnis noch nicht existiert, wird es angelegt.
Function getFolderObject(path)
	If (fso.FolderExists(path)) Then
		Set getFolderObject = fso.GetFolder(path)
	Else
		logAppend("--- Erstelle Ordner: " & path)
		On Error Resume Next
		Set getFolderObject = fso.CreateFolder(path)
		
		If Err.Number <> 0 Then
			On Error Goto 0
			criticalErrorHandler "getFolderObject()", "Konnte Zielordner nicht erstellen", Err.Number, Err.Description
		End If
		
		On Error Goto 0
	End If
End Function

'--- toCygwinPath(String) -----------------------------------------------------------------
' Wandelt einen Windows-Pfad in das Format, das Cygwin erwartet
Function toCygwinPath(path)
	Dim driveLetter, newPath
	driveLetter = Left(fso.GetDriveName(path), 1)
	newPath = Replace(path, "\", "/")
	newPath = Mid(newPath, 4)
	toCygwinPath = "/cygdrive/" & driveLetter & "/" & newPath
End Function

'--- toCrLf(String) -----------------------------------------------------------------------
' Ersetzt den von rsync ausgegebenen Unix-Zeilenumbruch (LF)
' durch das Windows-übliche Format (CRLF)
Function toCrLf(strText)
	toCrLf = Replace(strText, vbLf, vbCrLf)
End Function

'--- logAppend(String) --------------------------------------------------------------------
' hängt den übergebenen Text an die Log-Datei an
Function logAppend(string)
	const forAppend = 8
	dim f, errnum
	
	On Error Resume Next	
	Set f = fso.OpenTextFile(logFile, forAppend, true)
	errnum = Err.Number
	
	On Error Goto 0
	If errnum = 0 Then
		f.WriteLine(string)
		f.Close()
	Else
		Err.Raise 1, "logAppend", "Konnte Logdatei nicht öffnen"
	End If
End Function

'--- getRecentFolder(String) ---------------------------------------------------------------
' Sortiert die im übergebenen Pfad enthaltenen Ordner nach Datum und liefert das jüngste
' Ordner-Objekt zurück
' Parameter: Pfad als String
Function getRecentFolder(path)
	Dim destinationFolder, rs
	Set destinationFolder = getFolderObject(path)
	Set rs=newFolderRecordSet(destinationFolder)
	
	If Not (rs.Eof) Then
		rs.sort = "date DESC"		' absteigend nach Erstellungszeitpunkt sortieren 
		rs.MoveFirst
		Set getRecentFolder= fso.GetFolder(rs.fields("name"))
	Else
		Set getRecentFolder = Nothing
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- newFolderRecordSet(Folder-Objekt) -----------------------------------------------------
' Füllt die Unterordner der übergebenen Folder-Objekts in ein neues RecordSet-Objekt,
' das zum Sortieren verwendet wird.
Function newFolderRecordSet(folder)
	Dim aFolder
	Set newFolderRecordSet = CreateObject("ADODB.Recordset")
	newFolderRecordSet.Fields.Append "name", adVarChar, 255
	newFolderRecordSet.Fields.Append "date", adDate
    newFolderRecordSet.Open
	
	For Each aFolder in folder.SubFolders
		If Left(aFolder.Name, 2) = "20" Then	' nur die Datumsordner in die Liste aufnehmen
			newFolderRecordSet.addnew rsFieldNames, Array(aFolder.Path, aFolder.DateCreated)
		End if
	Next	
End Function

'--- rotate(fromFolder, toFolder, numberToKeep, diffInterval) ------------------------------
' Verschiebt oder löscht die Backup-Ordner. Fürjedes Zeitintervall (Tag, Woche, Monat) wird
' jeweils das zuletzt erstellte Backup archiviert.
Function rotate(fromFolder, toFolder, numberToKeep, diffInterval)
	
	Dim rs, aFolder, lastFolder, i, recentBackup, errNr, errDesc
	
	Set rs=newFolderRecordSet(fromFolder)
	
	If Not (rs.Eof) Then
		rs.Sort = "date DESC"
		rs.MoveFirst
		i = 0
		Do until rs.Eof
			If i >= numberToKeep Then
				'MsgBox("übrig:" & rs.fields("name"))
				'Das jüngste Backup dieses Datums aus dem toFolder holen. Wenn neuer, ersetzen.
				Set recentBackup = getRecentBackupForDate(toFolder, rs.fields("date"), diffInterval)
				On Error Resume Next
				If Not recentBackup Is Nothing Then
					' Wenn das gewählte Backup vom selben Zeitintervall (Tag) ist und
					' später erstellt wurde, soll es das Backup im Zielordner ersetzen.
					If DateDiff("s", recentBackup.DateCreated, rs.fields("date")) > 0 Then
						'MsgBox("selber Tag & neuer: bewegen")
						logAppend("--- bewege " & rs.fields("name") & " nach " & toFolder.Path)
						fso.MoveFolder fso.GetFolder(rs.fields("name")), toFolder.Path & "\"
						If Err.Number <> 0 Then 
							ErrNr=Err.Number
							ErrDesc=Err.Description
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner '" & rs.fields("name") & "' nicht nach '" & toFolder.Path & "\' bewegen", ErrNr, ErrDesc
						End If
						'MsgBox("Vorgänger löschen.")
						logAppend("--- Vorgänger löschen " & recentBackup)
						fso.DeleteFolder recentBackup, true
						
						If Err.Number <> 0 Then
							ErrNr=Err.Number
							ErrDesc=Err.Description
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner '" & recentBackup & "' nicht löschen", ErrNr, ErrDesc
						End If					
					Else
						logAppend("--- lösche " & rs.fields("name"))
						'MsgBox("selber Tag & älter: weg damit.")
						fso.DeleteFolder fso.GetFolder(rs.fields("name")), true
					
						If Err.Number <> 0 Then 
							ErrNr=Err.Number
							ErrDesc=Err.Description
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner '" & rs.fields("name") & "' nicht löschen", ErrNr, ErrDesc
						End If
					End If
				Else
					' Vom diesem Tag existiert noch kein Backup
					'MsgBox("noch nicht da, bewegen!")
					logAppend("--- bewege " & rs.fields("name") & " nach " & toFolder.Path)
					fso.MoveFolder fso.GetFolder(rs.fields("name")), toFolder.Path & "\"
					If Err.Number <> 0 Then 
						ErrNr=Err.Number
						ErrDesc=Err.Description
						On Error Goto 0
						criticalErrorHandler "rotate()", "Konnte Ordner '" & rs.fields("name") & "' nicht nach '" & toFolder.Path & "\' bewegen", ErrNr, ErrDesc
					End If	
				End If
				On Error Goto 0
			End If
			i = i + 1
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- getRecentBackupForDate(folderObj, aDate, diffInterval) -----------------------------
' Sortiert die Unterverzeichnisse mit Hilfe des ADO RecordSet und liefert
' das das letzte Backup des angegeben Tages/der Woche/des Monats --> diffInterval
Function getRecentBackupForDate(folderObj, aDate, diffInterval)
	Dim rs, exitLoop
	Set getRecentBackupForDate = Nothing
	Set rs=newFolderRecordSet(folderObj)
	If Not (rs.Eof) Then
		rs.Sort = "date DESC"
		rs.MoveFirst
		exitLoop=false 
		Do until rs.Eof Or exitLoop
			If DateDiff(diffInterval, rs.fields("date"), aDate) = 0 Then
				set getRecentBackupForDate = fso.GetFolder(rs.fields("name"))
				exitLoop = true
			End If
		   rs.MoveNext
		Loop	  
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- criticalErrorHandler(source, description, errNumber, errDescription) ---------------
' Kritischen Fehler loggen und Programm abbrechen. Vor dem Aufruf muss die
' Fehlerbehandlung mit "On Error Goto 0" wieder eingeschaltet werden, damit das Skript
' mit dem neu erzeugten Fehler abbricht.
Function criticalErrorHandler(source, description, errNumber, errDescription)
	logAppend("--- Fehler: Funktion " & source & ", " & description)
	logAppend("            Err.Number: " & errNumber & " Err.Description:" & errDescription)
	logAppend("-------- Stop: " & Now & " --------------------------------------------")
	Err.Raise 1, source, description
End Function

function GetBackupFolder()
   
   Dim fso2
   Dim file
   Dim s
   Dim i
   Dim strBuffer
   
   Set fso2 = createobject("Scripting.FileSystemObject")
   Set file = fso2.GetFile(BACKUPFOLDER)
   Set s = file.OpenAsTextStream(1) '1 = ForReading
   
   Do Until s.AtEndOfStream
	  i = s.ReadLine
	  if not left(trim(i),1) = "'" then
          strBuffer = strBuffer & i & vbCrLf
	  end if
   Loop
   
   strbuffer = left(strbuffer, len(strbuffer)-len(vbcrlf))

   getbackupfolder = split(strbuffer, vbcrlf)
   
end function