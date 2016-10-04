# HardlinkBackup
Einfaches, aber robuste Backuplösung rund um NTFS-Hardlinks.

Die Zeitschrift c't hat in ihrer Ausgabe 09/06, S. 126 eine einfache Backuplösung rund um die Windowsportierung 
von rsync vorgestellt (siehe http://www.heise.de/ct/ftp/06/09/126/). Rsync für Windows hat haber den Nachteil,
dass es cygwin als Umgebung braucht, wodurch die Performance unter Windows ziemlich schlecht ist. Bei größeren
zu sichernden Verzeichnisstrukturen kann die Dauer für einen Backuplauf daher schnell auf 8 Stunden  und mehr 
steigen, was einfach unpraktikabel ist.

Aus diesem Grund habe ich eine eigene simple Implementierung von rsync für Windows geschrieben (https://github.com/linuzer/HardLink),
die nativ auf dem Windows-API läuft und darüber hinaus multithreaded implementiert ist.

Die hier vorgestellte Backup-Lösung ist also im Prinzip die originale c't-Lösung, nur adaptiert auf das
eigene, wesentlich performantere HardLink.

# Installation
Die 3 Dateien in ein Verzeichnis laden, die HardLink.exe dazu packen und in der hardlinkbackup.vbs gemäß den
Kommentaren die Variablen anpassen.

