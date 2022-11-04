# Anwendung
$k=[kzven]::new()
$k.kzv|select kzvnummer, Name, Homepage | sort kzvnummer
$k.kzv|select kzvnummer, Name, Homepage | sort Name
$k.kzv|select Kurzname, PreisCSVLink | sort Kurzname
$k.kzv|select Kurzname, kzvnummer, @{N='CSVName';E={$_.PreisCSVLink.Segments[-1]}} | sort Kurzname

# Preise können generell von KZV, Innung, oder Softwarehersteller kommen
# in der Regel in der Form CSV-, PDF- oder Excel-Datei

# für XLSX-Dateien kann man das Modul ImportExcel verwenden, ist Core kompatibel
# Install-Module ImportExcel -Scope CurrentUser

# eine bestimmte KZV auswählen
$kzv = $k.kzv|where name -eq "Thüringen"

