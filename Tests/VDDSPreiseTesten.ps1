# Anwendung
$k=[kzven]::new()
$k.kzv|select kzvnummer, Name, Homepage | sort-object kzvnummer
$k.kzv|select kzvnummer, Name, Homepage | sort-object Name
$k.kzv|select Kurzname, PreisCSVLink | sort-object Kurzname
$k.kzv|select Kurzname, kzvnummer, @{N='CSVName';E={$_.PreisCSVLink.Segments[-1]}} | sort-object Kurzname

# Preise können generell von KZV, Innung, oder Softwarehersteller kommen
# in der Regel in der Form CSV-, PDF- oder Excel-Datei

# für XLSX-Dateien kann man das Modul ImportExcel verwenden, ist Core kompatibel
# Install-Module ImportExcel -Scope CurrentUser

# eine bestimmte KZV auswählen
$kzv = $k.kzv|where name -eq "Thüringen"

