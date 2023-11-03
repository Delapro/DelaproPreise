# Anwendung
$k=[kzven]::new()
"Nach Kzv-Nummer sortiert:"
$k.kzv|select kzvnummer, Name, Homepage | sort-object kzvnummer
"Nach Name sortiert:"
$k.kzv|select kzvnummer, Name, Homepage | sort-object Name
"Nach Kurzname sortiert mit CSVLink:"
$k.kzv|select Kurzname, PreisCSVLink | sort-object Kurzname |fl
"Nach Kurzname sortiert mit CSV-Dateiname:"
$k.kzv|select Kurzname, kzvnummer, @{N='CSVName';E={$_.PreisCSVLink.Segments[-1]}} | sort-object Kurzname |fl
"Ausgabe der Preisseiten, nach Name sortiert:"
$k.kzv|select kzvnummer, Name, HomepagePreise | sort-object Name

# Preise können generell von KZV, Innung, oder Softwarehersteller kommen
# in der Regel in der Form CSV-, PDF- oder Excel-Datei

# für XLSX-Dateien kann man das Modul ImportExcel verwenden, ist Core kompatibel
# Install-Module ImportExcel -Scope CurrentUser

# eine bestimmte KZV auswählen
$kzv = $k.kzv|where name -eq "Thüringen"

"(Get-Host).Ui.RawUI:"
(Get-host).ui.RawUI
