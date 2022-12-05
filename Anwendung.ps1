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
# Homepage mit Preisen öffnen
Start-Process $kzv.HomepagePreise

If ($kzv.PreisCSVLink) {
    $url=[uri]$kzv.PreisCSVLink
    If ($url.Segments[-1].ToLower() -match '.csv$') {
        # CSV-Dateiname ist direkt in der URL angegeben
        $url.Segments[-1]
    }
    $sa=Invoke-WebRequest -UseBasicParsing -Uri $url
    $bsa=$sa.RawContentStream.ToArray()   # muss sein, sonst gibt es Streß mit Umlauten
    $saCSV=[System.Text.Encoding]::UTF8.GetString($bsa)
} else {
    # alternativ aus Datei einlesen
    $saCSV=Get-Content .\55la0320.csv -Encoding oem
}
$saCsvNew=$saCSV -split "`n"
# Anzahl der Felder ermitteln
$Spalten = ($saCsvNew[0] -split ';').count
If ($Spalten -eq 14) {
    # Standardformat
    $VDDSCSVHeader = @('Kürzel','Nr','Bezeichnung','Kassenart','PreisPraxisLabor','PreisGewerbeLabor', 'PreisZEPraxis', 'PreisZEGewerbe', 'PreisKFOPraxis', 'PreisKFOGewerbe', 'PreisKBPraxis', 'PreisKBGewerbe', 'PreisPAPraxis', 'PreisPAGewerbe')
} else {
    $VDDSCSVHeader = @('Kürzel','Nr','Bezeichnung','Kassenart','PreisPraxisLabor','PreisGewerbeLabor', 'PreisZEPraxis', 'PreisZEGewerbe', 'PreisKFOPraxis', 'PreisKFOGewerbe', 'PreisKBPraxis', 'PreisKBGewerbe', 'PreisPAPraxis', 'PreisPAGewerbe')
}
# möglichen Header entfernen
If ($saCsvNew[0] -match 'BEL') {
    $saCsvNew=$saCsvNew[1..($saCsvNew.Length)-1]
}
# Leereintragungen entfernen
$saCsvNew = $saCsvNew | where {$_ -Match ';;;;'}

$saPreise=$saCsvNew|ConvertFrom-Csv -Delimiter ';' -Header $VDDSCSVHeader
$saPreise|Out-GridView


$VDDSHeaderConvert = @('Kürzel','Nr','Bezeichnung','Kassenart',@{N='PreisGewerbeLabor';E={[decimal]$_.PreisGewerbeLabor.replace(',','.')}})
$saPNeu = $saPreise| select -Property $VDDSHeaderConvert

$saPNeu| measure -Property preisgewerbelabor -AllStats

# URL-Check der Homepage
# alte Variante: $k.kzv|select kzvnummer, Name, Homepage, @{N='Erreichbar';E={(Invoke-WebRequest -Uri $_.Homepage -Headers $_.RequestHeaders).StatusCode -eq 200}} | sort kzvnummer
$k.kzv|select kzvnummer, Name, Homepage, @{N='Erreichbar';E={$_.HomepageErreichbar()}} | sort kzvnummer

$k.kzv|select kzvnummer, Name, HomepagePreise, @{N='Erreichbar';E={$_.HomepagePreiseErreichbar()}} | sort kzvnummer

# alle PDF-Links einer Seite ermitteln:
# alte Variante: (Invoke-WebRequest -uri $k.kzv[-3].HomepagePreise).links.href| where {$_ -match '.pdf'}
$k.KZV[-3].GetPdfPreiseLinks() | Select AbsoluteUri

# alle Links auf CSV-Dateien einer Seite ermitteln:
# alte Variante: (Invoke-WebRequest -uri $k.kzv[-3].HomepagePreise).links.href| where {$_ -match '.csv'}
$k.KZV[-3].GetCsvPreiseLinks() | Select AbsoluteUri

# Nordrhein Praxispreise: https://www.kzvnr.de/medien/PDFs/Zahn%C3%A4rzteseite/BEL-Listen_II__Zahnersatz_/BEL_II_ab_01.01.2020.csv

#VDZI: https://www.vdzi.net/fileadmin/user_uploads/downloads/pdf/betriebswirtschaft/BEL/BEL_Baden-Wuerttemberg.pdf
