# Anwendung
$k=[kzven]::new()
$k.kzv|select kzvnummer, Name, Homepage | sort kzvnummer
$k.kzv|select kzvnummer, Name, Homepage | sort Name
$k.kzv|select Kurzname, PreisCSVLink | sort Kurzname
# $k.kzv|select Kurzname, kzvnummer, @{N='CSVName';E={$_.PreisCSVLink.Segments[-1]}} | sort Kurzname
# bessere Variante die CSV-Dateinamen zu ermitteln, wegen indirekten Links:
$k.kzv|select Kurzname, kzvnummer, @{N='CSVName';E={$_.GetCSVName()}} | sort Kurzname

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

# Alternative wenn keine Kassenart mitangegeben wurde und allgemeinen Preisheader:
$SaAn=Get-Content .\54la0124.csv|ConvertFrom-Csv -Delimiter ';' -Header @('belnr','nr','bez','p1','p2','p3','p4','p5','p6','p7','p8','p9','p10','p11')|select belnr, nr, bez, @{n='p1';e={[decimal]$_.p1.replace(',','.')}}, @{n='p2';e={[decimal]$_.p2.replace(',','.')}}, @{n='p3';e={[decimal]$_.p3.replace(',','.')}}, @{n='p4';e={[decimal]$_.p4.replace(',','.')}}, @{n='p5';e={[decimal]$_.p5.replace(',','.')}}, @{n='p6';e={[decimal]$_.p6.replace(',','.')}}, @{n='p7';e={[decimal]$_.p7.replace(',','.')}}, @{n='p8';e={[decimal]$_.p8.replace(',','.')}}, @{n='p9';e={[decimal]$_.p9.replace(',','.')}}, @{n='p10';e={[decimal]$_.p10.replace(',','.')}}, @{n='p11';e={[decimal]$_.p11.replace(',','.')}}

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
