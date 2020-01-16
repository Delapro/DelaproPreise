# Requires PS7

$url='https://www.kzv-lsa.de/files/KZV-SA/Inhalte/Dokumente/BEL-Preisliste/2020/54la0120.csv'
$sa=Invoke-WebRequest -UseBasicParsing -Uri $url
$bsa=$sa.RawContentStream.ToArray()   # muss sein, sonst gibt es Streß mit Umlauten
$saCSV=[System.Text.Encoding]::UTF8.GetString($bsa)
$saCsvNew=$saCSV -split "`n"
$VDDSCSVHeader = @('Kürzel','Nr','Bezeichnung','Kassenart','PreisPraxisLabor','PreisGewerbeLabor', 'PreisZEPraxis', 'PreisZEGeewerbe', 'PreisKFOPraxis', 'PreisKFOGewerbe', 'PreisKBPraxis', 'PreisKBGewerbe', 'PreisPAPraxis', 'PreisPAGewerbe')
$VDDSHeaderConvert = @('Kürzel','Nr','Bezeichnung','Kassenart','@{N=''PreisGewerbeLabor'';E={[decimal]$_.PreisGewerbeLabor.replace('','',''.'')}}')

$saPreise=$saCsvNew|ConvertFrom-Csv -Delimiter ';' -Header $VDDSCSVHeader
$saPreise|Out-GridView

