#Requires -Version 7

$KZVBereich = @('Baden-Württemberg', 'Bayern', 'Hessen', 'Rheinland-Pfalz', 'Saarland, ''Nordrhein', 'Westfalen', 'Niedersachsen', 'Bremen', 'Hamburg', 'Schleswig-Holstein', 'Mecklenburg-Vorpommern', 'Brandenburg', 'Berlin', 'Sachsen', 'Sachsen-Anhalt', 'Thüringen')
$Kassenart = @{'Alle Kassen'=0; 'Primärkassen'=1; 'Ersatzkassen'=2}

Function GetAllLinksForFileExtension {
    [CmdletBinding()]
    [OutputType([uri[]])]
    Param (
        [parameter(Mandatory=$true)]
        [uri]$root,
        [parameter(Mandatory=$true)]
        [uri]$site,
        [parameter(Mandatory=$true)]
        [string]$fileExtension
    )

    [uri[]]$links=$null

    if ($site) {
        $links = (Invoke-WebRequest -uri $site).links.href| Where-Object {$_ -match $fileExtension}
        If ($links) {
            $links = $links | ForEach-Object {Join-Uri -Link $_ -Root $root}
        }
    }
    return $links
}

# wegen Problemen mit Komprimieren von URIs: https://github.com/dotnet/runtime/issues/31300
# gibt es diese Routine, welche unnötige / beim Zusammenketten von Urls entfernt
Function Join-Uri {
    [CmdletBinding()]
    [OutputType([uri])]
    Param (
        [Uri]$Link,
        [Uri]$Root
    )

    If ($Link.IsAbsoluteUri) {
        $Link
    } else {
        If (($Root.OriginalString.Substring($Root.OriginalString.Length -1) -eq '/') -and
              ($Link.OriginalString.Substring(0, 1) -eq '/')) {
                [uri]"$($Root.OriginalString.Substring(0,$Root.OriginalString.LastIndexOf('/')))$($Link.OriginalString)"
        } else {
            [uri]"$($Root.OriginalString)$($Link.OriginalString)"
        }
    }
}

class Preis {
    [DateTime]$GueltigAb
    [Uri]$PreisCSVLink
    [Uri]$PreisPDFLink
    [Uri]$PreisKFOCSVLink
    [Uri]$PreisKFOPDFLink
    [System.Text.Encoding]$Encoding
    [System.IO.FileInfo]$AlternateFile

}

class Preise {
    # ``1 ist wichtig, sonst gibts Stress bei Add. ``1 bedeutet einen Typ in der generischen Liste
    [System.Collections.Generic.List``1[Preis]]$Preis

    Preise ([DateTime]$GueltigAb, $PreisCSVLink) {
        If ($null -eq $this.Preis) {
            $this.Preis = [System.Collections.Generic.List``1[Preis]]::new()
        }
        $p = [Preis]::new()
        $p.PreisCSVLink = $PreisCSVLink
        $this.Preis.Add($p)
    }
    Preise ([DateTime]$GueltigAb, $PreisCSVLink, $PreisPDFLink) {
        If ($null -eq $this.Preis) {
            $this.Preis = [System.Collections.Generic.List``1[Preis]]::new()
        }
        $p = [Preis]::new()
        $p.PreisCSVLink = $PreisCSVLink
        $p.PreisPDFLink = $PreisPDFLink
        $this.Preis.Add($p)
    }

    Add ([Preis]$preis) {
        $this.Preis.Add($preis)
    }
}

class KZV {
    [String]$Name
    [String]$Kurzname  # wie von Delapro verwendet
    [String]$KZVNummer
    [Uri]$Homepage
    [Uri]$HomepagePreise
    [Uri]$PreisCSVLink
    [Uri]$PreisPDFLink
    [System.Text.Encoding]$Encoding
    [System.IO.FileInfo]$AlternateFile
    [hashtable]$RequestHeaders

    KZV ($Name, $Kurzname, $KZVNummer, $Homepage, $HomepagePreise, $PreisCSVLink) {
        $this.Name = $Name
        $this.KZVNummer = $KZVNummer
        $this.Kurzname = $Kurzname
        $this.Homepage = $Homepage
        $this.HomepagePreise = $HomepagePreise
        $this.PreisCSVLink = $PreisCSVLink
    }

    KZV ($Name, $Kurzname, $KZVNummer, $Homepage, $HomepagePreise, $PreisCSVLink, $PreisPDFLink) {
        $this.Name = $Name
        $this.KZVNummer = $KZVNummer
        $this.Kurzname = $Kurzname
        $this.Homepage = $Homepage
        $this.HomepagePreise = $HomepagePreise
        $this.PreisCSVLink = $PreisCSVLink
        $this.PreisPDFLink = $PreisPDFLink
    }

    [bool]HomepageErreichbar () {
        [bool]$erg = $false;

        try {
            $erg = (Invoke-WebRequest -Uri $this.Homepage -Headers $this.RequestHeaders).StatusCode -eq 200
        } catch {
            $erg = $false
        }
        return $erg
    }

    [bool]HomepagePreiseErreichbar () {
        [bool]$erg = $false;

        if ($this.HomepagePreise) {
            try {
                $erg = (Invoke-WebRequest -Uri $this.HomepagePreise -Headers $this.RequestHeaders).StatusCode -eq 200
            } catch {
                $erg = $false
            }
        }
        return $erg
    }

    [uri[]]GetPDFPreiseLinks () {
        return GetAllLinksForFileExtension -root $this.Homepage -site $this.HomepagePreise -fileExtension '.pdf'
    }

    [uri[]]GetCSVPreiseLinks () {
        return GetAllLinksForFileExtension -root $this.Homepage -site $this.HomepagePreise -fileExtension '.csv'
    }
}

class KZVen {
    [KZV[]]$KZV

    KZVen () {
        $this.KZV = [KZV[]]::new(17)

        # hier sollten immer die Links zu allen aktuellen KZVen stehen:
        # https://www.kzbv.de/bundeseinheitliches-kassenverzeichnis-bkv.951.de.html
        # ahuch hier gibts Infos und Links: https://www.vdds.de/schnittstellen/labor-preise/#top

        # TODO: Verweise auf CSV- und PDF-Preislisten entkoppeln, damit die Jahre zugeordnet werden können
        # TODO: Splitting von CSV- und PDF-Preislisten nach ZE und KFO vorsehen, z. B. hat Bayern in 2020 zwei PDF-Dateien obwohl nur eine CSV
        $this.KZV[0] = [KZV]::new('Baden-Württemberg', 'BaWu', '02', 'http://www.kzvbw.de/site/', '', '')

        # KFO: https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/BEL_KB_KFO_012020.pdf
        $this.KZV[1] = [KZV]::new('Bayern', 'Baye', '11', 'https://www.kzvb.de/', 'https://www.kzvb.de/abrechnung/bel-preise', 'https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/11la0120.csv', 'https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/BEL_ZE_012020.pdf')

        $this.KZV[2] = [KZV]::new('Berlin', 'Berl', '30', 'https://www.kzv-berlin.de/', 'https://www.kzv-berlin.de/praxis-service/abrechnung/bel-ii-laborpreise/', 'https://www.kzv-berlin.de/fileadmin/user_upload/Praxis-Service/1_Abrechnung/8_BEL_II__Laborpreise/30la0320.csv', 'https://www.kzv-berlin.de/fileadmin/user_upload/Praxis-Service/1_Abrechnung/8_BEL_II__Laborpreise/Laborpreise_seit_2020_03_01.pdf')
        $this.KZV[3] = [KZV]::new('Brandenburg', 'Bran', '53', 'http://www.kzvlb.de/', 'https://verwaltung.kzvlb.de/info.php', '', 'https://www.kzvlb.de/fileadmin/user_upload/sw/BEL_II_GewerblicheLabore_2003.pdf')
        $this.KZV[4] = [KZV]::new('Bremen', 'Brem', '31', 'https://www.kzv-bremen.de/', 'https://www.kzv-bremen.de/praxis/abrechnung/bel', 'https://www.kzv-bremen.de/praxis/abrechnung/files/31la0120.csv', 'https://www.kzv-bremen.de/praxis/files/21_zahntechnikvertraege.pdf')
        # TODO: Hamburg hat Splitting von ZE und KFO!!
        $this.KZV[5] = [KZV]::new('Hamburg', 'Hamb', '32', 'https://www.zahnaerzte-hh.de/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/praxis/abrechnung/kassenabrechnung-kzv/punktwerte-laborpreise-bel-materialkosten/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/mediathek/download-center/geschuetztes-dokument/file/download/36337/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/mediathek/download-center/geschuetztes-dokument/file/download/36331/')
        $this.KZV[6] = [KZV]::new('Hessen', 'Hess', '20', 'https://www.kzvh.de/index.html', '', '')
        $this.KZV[7] = [KZV]::new('Mecklenburg-Vorpommern', 'MeVo', '52', 'http://www.kzvmv.de/', '', '')
        $this.KZV[8] = [KZV]::new('Niedersachsen', 'Nisa', '04', 'https://www.kzvn.de/', '', '')
        $this.KZV[9] = [KZV]::new('Nordrhein', 'Nord', '13', 'http://www.kzvnr.de/', '', '')
        # WL dazu noch KFO: https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/Laborpreisliste_KFO_2021.pdf
        $this.KZV[10] = [KZV]::new('Westfalen-Lippe', 'West', '37', 'https://www.zahnaerzte-wl.de/', 'https://www.zahnaerzte-wl.de/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte.html', 'https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/37la0121.csv', 'https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/Laborpreisliste_ZE_2021.pdf')
        # alte Preise 2019 RPF: https://www.kzvrlp.de/fileadmin/KZV/Downloads/Mitglieder/Abrechnung/Aktuelles/BEL-Preise/06la0119.csv
        $this.KZV[11] = [KZV]::new('Rheinland-Pfalz', 'RhPf', '06', 'https://www.kzvrlp.de/', 'https://www.kzvrlp.de/mitglieder/abrechnung/bel-ii/', 'https://www.kzvrlp.de/fileadmin/KZV/Downloads/Mitglieder/Abrechnung/Aktuelles/BEL-Preise/06la0120.csv', 'https://www.kzvrlp.de/index.php?eID=tx_securedownloads&p=256&u=0&g=0&t=1579339965&hash=7263f05c3c30c30ec966430fc46de459d006977e&file=fileadmin/KZV/Downloads/Mitglieder/Abrechnung/Aktuelles/BEL-Preise/Kurzfassung_BEL_Preise_01.01.2020.pdf')
        $this.KZV[12] = [KZV]::new('Saarland', 'Saar', '35', 'https://www.zahnaerzte-saarland.de', 'https://www.zahnaerzte-saarland.de/praxisteam/index.php?idx=4&idxx=14', '')
        
        # KFO: https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/2020schnelluebersicht_laborpreise_paragraph_88.pdf
        $this.KZV[13] = [KZV]::new('Sachsen', 'Sach', '56', 'https://www.zahnaerzte-in-sachsen.de/', 'https://www.zahnaerzte-in-sachsen.de/praxis/abrechnung/zahntechnik/', 'https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/56la0220.csv', 'https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/2020schnelluebersicht_laborpreise_paragraph_57.pdf')

        $this.KZV[14] = [KZV]::new('Sachsen-Anhalt', 'SaAn', '54', 'https://www.kzv-lsa.de/', 'https://www.kzv-lsa.de/f%C3%BCr-die-praxis/abrechnung/bel-liste.html', 'https://www.kzv-lsa.de/files/Inhalte/Abrechnung/BEL/2020/54la0120.csv', 'https://www.kzv-lsa.de/files/Inhalte/Abrechnung/BEL/2020/HB_Fach_5.3_Hoechstpreisliste.pdf')
        # Sachsen-Anhalt Homepage benötigt zwingend diesen Header:
        $this.KZV[14].RequestHeaders = @{"Accept-Language"="de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7"}

        $this.KZV[15] = [KZV]::new('Schleswig-Holstein', 'SHol', '36', 'http://www.kzv-sh.de/', 'https://www.kzv-sh.de/fuer-die-praxis/abrechnung/bel-ii/', '')
        $this.KZV[16] = [KZV]::new('Thüringen', 'Thue', '55', 'https://www.kzvth.de/', 'https://www.kzvth.de/bel-beb', '')
    }
}

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
    $VDDSCSVHeader = @('Kürzel','Nr','Bezeichnung','Kassenart','PreisPraxisLabor','PreisGewerbeLabor', 'PreisZEPraxis', 'PreisZEGeewerbe', 'PreisKFOPraxis', 'PreisKFOGewerbe', 'PreisKBPraxis', 'PreisKBGewerbe', 'PreisPAPraxis', 'PreisPAGewerbe')
} else {
    $VDDSCSVHeader = @('Kürzel','Nr','Bezeichnung','Kassenart','PreisPraxisLabor','PreisGewerbeLabor', 'PreisZEPraxis', 'PreisZEGeewerbe', 'PreisKFOPraxis', 'PreisKFOGewerbe', 'PreisKBPraxis', 'PreisKBGewerbe', 'PreisPAPraxis', 'PreisPAGewerbe')
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
