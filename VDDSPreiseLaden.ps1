#Requires -Version 7

$KZVBereich = @('Baden-Württemberg'
    ,'Bayern'
    ,'Hessen'
    ,'Rheinland-Pfalz'
    ,'Saarland'
    ,'Nordrhein'
    ,'Westfalen'
    ,'Niedersachsen'
    ,'Bremen'
    ,'Hamburg'
    ,'Schleswig-Holstein'
    ,'Mecklenburg-Vorpommern'
    ,'Brandenburg'
    ,'Berlin'
    ,'Sachsen'
    ,'Sachsen-Anhalt'
    ,'Thüringen'
)
$Kassenart = @{'Alle Kassen'=0
    ;'Primärkassen'=1
    ;'Ersatzkassen'=2
}
$LeistungsgruppenBEL2 = @{'Arbeitsvorbereitung'=0
    ;'Festsitzender Zahnersatz'=1
    ;'Herausnehmbarer Zahnersatz aus Legierungen (Modellguss)'=2
    ;'Herausnehmbarer Zahnersatz aus Kunststoff'=3
    ;'Aufbissbehelfe'=4
    ;'Unterkieferprotrusionsschiene'=5
    ;'Kieferorthopädie'=7
    ;'Wiederherstellungen und Erweiterungen'=8
    ;'Materialien und Sonstiges'=9
}

# ermittelt alle Href Links auf einer HTML-Seite die eine bestimmte Dateiendung haben
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

    if (-not ([string]::IsNullOrEmpty($site))) {
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
    [System.IO.FileInfo]$AlternateFile  # kann verwendet werden um lokale Dateien einbinden zu können
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

    # gibt alle Links zu PDF-Dateien zurück
    [uri[]]GetPDFPreiseLinks () {
        return GetAllLinksForFileExtension -root $this.Homepage -site $this.HomepagePreise -fileExtension '.pdf'
    }

    # gibt alle Links zu CSV-Dateien zurück
    [uri[]]GetCSVPreiseLinks () {
        return GetAllLinksForFileExtension -root $this.Homepage -site $this.HomepagePreise -fileExtension '.csv'
    }
    
    [string]GetCSVName () {
        $CsvName = ""
        If ($this.PreisCSVLink.Segments.Length -gt 0) {
            $CsvName = $this.PreisCSVLink.Segments[-1]
            If (-Not ($CsvName -match '\.csv')) {
                # Prüfen, ob es ein CSV-Dateiname mittels Header ausgelesen werden kann
                $r = Invoke-WebRequest $this.PreisCSVLink -UseBasicParsing
                $disposition = [System.Net.Mime.ContentDisposition]::new($r.Headers.'Content-Disposition')
                $CsvName = $disposition.Parameters['filename']
            }
        }
        return $CsvName
    }
}

class KZVen {
    [KZV[]]$KZV

    KZVen () {
        $this.KZV = [KZV[]]::new(17)

        # hier sollten immer die Links zu allen aktuellen KZVen stehen:
        # https://www.kzbv.de/bundeseinheitliches-kassenverzeichnis-bkv.951.de.html
        # auch hier gibts Infos und Links: https://www.vdds.de/schnittstellen/labor-preise/#top

        # TODO: Verweise auf CSV- und PDF-Preislisten entkoppeln, damit die Jahre zugeordnet werden können
        # TODO: Splitting von CSV- und PDF-Preislisten nach ZE und KFO vorsehen, z. B. hat Bayern in 2020 zwei PDF-Dateien obwohl nur eine CSV
        $this.KZV[0] = [KZV]::new('Baden-Württemberg', 'BaWu', '02', 'https://www.kzvbw.de/', 'https://www.kzvbw.de/zahnaerzte/abrechnung/punktwerte-formulare-vordrucke/bel-leistungen-download/', 'https://www.kzvbw.de/wp-content/uploads/02la0122.csv', 'https://www.kzvbw.de/wp-content/uploads/Preisliste-Gewerbelabor-und-Praxislabor-2022.pdf')

        # KFO: https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/BEL_Preise_KFO_KB_012021.pdf
        $this.KZV[1] = [KZV]::new('Bayern', 'Baye', '11', 'https://www.kzvb.de/', 'https://www.kzvb.de/abrechnung/bel-preise', 'https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/11la0122.csv', 'https://www.kzvb.de/fileadmin/user_upload/Abrechnung/BEL/BEL_Preise_ZE_012022.pdf')

        $this.KZV[2] = [KZV]::new('Berlin', 'Berl', '30', 'https://www.kzv-berlin.de/', 'https://www.kzv-berlin.de/praxis-service/abrechnung/bel-ii-laborpreise/', 'https://www.kzv-berlin.de/fileadmin/user_upload/Praxis-Service/1_Abrechnung/8_BEL_II__Laborpreise/30la0222.csv', 'https://www.kzv-berlin.de/fileadmin/user_upload/Praxis-Service/1_Abrechnung/8_BEL_II__Laborpreise/Laborpreise_ab_2022_02_01.pdf')
        $this.KZV[3] = [KZV]::new('Brandenburg', 'Bran', '53', 'https://www.kzvlb.de/', 'https://verwaltung.kzvlb.de/info.php', 'https://verwaltung.kzvlb.de/sw/53la0222.csv', 'https://verwaltung.kzvlb.de/sw/BEL_II_2202.pdf')
        $this.KZV[4] = [KZV]::new('Bremen', 'Brem', '31', 'https://www.kzv-bremen.de/', 'https://www.kzv-bremen.de/praxis/abrechnung/bel', 'https://www.kzv-bremen.de/praxis/abrechnung/files/31la0122.csv', 'https://www.kzv-bremen.de/praxis/files/21_zahntechnikvertraege.pdf')
        # TODO: Hamburg hat Splitting von ZE und KFO!!
        $this.KZV[5] = [KZV]::new('Hamburg', 'Hamb', '32', 'https://www.zahnaerzte-hh.de/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/praxis/abrechnung/kassenabrechnung-kzv/punktwerte-laborpreise-bel-materialkosten/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/mediathek/download-center/geschuetztes-dokument/file/download/44877/', 'https://www.zahnaerzte-hh.de/zahnaerzte-portal/mediathek/download-center/geschuetztes-dokument/file/download/44875/')
        $this.KZV[6] = [KZV]::new('Hessen', 'Hess', '20', 'https://www.kzvh.de/index.html', 'https://www.kzvh.de/BEL-Preisliste/index.html', 'https://www.kzvh.de/wcm/idc/groups/public/documents/web/mdiw/bgew/~edisp/20la0122.csv')

        # MeVo: die Preise sind unter BKV-Download zu finden!
        $this.KZV[7] = [KZV]::new('Mecklenburg-Vorpommern', 'MeVo', '52', 'https://www.kzvmv.de/', 'https://www.kzvmv.de/bkv-download/', 'https://www.kzvmv.de/dokumente/52la0122.csv')
        $this.KZV[8] = [KZV]::new('Niedersachsen', 'Nisa', '04', 'https://www.kzvn.de/', '', '')
        $this.KZV[9] = [KZV]::new('Nordrhein', 'Nord', '13', 'https://www.kzvnr.de/', '', '')
        # WL dazu noch KFO: https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/Laborpreisliste_KFO_2021.pdf
        $this.KZV[10] = [KZV]::new('Westfalen-Lippe', 'West', '37', 'https://www.zahnaerzte-wl.de/', 'https://www.zahnaerzte-wl.de/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte.html', 'https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/37la0121.csv', 'https://www.zahnaerzte-wl.de/images/kzvwl/praxisteam/abrechnung/aktuelle-abrechnungsinfos-fuer-vertragszahnaerzte/Laborpreisliste_ZE_2021.pdf')
        # alte Preise 2019 RPF: https://www.kzvrlp.de/fileadmin/KZV/Downloads/Mitglieder/Abrechnung/Aktuelles/BEL-Preise/06la0119.csv
        $this.KZV[11] = [KZV]::new('Rheinland-Pfalz', 'RhPf', '06', 'https://www.kzvrlp.de/', 'https://www.kzvrlp.de/mitglieder/abrechnung/bel-ii/', 'https://www.kzvrlp.de/securedl/sdl-eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzA0MTA2NjEsImV4cCI6MTY3MDUwMDY2MSwidXNlciI6MCwiZ3JvdXBzIjpbMCwtMV0sImZpbGUiOiJcL2ZpbGVhZG1pblwvS1pWXC9Eb3dubG9hZHNcL01pdGdsaWVkZXJcL0FicmVjaG51bmdcL0FrdHVlbGxlc1wvQkVMLVByZWlzZVwvMDZsYTA1MjIuY3N2IiwicGFnZSI6MzIyfQ.FN6gLVnNOHXHQcMn6eaJ0jSQQGRT8_x-6Xm2hXuHB20/06la0522.csv', 'https://www.kzvrlp.de/securedl/sdl-eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzAzMjUxMzMsImV4cCI6MTY3MDQxNTEzMywidXNlciI6MCwiZ3JvdXBzIjpbMCwtMV0sImZpbGUiOiJmaWxlYWRtaW5cL0taVlwvRG93bmxvYWRzXC9NaXRnbGllZGVyXC9BYnJlY2hudW5nXC9Ba3R1ZWxsZXNcL0JFTC1QcmVpc2VcL0t1cnpmYXNzdW5nX0JFTF9QcmVpc2VfMDEuMDUuMjAyMi5wZGYiLCJwYWdlIjozMjJ9.PsXnSmID9cn8oL61DCkSGxte09T6U98ahBK5bEmJ4IA/Kurzfassung_BEL_Preise_01.05.2022.pdf')
        $this.KZV[12] = [KZV]::new('Saarland', 'Saar', '35', 'https://www.zahnaerzte-saarland.de', 'https://www.zahnaerzte-saarland.de/praxisteam/index.php?idx=4&idxx=14', 'https://www.zahnaerzte-saarland.de/meine_kzv/data/35la0122.csv', 'https://www.zahnaerzte-saarland.de/meine_kzv/data/BEL_II_2022per01012022Kurztexte.pdf')
        
        # KFO: https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/2021schnelluebersicht_laborpreise_paragraph_88.pdf
        $this.KZV[13] = [KZV]::new('Sachsen', 'Sach', '56', 'https://www.zahnaerzte-in-sachsen.de/', 'https://www.zahnaerzte-in-sachsen.de/praxis/abrechnung/zahntechnik/', 'https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/56la0122.csv', 'https://www.zahnaerzte-in-sachsen.de/fileadmin/Praxis/KZVS/Abrechnung/BEL_II/verguetungsvereinbarung_zt_paragraph57.pdf')

        $this.KZV[14] = [KZV]::new('Sachsen-Anhalt', 'SaAn', '54', 'https://www.kzv-lsa.de/', 'https://www.kzv-lsa.de/f%C3%BCr-die-praxis/abrechnung/bel-liste.html', 'https://www.kzv-lsa.de/files/Inhalte/Abrechnung/BEL/2022/54la0822.csv', 'https://www.kzv-lsa.de/files/Inhalte/Abrechnung/BEL/2022/HB_Fach_5.3_Hoechstpreisliste.pdf')
        # Sachsen-Anhalt Homepage benötigt zwingend diesen Header:
        $this.KZV[14].RequestHeaders = @{"Accept-Language"="de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7"}

        # KFO: https://www.kzv-sh.de/wp-content/uploads/2021/01/BEL-II-Preislisten-KFO-01.01.2021-Gewerbelabor.pdf
        $this.KZV[15] = [KZV]::new('Schleswig-Holstein', 'SHol', '36', 'https://www.kzv-sh.de/', 'https://www.kzv-sh.de/fuer-die-praxis/abrechnung/bel-ii/', 'https://www.kzv-sh.de/wp-content/uploads/2022/01/36la0122.csv', 'https://www.kzv-sh.de/wp-content/uploads/2021/12/BEL-II-Preislisten-Regelversorgung-01.01.2022-Gewerbelabor.pdf')
        $this.KZV[16] = [KZV]::new('Thüringen', 'Thue', '55', 'https://www.kzvth.de/', 'https://www.kzvth.de/bel-beb', 'https://www.kzvth.de/services/asset/KZVTh/Downloadbereich/BEL/BEL%202022/55Ia0122.csv', 'https://www.kzvth.de/services/asset/KZVTh/Downloadbereich/BEL/BEL%202022/Bel_II_Kurzversion_ab_01012022_KZV.pdf')
    }
}

# Funktion zum Auslesen der Leistungsbeschreibungen aus den offizellen Beschreibungen
Function ConvertFrom-BEL2Beschreibung {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [String]$Block
    )
  
    Begin {}
  
    Process {
        $BeginLine = 0
        # Sonderfall, unbekanntes Zeichen wird in Bindestrich gewandelt
        $unbekannt=[char]61485
        $Block = $Block -replace $unbekannt,"-" 
        $BlockLines = $Block -split "`r?`n"  # wichtig " zu verwenden!
  
        $Teiltext = ""
        $Pos = $BlockLines | Select-String 'Kurztext laut Anlage 2'
        If ($Pos) {
            $LastLine = ($Pos.LineNumber-2) # immer -2 um die unnötigen Zeilen wegzubekommen
            $Teiltext = ($BlockLines[$BeginLine..$LastLine] | % {$_.Trim()} | Out-String).Trim()
            $BeginLine = $LastLine +2 # um die unnötigen Zeilen wegzubekommen
        }
        $Bezeichnung = $Teiltext
  
        $Teiltext = ""
        $Pos = $BlockLines | Select-String 'Erläuterung(?:en)? zum Leistungsinhalt'
        If ($Pos -is [array]) {
            $Pos = $Pos[0]
        }
        If ($Pos) {
            $LastLine = ($Pos.LineNumber-2) # immer -2 um die unnötigen Zeilen wegzubekommen
            $Teiltext = ($BlockLines[$BeginLine..$LastLine] | Out-String).Trim()
            $BeginLine = $LastLine +2 # um die unnötigen Zeilen wegzubekommen
        } else {
            $Pos = $BlockLines | Select-String 'Erläuterung(?:en)? zur Abrechnung'
            If ($Pos -is [array]) {
                $Pos = $Pos[0]
            }
            If ($Pos) {
                $LastLine = ($Pos.LineNumber-2) # immer -2 um die unnötigen Zeilen wegzubekommen
                $Teiltext = ($BlockLines[$BeginLine..$LastLine] | Out-String).Trim()
                $BeginLine = $LastLine +2 # um die unnötigen Zeilen wegzubekommen
            } else {
                # absoluter Sonderfall, keine Leistungs- und Abrechungsbeschreibung
                $TeilText = $BlockLines[-1].Trim()
                $BeginLine = $BlockLines.Length
            }
        }
        $Kurztext = $Teiltext
  
        $Teiltext = ""
        $Pos = $BlockLines | Select-String 'Erläuterung(?:en)? zur Abrechnung'
        If ($Pos -is [array]) {
            $Pos = $Pos[0]
        }
        If ($Pos) {
            $LastLine = ($Pos.LineNumber-2) # immer -2 um die unnötigen Zeilen wegzubekommen
            If ($BeginLine -le $LastLine) {
                $Teiltext = ($BlockLines[$BeginLine..$LastLine] | % {$_.Trim()} | Out-String).Trim()
            } else {
                #$Teiltext = $null
            }
            $BeginLine = $LastLine +2 # um die unnötigen Zeilen wegzubekommen
        } else {
            # evtl. Sonderfall weil es keine Erläuterung zur Abrechnung gibt
            $LastLine = $Blocklines.Length-1
            If ($BeginLine -le $LastLine) {
                $Teiltext = ($BlockLines[$BeginLine..$LastLine] | Out-String).Trim()
            }
        }
        $Leistungsinhalt = $Teiltext
  
        $Teiltext = ""
        $Pos = $BlockLines | Select-String 'Erläuterung(?:en)? zur Abrechnung'
        If ($Pos -is [array]) {
            $Pos = $Pos[0]
        }
        If ($Pos) {
            $LastLine = $Blocklines.Length-1
            $Teiltext = ($BlockLines[$BeginLine..$LastLine] | % {$_.Trim()} | Out-String).Trim()
        }
        $Abrechnung = $Teiltext
  
        $BelNummer = $Kurztext.Substring(0,5)
        $Verlinkt = $Abrechnung -replace $BelNummer,'' | Select-String '\d{3,3} \d{1,1}' -AllMatches
        If ($Verlinkt) {
            $Verlinkt = ($Verlinkt.Matches).Value
        } else {
            $Verlinkt = $null
        }
        [PSCustomObject]@{BelNummer=$BelNummer
                         ;Bezeichnung=($Bezeichnung -replace $BelNummer,'')
                         ;Kurztext=($Kurztext -replace $BelNummer,'').Trim()
                         ;Leistungsinhalt=$Leistungsinhalt
                         ;Abrechnung=$Abrechnung
                         ;Verlinkt=$Verlinkt}    
    } # Process
}

