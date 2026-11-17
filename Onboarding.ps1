<#
.SYNOPSIS
    Legt neue Mitarbeitende aus einer CSV-Datei in Microsoft Entra ID an.

.DESCRIPTION
    Für jede Zeile der CSV-Datei wird ein Benutzerkonto nach der Namenskonvention
    vorname.nachname@<Domain> angelegt. Abteilung, Position und Nutzungsstandort werden gesetzt.

    Gruppen werden bewusst NICHT im Skript vergeben: Die Mitgliedschaften entstehen automatisch
    über die dynamischen Gruppen, deren Regel auf das Attribut "department" (Abteilung) prüft.

    Die Startkennwörter werden in eine Ergebnisdatei geschrieben. Das ist nur für die
    Testumgebung gedacht; produktiv würde man z. B. einen Temporary Access Pass verwenden.

.PARAMETER CsvPfad
    Pfad zur CSV-Datei (Trennzeichen Semikolon, Spalten: Vorname;Nachname;Abteilung;Position).

.PARAMETER Domain
    Domain für den Benutzernamen.

.EXAMPLE
    .\Onboarding.ps1 -CsvPfad .\neue-mitarbeitende.csv -WhatIf
    Testlauf: zeigt nur an, was passieren würde.

.EXAMPLE
    .\Onboarding.ps1 -CsvPfad .\neue-mitarbeitende.csv
    Legt die Benutzer tatsächlich an.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$CsvPfad,

    [string]$Domain = "nordlichtlab.onmicrosoft.com",

    [string]$Nutzungsstandort = "DE"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

# Macht aus "Krüger" den Kontonamen "krueger" (Kleinbuchstaben, Umlaute ersetzt, Sonderzeichen entfernt)
function ConvertTo-Kontoname {
    param([string]$Text)
    $t = $Text.Trim().ToLower()
    $t = $t -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'ß', 'ss'
    $t = $t -replace '[^a-z0-9\-]', ''
    return $t
}

# Erzeugt ein zufälliges Startkennwort mit 16 Zeichen aus allen vier Zeichengruppen
function New-Startkennwort {
    $gross  = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $klein  = 'abcdefghijkmnopqrstuvwxyz'.ToCharArray()
    $ziffer = '23456789'.ToCharArray()
    $sonder = '!$%&?#+-'.ToCharArray()
    $alle   = $gross + $klein + $ziffer + $sonder

    $zeichen = @()
    $zeichen += $gross  | Get-Random -Count 2
    $zeichen += $klein  | Get-Random -Count 2
    $zeichen += $ziffer | Get-Random -Count 2
    $zeichen += $sonder | Get-Random -Count 2
    $zeichen += $alle   | Get-Random -Count 8

    return -join ($zeichen | Sort-Object { Get-Random })
}

# ---------------------------------------------------------------------------
# Vorprüfungen
# ---------------------------------------------------------------------------

if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
    throw "Das Microsoft-Graph-Modul ist nicht installiert. Siehe Anleitung, Teil A."
}
if (-not (Get-MgContext)) {
    throw "Keine Verbindung zu Microsoft Graph. Bitte zuerst Connect-MgGraph ausführen."
}
if (-not (Test-Path $CsvPfad)) {
    throw "CSV-Datei nicht gefunden: $CsvPfad"
}

$csvVoll    = (Resolve-Path $CsvPfad).Path
$eintraege  = Import-Csv -Path $csvVoll -Delimiter ';' -Encoding UTF8
$ergebnisse = @()

Write-Host "`nOnboarding für $($eintraege.Count) Einträge aus $csvVoll`n" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Verarbeitung
# ---------------------------------------------------------------------------

foreach ($e in $eintraege) {
    $vorname   = $e.Vorname.Trim()
    $nachname  = $e.Nachname.Trim()
    $abteilung = $e.Abteilung.Trim()
    $upn       = "{0}.{1}@{2}" -f (ConvertTo-Kontoname $vorname), (ConvertTo-Kontoname $nachname), $Domain

    # Gibt es den Benutzer schon? Dann überspringen (das Skript kann gefahrlos mehrfach laufen)
    $vorhanden = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
    if ($vorhanden) {
        Write-Warning "$upn existiert bereits und wird übersprungen."
        $ergebnisse += [pscustomobject]@{ Benutzer = $upn; Abteilung = $abteilung; Status = "Übersprungen (existiert bereits)"; Startkennwort = "" }
        continue
    }

    $kennwort  = New-Startkennwort
    $parameter = @{
        AccountEnabled    = $true
        DisplayName       = "$vorname $nachname"
        GivenName         = $vorname
        Surname           = $nachname
        UserPrincipalName = $upn
        MailNickname      = $upn.Split('@')[0]
        Department        = $abteilung
        JobTitle          = $e.Position.Trim()
        UsageLocation     = $Nutzungsstandort
        PasswordProfile   = @{
            Password                      = $kennwort
            ForceChangePasswordNextSignIn = $true
        }
    }

    if ($PSCmdlet.ShouldProcess($upn, "Benutzer anlegen (Abteilung: $abteilung)")) {
        try {
            New-MgUser @parameter -ErrorAction Stop | Out-Null
            Write-Host "Angelegt: $upn ($abteilung)" -ForegroundColor Green
            $ergebnisse += [pscustomobject]@{ Benutzer = $upn; Abteilung = $abteilung; Status = "Angelegt"; Startkennwort = $kennwort }
        }
        catch {
            Write-Warning "Fehler bei ${upn}: $($_.Exception.Message)"
            $ergebnisse += [pscustomobject]@{ Benutzer = $upn; Abteilung = $abteilung; Status = "Fehler: $($_.Exception.Message)"; Startkennwort = "" }
        }
    }
}

# ---------------------------------------------------------------------------
# Ergebnis
# ---------------------------------------------------------------------------

if ($ergebnisse.Count -gt 0) {
    $ergebnisse | Format-Table Benutzer, Abteilung, Status -AutoSize

    if ($WhatIfPreference) {
        Write-Host "Testlauf (-WhatIf): Es wurde nichts angelegt und keine Ergebnisdatei geschrieben.`n" -ForegroundColor Yellow
    }
    else {
        $datei = Join-Path (Split-Path $csvVoll) ("Onboarding-Ergebnis_{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmm"))
        $csvOptionen = @{ Path = $datei; Delimiter = ';' }
        if ($PSVersionTable.PSVersion.Major -ge 6) { $csvOptionen.Encoding = 'utf8BOM' } else { $csvOptionen.Encoding = 'UTF8'; $csvOptionen.NoTypeInformation = $true }
        $ergebnisse | Export-Csv @csvOptionen
        Write-Host "Ergebnis gespeichert: $datei" -ForegroundColor Cyan
        Write-Host "Achtung: Die Datei enthält Startkennwörter. Nach der Weitergabe löschen." -ForegroundColor Yellow
        Write-Host "Die Gruppenmitgliedschaften entstehen automatisch über die dynamischen Gruppen (kann einige Minuten dauern).`n"
    }
}
