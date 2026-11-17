<#
.SYNOPSIS
    Führt das technische Offboarding eines Benutzers in Microsoft Entra ID durch.

.DESCRIPTION
    Schritte:
      1. Anmeldung sperren
      2. Alle Sitzungen widerrufen (bestehende Tokens werden ungültig)
      3. Kennwort auf einen zufälligen Wert zurücksetzen
      4. Abteilung auf "Ausgeschieden" setzen -> dadurch endet die Mitgliedschaft in dynamischen Gruppen
      5. Mitgliedschaften in zugewiesenen Gruppen entfernen
         (Gruppen, über die eine Lizenz vergeben wird, werden bewusst ausgelassen, siehe unten)
      6. Admin-Rollen und Lizenzen prüfen und als Hinweis ausgeben
      7. Alle Schritte in einer Protokolldatei festhalten

    Benötigte Berechtigungen beim Verbinden: User.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All,
    User-PasswordProfile.ReadWrite.All (für das Zurücksetzen des Kennworts).

    Lizenzen werden bewusst NICHT automatisch entfernt: Wird die Lizenz entzogen, bevor das Postfach
    in ein freigegebenes Postfach umgewandelt wurde, wird das Postfach nach 30 Tagen gelöscht.

.PARAMETER UserPrincipalName
    Anmeldename des Benutzers, z. B. jonas.richter@nordlichtlab.onmicrosoft.com

.PARAMETER Force
    Ohne Sicherheitsabfrage ausführen.

.EXAMPLE
    .\Offboarding.ps1 -UserPrincipalName jonas.richter@nordlichtlab.onmicrosoft.com -WhatIf
    Testlauf: zeigt nur an, was passieren würde.

.EXAMPLE
    .\Offboarding.ps1 -UserPrincipalName jonas.richter@nordlichtlab.onmicrosoft.com
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,

    [string]$ProtokollPfad = ".\Offboarding-Protokoll.csv",

    [switch]$Force
)

$ErrorActionPreference = "Stop"

function New-Zufallskennwort {
    $alle = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!$%&?#+-'.ToCharArray()
    return 'Aa1!' + (-join ($alle | Get-Random -Count 24))
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

try {
    $benutzer = Get-MgUser -UserId $UserPrincipalName `
        -Property Id, DisplayName, UserPrincipalName, Department, AccountEnabled, AssignedLicenses, LicenseAssignmentStates
}
catch {
    throw "Benutzer $UserPrincipalName wurde nicht gefunden."
}

$id        = $benutzer.Id
$protokoll = @()
function Add-Protokoll([string]$Schritt, [string]$Ergebnis) {
    $script:protokoll += [pscustomobject]@{
        Zeitpunkt      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Benutzer       = $benutzer.UserPrincipalName
        Schritt        = $Schritt
        Ergebnis       = $Ergebnis
        AusgefuehrtVon = (Get-MgContext).Account
    }
}

Write-Host "`nOffboarding für $($benutzer.DisplayName) ($($benutzer.UserPrincipalName))" -ForegroundColor Cyan
Write-Host "Abteilung: $($benutzer.Department) | Konto aktiv: $($benutzer.AccountEnabled)`n"

if (-not $WhatIfPreference -and -not $Force) {
    $antwort = Read-Host "Offboarding wirklich durchführen? (j/n)"
    if ($antwort -ne 'j') { Write-Host "Abgebrochen."; return }
}

# Gruppen, über die der Benutzer eine Lizenz erhält
$lizenzGruppen = @($benutzer.LicenseAssignmentStates | Where-Object { $_.AssignedByGroup } | ForEach-Object { $_.AssignedByGroup })

# Führt einen Schritt aus, fängt Fehler ab und protokolliert das echte Ergebnis.
# Wichtig: Die Graph-Befehle brauchen -ErrorAction Stop, sonst würde ein Fehler nicht abgefangen.
$script:fehlerAnzahl = 0
function Invoke-Schritt {
    param([string]$Nr, [string]$Beschreibung, [string]$Erfolgstext, [scriptblock]$Aktion)
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, $Beschreibung)) {
        try {
            & $Aktion
            Write-Host "[$Nr] $Erfolgstext" -ForegroundColor Green
            Add-Protokoll $Beschreibung "OK"
        }
        catch {
            $meldung = ($_.Exception.Message -split "`n")[0]
            Write-Warning "[$Nr] $Beschreibung FEHLGESCHLAGEN: $meldung"
            Add-Protokoll $Beschreibung "FEHLER: $meldung"
            $script:fehlerAnzahl++
        }
    }
}

# ---------------------------------------------------------------------------
# 1. bis 4.: Konto absichern
# ---------------------------------------------------------------------------
Invoke-Schritt "1" "Anmeldung sperren" "Anmeldung gesperrt" {
    Update-MgUser -UserId $id -AccountEnabled:$false -ErrorAction Stop
}

Invoke-Schritt "2" "Alle Sitzungen widerrufen" "Sitzungen widerrufen" {
    Revoke-MgUserSignInSession -UserId $id -ErrorAction Stop | Out-Null
}

# Benötigt die Berechtigung User-PasswordProfile.ReadWrite.All (siehe Connect-MgGraph in der README)
Invoke-Schritt "3" "Kennwort auf Zufallswert zurücksetzen" "Kennwort zurückgesetzt" {
    Update-MgUser -UserId $id -PasswordProfile @{ Password = (New-Zufallskennwort); ForceChangePasswordNextSignIn = $true } -ErrorAction Stop
}

Invoke-Schritt "4" "Abteilung von '$($benutzer.Department)' auf 'Ausgeschieden' setzen" "Abteilung auf 'Ausgeschieden' gesetzt (dynamische Gruppen werden automatisch angepasst)" {
    Update-MgUser -UserId $id -Department "Ausgeschieden" -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# 5. Zugewiesene Gruppen entfernen, Rollen prüfen
# ---------------------------------------------------------------------------
$mitgliedschaften = Get-MgUserMemberOf -UserId $id -All -ErrorAction Stop

foreach ($m in $mitgliedschaften) {
    $typ  = $m.AdditionalProperties['@odata.type']
    $name = $m.AdditionalProperties['displayName']

    if ($typ -eq '#microsoft.graph.directoryRole') {
        Write-Warning "Admin-Rolle '$name' ist noch zugewiesen. Bitte manuell entziehen."
        Add-Protokoll "Admin-Rolle prüfen" "Noch zugewiesen: $name"
        continue
    }
    if ($typ -ne '#microsoft.graph.group') { continue }

    $gruppenTypen = $m.AdditionalProperties['groupTypes']
    if ($gruppenTypen -contains 'DynamicMembership') {
        Write-Host "[5] $name : dynamische Gruppe, Mitgliedschaft endet über die Abteilungsänderung"
        continue
    }
    if ($lizenzGruppen -contains $m.Id) {
        Write-Warning "$name vergibt eine Lizenz und wird NICHT entfernt. Erst Postfach umwandeln, dann manuell entfernen."
        Add-Protokoll "Gruppe entfernen" "Übersprungen (Lizenzgruppe): $name"
        continue
    }
    $gruppenId = $m.Id
    Invoke-Schritt "5" "Aus Gruppe '$name' entfernen" "Aus Gruppe '$name' entfernt" {
        Remove-MgGroupMemberByRef -GroupId $gruppenId -DirectoryObjectId $id -ErrorAction Stop
    }
}

# ---------------------------------------------------------------------------
# 6. Lizenzen: nur Hinweis
# ---------------------------------------------------------------------------
if ($benutzer.AssignedLicenses.Count -gt 0) {
    Write-Warning ("Der Benutzer hat noch {0} Lizenz(en) und damit vermutlich ein Postfach. " -f $benutzer.AssignedLicenses.Count +
        "Nächste Schritte: Postfach in ein freigegebenes Postfach umwandeln (Exchange Online PowerShell: " +
        "Set-Mailbox -Identity $UserPrincipalName -Type Shared), OneDrive-Daten übergeben, danach Lizenz entfernen.")
    Add-Protokoll "Lizenzen prüfen" "Noch $($benutzer.AssignedLicenses.Count) Lizenz(en), manuelle Nacharbeit nötig"
}
else {
    Write-Host "[6] Keine Lizenzen zugewiesen" -ForegroundColor Green
    Add-Protokoll "Lizenzen prüfen" "Keine Lizenzen"
}

# ---------------------------------------------------------------------------
# 7. Protokoll
# ---------------------------------------------------------------------------
if ($protokoll.Count -gt 0 -and -not $WhatIfPreference) {
    $csvOptionen = @{ Path = $ProtokollPfad; Delimiter = ';'; Append = $true }
    if ($PSVersionTable.PSVersion.Major -ge 6) { $csvOptionen.Encoding = 'utf8BOM' } else { $csvOptionen.Encoding = 'UTF8'; $csvOptionen.NoTypeInformation = $true }
    $protokoll | Export-Csv @csvOptionen
    Write-Host "`nProtokoll ergänzt: $((Resolve-Path $ProtokollPfad).Path)" -ForegroundColor Cyan
}
if ($script:fehlerAnzahl -gt 0) {
    Write-Warning "$($script:fehlerAnzahl) Schritt(e) sind fehlgeschlagen. Bitte die Meldungen oben prüfen und das Skript danach erneut ausführen."
}
Write-Host "Weitere manuelle Schritte: Geräte einsammeln und in Intune zurücksetzen, Zugänge im Passwortmanager entfernen.`n"
