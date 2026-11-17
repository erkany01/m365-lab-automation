# M365-Lab: Automatisierung von On- und Offboarding

PowerShell-Skripte für eine Microsoft-365-Testumgebung (fiktive Nordlicht GmbH).
Sie nutzen das Microsoft Graph PowerShell SDK.

## Skripte

| Datei | Zweck |
|---|---|
| `Onboarding.ps1` | Legt neue Mitarbeitende aus einer CSV-Datei an (Namenskonvention, Abteilung, Position, Nutzungsstandort). Gruppen entstehen automatisch über dynamische Gruppen. |
| `Offboarding.ps1` | Sperrt das Konto, widerruft Sitzungen, setzt das Kennwort zurück, beendet Gruppenmitgliedschaften und protokolliert jeden Schritt mit seinem tatsächlichen Ergebnis. Lizenzen werden bewusst nicht automatisch entfernt. |
| `neue-mitarbeitende.csv` | Beispieldaten (Trennzeichen Semikolon). |

## Voraussetzungen

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Groups -Scope CurrentUser
Connect-MgGraph -TenantId "<tenant>.onmicrosoft.com" -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.Read.All","User-PasswordProfile.ReadWrite.All"
```

## Verwendung

```powershell
.\Onboarding.ps1 -CsvPfad .\neue-mitarbeitende.csv -WhatIf   # Testlauf
.\Onboarding.ps1 -CsvPfad .\neue-mitarbeitende.csv
.\Offboarding.ps1 -UserPrincipalName jonas.richter@<tenant>.onmicrosoft.com -WhatIf
.\Offboarding.ps1 -UserPrincipalName jonas.richter@<tenant>.onmicrosoft.com
```

## Designentscheidungen

- **Testlauf mit `-WhatIf`:** Beide Skripte zeigen vor der Ausführung, was passieren würde.
- **Mehrfach ausführbar:** Bereits vorhandene Benutzer werden übersprungen.
- **Gruppen über Attribute:** Die Abteilung steuert die Mitgliedschaft in dynamischen Gruppen.
- **Kein Datenverlust beim Offboarding:** Lizenzgruppen und Lizenzen werden nicht automatisch entfernt, damit das Postfach vorher umgewandelt werden kann.

Hinweis: Testumgebung, alle Personen sind fiktiv.
