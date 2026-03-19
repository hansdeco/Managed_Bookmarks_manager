# Changelog — Decoster.tech.output.psm1

## 2026-03-15 (2)

### Test-DecosterOutputConfig — nieuw

- Nieuwe functie `Test-DecosterOutputConfig` voor config-validatie na `Import-PowerShellDataFile`.
- Parameters: `$Config` (hashtable, verplicht), `$RequiredKeys` (string-array van `"Section.Key"`-paden, verplicht).
- Controleert voor elke vereiste sleutel of de sectie en de sleutel bestaan én niet leeg zijn (`IsNullOrWhiteSpace`).
- Retourneert `[string[]]` met de ontbrekende of lege sleutels; lege array = config is volledig.
- Bedoeld om vanuit scripts aan te roepen vlak na de module-load, zodat de GUI een leesbare waarschuwing kan tonen i.p.v. stil te crashen bij een ontbrekende instelling.
- Geëxporteerd via `Export-ModuleMember`.

---

## 2026-03-15

### Format-DecosterOutputStep — nieuw
- Nieuwe functie `Format-DecosterOutputStep` voor genummerde of ongenummerde stapindicatoren bij meerstappen-operaties.
- Parameters: `$Message` (verplicht), `$Step` (optioneel), `$Total` (optioneel).
- Met `$Step` en `$Total`: prefix `[Step/Total]` — voorbeeld: `[1/3] Bestand laden`.
- Zonder stap-parameters: prefix `[>>]` — voorbeeld: `[>>] Validating input`.
- Geëxporteerd via `Export-ModuleMember`.

### Format-DecosterOutputSubItem — nieuw
- Nieuwe functie `Format-DecosterOutputSubItem` voor ingesprongen detail-regels visueel genest onder een voorgaande log-lijn.
- Gebruik: als extra context-regels onder een Ok / Warn / Fail entry, zonder de statusprefixes te herhalen.
- Prefix: `└ ` met overeenkomende inspringing als `Format-DecosterOutputFail` / `Format-DecosterOutputWarn`.
- Geëxporteerd via `Export-ModuleMember`.

### Export-ModuleMember bijgewerkt
- `Format-DecosterOutputStep` en `Format-DecosterOutputSubItem` toegevoegd aan de exportlijst.

---

## Initiële staat 

### Bestaande functies (ongewijzigd)
| Functie | Beschrijving |
|---|---|
| `Format-DecosterOutputHeader` | Sectie-header `=== Naam ===` |
| `Format-DecosterOutputOk` | Succes-lijn `[OK]` |
| `Format-DecosterOutputFail` | Fout-lijn `[FAIL]` met optionele detail |
| `Format-DecosterOutputSkip` | Overgeslagen item `[SKIP]` |
| `Format-DecosterOutputWarn` | Waarschuwing `[WARN]` met optioneel Detail en ErrorMessage |
| `Format-DecosterOutputWarnResult` | Vervolglijn na een `[WARN]` → `[OK]` of `[FAIL]` |
| `Format-DecosterOutputError` | Harde fout `[ERROR]` |
| `Format-DecosterOutputInfo` | Informatieve insprong-lijn (geen prefix) |
| `Format-DecosterOutputConnected` | Standaard verbinding-succes lijn |
| `Format-DecosterOutputConnectionFail` | Verbinding-fout lijn |
| `Format-DecosterOutputSummary` | Per-fase samenvatting: `X OK | Y skipped | Z failed` |
| `Format-DecosterOutputTotal` | Eindtotaal over alle fasen |
| `Format-DecosterOutputSeparator` | Dunne scheidingslijn `---` |
| `Format-DecosterOutputDivider` | Zware scheidingslijn `XXX...` |
| `Format-DecosterOutputWarningBlock` | Blok-waarschuwing met zware dividers |
| `Format-DecosterOutputFailureBlock` | Blok-fout met zware dividers |
| `Format-DecosterOutputKeyValue` | Sleutel : waarde insprong-lijn |
| `Add-DecosterOutputOutput` | Voegt `[string[]]` toe aan een ListBox |
| `Set-DecosterOutputProgress` | Stelt ProgressBar waarde en Label tekst in |
| `Set-DecosterOutputCredentialStatus` | Kleurt credential-indicator Panel + Label |
| `Set-DecosterOutputLabelStatus` | Stelt Label tekst en kleur in op basis van state |
| `Remove-DecosterOutputOldLogs` | Verwijdert logbestanden ouder dan retentieperiode |
| `Write-DecosterOutputLog` | Schrijft ListBox-inhoud naar timestamped logbestand |
