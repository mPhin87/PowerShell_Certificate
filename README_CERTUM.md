# 🔐 CERTUM Cloud Code Signing für PowerShell

Dieses Repository enthält Tools zur Signierung von PowerShell-Skripten mit CERTUM Cloud Code Signing Zertifikaten.

## 📁 Dateien

| Datei | Beschreibung |
|-------|--------------|
| `PhinIT_CERTUM_Setup.ps1` | Setup und Diagnose-Tool für CERTUM Zertifikate |
| `PhinIT_CloudCodeSigning_CERTUM.ps1` | Kommandozeilen-Tool zur Skript-Signierung |
| `PhinIT_CodeSigning_GUI.ps1` | Grafische Benutzeroberfläche zur Skript-Signierung |
| `EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer` | Öffentliches Zertifikat (für Trusted Root) |

## 🚀 Quick Start

1. **CERTUM Zertifikat installieren**
   ```powershell
   # Setup-Tool ausführen
   .\PhinIT_CERTUM_Setup.ps1
   ```

2. **Skript signieren (GUI)**
   ```powershell
   # Grafische Benutzeroberfläche starten
   .\PhinIT_CodeSigning_GUI.ps1
   ```

3. **Skript signieren (Kommandozeile)**
   ```powershell
   # Einzelnes Skript signieren
   .\PhinIT_CloudCodeSigning_CERTUM.ps1 -ScriptPath "C:\MeinScript.ps1"
   
   # Mit spezifischem Zertifikat
   .\PhinIT_CloudCodeSigning_CERTUM.ps1 -ScriptPath "C:\MeinScript.ps1" -CertThumbprint "7352 8B48 7285 1395"
   ```

## 📋 Voraussetzungen

### CERTUM Cloud Code Signing Zertifikat
- Gültiges CERTUM Cloud Code Signing Zertifikat
- Zertifikat im Personal Store installiert (`Cert:\CurrentUser\My`)
- Optional: SimplySign Desktop Application

### PowerShell Konfiguration
```powershell
# Empfohlene Execution Policy für signierte Skripte
Set-ExecutionPolicy -ExecutionPolicy AllSigned -Scope CurrentUser
```

## 🔧 Installation

### 1. CERTUM Zertifikat Installation

1. **Zertifikat vom CERTUM Portal herunterladen**
   - Besuchen Sie: https://panel.certum.eu/
   - Laden Sie Ihr Zertifikat im PFX/P12 Format herunter

2. **Zertifikat installieren**
   - Doppelklicken Sie die PFX-Datei
   - Installieren Sie im **Persönlichen Speicher** (Personal Store)
   - Wählen Sie "Aktueller Benutzer" als Speicherort

3. **Installation überprüfen**
   ```powershell
   # Alle Code Signing Zertifikate anzeigen
   Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert
   ```

### 2. SimplySign Desktop (Optional)

Für erweiterte CERTUM Cloud Code Signing Funktionen:
- Download: https://www.certum.eu/certum/cert,offer_en_simplysign_desktop.xml

## 📖 Verwendung

### Setup und Diagnose

Das Setup-Tool führt umfassende Überprüfungen durch:

```powershell
.\PhinIT_CERTUM_Setup.ps1
```

**Überprüft:**
- ✅ Installierte Zertifikate
- ✅ SimplySign Desktop Installation  
- ✅ PowerShell Execution Policy
- ✅ Test-Signierung

### GUI-Tool

Einfache grafische Benutzeroberfläche:

```powershell
.\PhinIT_CodeSigning_GUI.ps1
```

**Features:**
- 📁 Datei-Browser zur Skript-Auswahl
- 📋 Automatische Zertifikat-Erkennung
- ⏰ Timestamp Server Option
- ✅ Status-Anzeige
- 📦 Batch-Signierung möglich

### Kommandozeilen-Tool

Für Automatisierung und Batch-Processing:

```powershell
# Einzelnes Skript
.\PhinIT_CloudCodeSigning_CERTUM.ps1 -ScriptPath "C:\Scripts\MeinScript.ps1"

# Spezifisches Zertifikat verwenden
.\PhinIT_CloudCodeSigning_CERTUM.ps1 -ScriptPath "C:\Scripts\MeinScript.ps1" -CertThumbprint "1234567890ABCDEF"

# Beispiel für Batch-Verarbeitung
Get-ChildItem -Path "C:\Scripts" -Filter "*.ps1" | ForEach-Object {
    .\PhinIT_CloudCodeSigning_CERTUM.ps1 -ScriptPath $_.FullName
}
```

## 🔍 Zertifikat-Informationen

### Ihr CERTUM Zertifikat (basierend auf den Screenshots)

- **Typ:** Code Signing (für non-qualified certificates)
- **Gültigkeitsdauer:** 31.07.2026  
- **Schlüssel:** RSA, 4096 bits
- **Kartennummer:** 7352 8B48 7285 1395

### Automatische Zertifikat-Erkennung

Die Tools erkennen CERTUM Zertifikate automatisch anhand:
- Enhanced Key Usage: "Code Signing"
- Issuer enthält "CERTUM"
- Subject enthält relevante Begriffe

## 🛠️ Troubleshooting

### Häufige Probleme

**1. "Kein Code Signing Zertifikat gefunden"**
```powershell
# Alle Zertifikate im Personal Store anzeigen
Get-ChildItem -Path Cert:\CurrentUser\My | Format-List Subject, Thumbprint, EnhancedKeyUsageList
```

**2. "Zertifikat hat keinen privaten Schlüssel"**
- Script als Administrator ausführen
- SimplySign Desktop App installieren
- Zertifikat neu installieren

**3. "Timestamp Server nicht erreichbar"**
- Alternativer Timestamp Server: `http://timestamp.digicert.com`
- Signierung ohne Timestamp: Option in GUI deaktivieren

**4. "Signatur ungültig"**
```powershell
# Signatur überprüfen
Get-AuthenticodeSignature -FilePath "C:\MeinScript.ps1"
```

### Debug-Informationen sammeln

```powershell
# System-Informationen
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Host "Execution Policy: $(Get-ExecutionPolicy)"
Write-Host "Benutzer: $env:USERNAME"

# Zertifikat-Details
Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.EnhancedKeyUsageList -match "Code Signing" } | 
    Select-Object Subject, Thumbprint, NotBefore, NotAfter, HasPrivateKey, @{Name="KeyLength"; Expression={$_.PublicKey.Key.KeySize}}
```

## 🔗 Nützliche Links

- **CERTUM Portal:** https://panel.certum.eu/
- **SimplySign Desktop:** https://www.certum.eu/certum/cert,offer_en_simplysign_desktop.xml
- **CERTUM Support:** https://support.certum.eu/
- **Microsoft Code Signing:** https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-authenticodesignature

## 📝 Hinweise

- **Timestamp Server:** Empfohlen für langfristige Gültigkeit der Signatur
- **Execution Policy:** `AllSigned` für maximale Sicherheit
- **Zertifikat-Speicherort:** Persönlicher Speicher (CurrentUser\My)
- **Hash-Algorithmus:** SHA256 (Standard und empfohlen)

## 🤝 Support

Bei Problemen:
1. Setup-Tool ausführen: `.\PhinIT_CERTUM_Setup.ps1`
2. Debug-Informationen sammeln (siehe oben)
3. CERTUM Support kontaktieren: https://support.certum.eu/

---

*Erstellt für PhinIT - PowerShell Code Signing mit CERTUM Cloud Zertifikaten*
