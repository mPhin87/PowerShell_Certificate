# 🚀 PhinIT CERTUM Tool → EXE Konvertierung
# =====================================================

# ANLEITUNG: PowerShell Script zu EXE konvertieren
# Für maximale Benutzerfreundlichkeit und Umgehung der PowerShell ExecutionPolicy

## 🎯 WARUM EXE?
- ✅ Umgeht PowerShell ExecutionPolicy komplett
- ✅ Windows SmartScreen funktioniert normal mit Code Signing
- ✅ Benutzer brauchen keine PowerShell-Kenntnisse
- ✅ Einfache Verteilung wie normale Software
- ✅ Keine "nicht vertrauenswürdig" Probleme bei signierten EXEs

## 📦 METHODE 1: PS2EXE (Empfohlen - Kostenlos)

### Installation:
```powershell
Install-Module PS2EXE -Scope CurrentUser
```

### Konvertierung:
```powershell
# Basis-Konvertierung
PS2EXE -inputFile "PhinIT_CodeSigning_CERTUM_V0.2.ps1" -outputFile "PhinIT_CodeSigning_CERTUM.exe"

# Erweiterte Optionen (empfohlen):
PS2EXE -inputFile "PhinIT_CodeSigning_CERTUM_V0.2.ps1" `
       -outputFile "PhinIT_CodeSigning_CERTUM.exe" `
       -iconFile "icon.ico" `
       -title "PhinIT CERTUM Code Signing Tool" `
       -description "Advanced Code Signing Tool for CERTUM Certificates" `
       -company "PhinIT" `
       -product "CERTUM Code Signing Tool" `
       -copyright "© 2025 PhinIT" `
       -version "0.2.0.0" `
       -requireAdmin `
       -supportOS `
       -longPaths
```

## 📦 METHODE 2: PowerShell Studio (Kommerziell)
- Professionellste Lösung
- Beste EXE-Qualität  
- GUI-Designer integriert
- Erweiterte Packaging-Optionen

## 📦 METHODE 3: Compile-PS1 (Kostenlos Alternative)
```powershell
# GitHub: https://github.com/MScholtes/PS2EXE
iex (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/MScholtes/PS2EXE/master/src/PS2EXE.ps1')
```

## 🔐 NACH DER KONVERTIERUNG: EXE SIGNIEREN

### Mit signtool.exe:
```cmd
signtool sign /sha1 "[CERTUM_THUMBPRINT]" /tr http://time.certum.pl /td sha256 /fd sha256 /v "PhinIT_CodeSigning_CERTUM.exe"
```

### Mit dem Tool selbst:
1. EXE erstellen
2. Mit der ursprünglichen PowerShell-Version die neue EXE signieren
3. Signierte EXE an Benutzer verteilen

## 📋 DEPLOYMENT-WORKFLOW

### Für Entwickler:
1. PowerShell Script entwickeln
2. Mit PS2EXE zu EXE konvertieren  
3. EXE mit CERTUM Zertifikat signieren
4. Optional: MSI-Installer erstellen
5. An Benutzer verteilen

### Für Endbenutzer:
1. Signierte EXE herunterladen
2. Windows SmartScreen erkennt Code Signing automatisch
3. EXE starten - keine PowerShell ExecutionPolicy Probleme
4. Tool installiert automatisch Trust für zukünftige PowerShell Scripts

## 🎯 VORTEILE DER EXE-LÖSUNG

**Technisch:**
- Keine PowerShell ExecutionPolicy Konflikte
- Normale Windows Code Signing Regeln
- SmartScreen Integration
- Bessere Antivirus-Kompatibilität

**Benutzerfreundlich:**
- Doppelklick zum Starten
- Keine PowerShell-Kenntnisse nötig
- Normale Software-Installation
- Professional wirkendes Deployment

**Entwickler:**
- Ein einmaliger Build-Prozess
- Einfachere Verteilung
- Weniger Support-Anfragen
- Professionellere Software-Distribution

## ⚡ NÄCHSTE SCHRITTE

1. **Sofort:** Batch-Starter verwenden (bereits erstellt)
2. **Kurz:** PS2EXE installieren und EXE erstellen
3. **Mittel:** MSI-Installer für professionelle Verteilung
4. **Lang:** Automatischen Build-Pipeline einrichten
