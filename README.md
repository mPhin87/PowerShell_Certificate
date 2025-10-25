
# ✅ PhinIT Trusted Script Signing Certificate

Dieses Repository enthält das öffentliche **Code-Signing-Zertifikat** (self-signed) von **PHINIT / easyIT**, mit dem alle PowerShell-Skripte aus den zugehörigen Repositories signiert wurden. Damit Deine Systeme den signierten Skripten vertrauen, kannst Du das Zertifikat manuell oder automatisiert hinterlegen.

This repository provides the public **self-signed code-signing certificate** from **PHINIT / easyIT**, used to sign all PowerShell scripts across related repositories. To ensure your systems trust these signed scripts, you can import the certificate manually or via PowerShell.

## 📚 Inhalt / Table of Contents

- [🇩🇪 Anleitung (Deutsch)](#-anleitung-deutsch)
  - [📦 Was ist enthalten?](#-was-ist-enthalten)
  - [🛠️ Manuelle Einrichtung](#️-manuelle-einrichtung)
  - [⚙️ Automatisierter Import](#️-automatisierter-import)
- [🇬🇧 Instructions (English)](#-instructions-english)
  - [📦 Included Files](#-included-files)
  - [🛠️ Manual Setup](#️-manual-setup)
  - [⚙️ Automated Import](#️-automated-import)
- [🔐 Sicherheitshinweis / Security Notice](#-sicherheitshinweis--security-notice)

---
## 🇩🇪 Anleitung (Deutsch)

### 📦 Was ist enthalten?

- `EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer`  
  → Öffentliches Zertifikat zum Import in den Zertifikatspeicher

- `PhinIT_TrustetScripts.ps1`  
  → PowerShell-Skript zum automatisierten Hinterlegen des Zertifikats  
  **Hinweis:** Skript **immer mit Admin-Rechten** ausführen!

---

### 🛠️ Manuelle Einrichtung

1. Lade `EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer` herunter.
2. Starte `mmc.exe` → Snap-In *Zertifikate (Lokaler Computer)* hinzufügen.
3. Importiere das Zertifikat in **beide Speicherorte**:
   - **Vertrauenswürdige Herausgeber** → Zertifikate
   - **Vertrauenswürdige Stammzertifizierungsstellen** → Zertifikate

---

### ⚙️ Automatisierter Import

```powershell
# Als Administrator ausführen:
.\PhinIT_TrustetScripts.ps1
```

Das Skript prüft die erforderlichen Speicherorte und importiert das Zertifikat für den lokalen Computerkontext.

---
## 🇬🇧 Instructions (English)

### 📦 Included Files

- `EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer`  
  → Public certificate for importing into the certificate store

- `PhinIT_TrustetScripts.ps1`  
  → PowerShell script to automate certificate import  
  **Note:** Always run this script **with administrative privileges**!

---

### 🛠️ Manual Setup

1. Download `EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer`.
2. Start `mmc.exe` → Add *Certificates* snap-in for *Local Computer*.
3. Import the certificate into both stores:
   - **Trusted Publishers** → Certificates
   - **Trusted Root Certification Authorities** → Certificates

---

### ⚙️ Automated Import

```powershell
# Run as Administrator:
.\PhinIT_TrustetScripts.ps1
```

The script automatically imports the certificate to the correct stores under the local machine context.

---

## 🔐 Sicherheitshinweis / Security Notice

Dieses Zertifikat ist **selbstsigniert** und ausschließlich für Skripte aus dem [PHINIT easyIT Projekten](https://github.com/PS-easyIT) vorgesehen.  
This certificate is **self-signed** and intended solely for scripts from the [PHINIT easyIT projects](https://github.com/PS-easyIT).  
