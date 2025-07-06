<#
.SYNOPSIS
  Importiert ein öffentliches Zertifikat (.cer) in den Store "Vertrauenswürdige Stammzertifizierungsstellen" (LocalMachine\Root), damit signierte Skripte ohne Warnungen ausgeführt werden können.
#>

# Laden der Windows Forms Assembly für MessageBox
Add-Type -AssemblyName System.Windows.Forms

# Pfad zur CER-Datei im gleichen Verzeichnis wie das Skript
$CerPath = Join-Path -Path $PSScriptRoot -ChildPath 'EA4D3E80D6712E4FD7F39B32B359BC48D36F8F94.cer'

# Überprüfen, ob die CER-Datei existiert
if (-not (Test-Path -Path $CerPath)) {
    [System.Windows.Forms.MessageBox]::Show("CER-Datei nicht gefunden: $CerPath", "Fehler")
    exit
}

# Thumbprint der CER ermitteln
try {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CerPath)
    $thumbprint = $cert.Thumbprint
} catch {
    [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden der CER-Datei: $_", "Fehler")
    exit
}

# Store-Location für Trusted Root CA
$storeRoot = 'Cert:\LocalMachine\Root'

# Prüfen, ob das Zertifikat bereits im Root-Store vorhanden ist
$existsRoot = Get-ChildItem -Path $storeRoot -Recurse | Where-Object Thumbprint -EQ $thumbprint

if ($existsRoot) {
    [System.Windows.Forms.MessageBox]::Show("Zertifikat bereits im Trusted Root Store vorhanden (Thumbprint: $thumbprint).", "Information")
} else {
    try {
        Import-Certificate -FilePath $CerPath -CertStoreLocation $storeRoot | Out-Null
        [System.Windows.Forms.MessageBox]::Show("Import abgeschlossen (Thumbprint: $thumbprint).", "Erfolg")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Importieren des Zertifikats: $_", "Fehler")
    }
}