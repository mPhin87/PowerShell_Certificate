<#
.SYNOPSIS
    PhinIT CERTUM Code Signing Tool V0.2 - Trust Manager Edition
    Advanced GUI-Tool zum Signieren von PowerShell-Skripten mit CERTUM Zertifikaten
    und zur Verwaltung von Trust-Beziehungen fuer vertrauenswuerdige Publisher

.DESCRIPTION
    Erweiterte Version des PhinIT Code Signing Tools mit folgenden Features:
    - Integrierter PowerShell Datei-Explorer
    - Signierung mit CERTUM Cloud Code Signing Zertifikaten
    - Automatische Trust-Installation fur vertrauenswurdige Publisher
    - Loesung fur "nicht vertrauenswuerdig" Problem bei signierten Skripten
    - Administrator-Privilegien-Erkennung
    - Timestamp-Server Support
    - Signatur-Status-Anzeige im Explorer

    Das Tool loest das haeufige Problem, dass signierte PowerShell-Skripte vom System
    als "nicht vertrauenswuerdig" eingestuft werden, obwohl die Root CA installiert ist.
    Die Loesung liegt in der Installation des Code Signing Zertifikats als
    vertrauenswuerdiger Publisher im lokalen Computer Store.

.PARAMETER None
    Dieses Skript wird als GUI-Anwendung ohne Parameter gestartet.

.NOTES
    Dateiname     : PhinIT_CodeSigning_CERTUM_V0.2.ps1
    Autor         : Andreas Hepp / PhinIT
    Version       : 0.2
    Erstellt      : 2025
    Abhaengigkeiten: CERTUM Code Signing Zertifikat, Windows Forms

    CERTUM Cloud Code Signing Voraussetzungen:
    - CERTUM Code Signing Zertifikat installiert (Cert:\CurrentUser\My)
    - SimplySign Desktop App (optional fuer Cloud-Zertifikate)
    - Fuer Trust-Installation: Administrator-Rechte

.EXAMPLE
    .\PhinIT_CodeSigning_CERTUM_V0.2.ps1
    Startet die GUI-Anwendung mit Datei-Explorer und Signierungsfunktionen

.LINK
    https://github.com/PhinIT
#>

#Requires -Version 5.1
[CmdletBinding()]
param()

# =============================================================================
# ASSEMBLY LADEN FUER WINDOWS FORMS
# =============================================================================

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Write-Host "Windows Forms Assemblies erfolgreich geladen" -ForegroundColor Green
}
catch {
    Write-Error "Fehler beim Laden der Windows Forms Assemblies: $($_.Exception.Message)"
    Write-Error "Bitte stellen Sie sicher, dass .NET Framework installiert ist."
    exit 1
}

# =============================================================================
# REGISTRY EINSTELLUNGEN
# =============================================================================

# Registry-Pfad fuer Einstellungen
$registryPath = "HKCU:\Software\easyIT\PSS2ES"

# Standard-Einstellungen
$defaultSettings = @{
    PS2EXEPath = Join-Path $PSScriptRoot "ps2exe"
    DefaultFolder = [Environment]::GetFolderPath("MyDocuments")
    IconPath = ""
    AppAuthor = "PhinIT"
    AppCompany = "PhinIT"
    AppProduct = "PowerShell Tool"
    AppCopyright = "(c) 2025 PhinIT"
    AppVersion = "1.0.0.0"
    RequireAdmin = $false
    NoConsole = $true
    CPUArch = "AnyCPU"
    TimestampServer = "http://time.certum.pl"
}

# Aktuelle Einstellungen (werden aus Registry geladen)
$script:settings = $defaultSettings.Clone()

# Globale Variablen mit sicheren Standardwerten initialisieren
$script:selectedScriptPath = ""
$script:selectedCertificate = $null
$script:currentDirectory = [Environment]::GetFolderPath("MyDocuments")

# Debug-Modus (nur für Entwicklung)
$script:debugMode = $false

# Sicherstellen, dass settings Hashtable existiert
if (-not $script:settings) {
    $script:settings = $defaultSettings.Clone()
}

# Hilfsfunktion für Debug-Ausgaben
function Write-DebugInfo {
    param([string]$Message)
    if ($script:debugMode) {
        Write-Host $Message -ForegroundColor Cyan
    }
}

# Registry-Funktionen
function Save-SettingsToRegistry {
    param([hashtable]$Settings)
    
    try {
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        
        foreach ($key in $Settings.Keys) {
            Set-ItemProperty -Path $registryPath -Name $key -Value $Settings[$key] -Type String -Force
        }
        
        Update-Info "Einstellungen erfolgreich in Registry gespeichert"
        return $true
    }
    catch {
        Update-Info "Fehler beim Speichern der Einstellungen: $($_.Exception.Message)"
        return $false
    }
}

function Load-SettingsFromRegistry {
    try {
        if (Test-Path $registryPath) {
            $registrySettings = Get-ItemProperty -Path $registryPath
            
            foreach ($key in $defaultSettings.Keys) {
                if ($registrySettings.PSObject.Properties.Name -contains $key) {
                    # Konvertierung fuer Boolean-Werte
                    if ($key -in @("RequireAdmin", "NoConsole")) {
                        $script:settings[$key] = [bool]::Parse($registrySettings.$key)
                    } elseif ($key -eq "CPUArch") {
                        # CPUArch Validierung
                        $validArchs = @("AnyCPU", "x86", "x64")
                        if ($registrySettings.$key -in $validArchs) {
                            $script:settings[$key] = $registrySettings.$key
                        } else {
                            $script:settings[$key] = $defaultSettings[$key]
                        }
                    } elseif ($key -eq "DefaultFolder") {
                        # DefaultFolder Validierung
                        $folderValue = $registrySettings.$key
                        if ($folderValue -and $folderValue.Trim() -ne "" -and (Test-Path $folderValue)) {
                            $script:settings[$key] = $folderValue
                        } else {
                            # Ungueltiger Pfad - verwende Standardwert
                            $script:settings[$key] = $defaultSettings[$key]
                            Update-Info "Ungueltiger DefaultFolder in Registry gefunden, verwende Standardwert: $($defaultSettings[$key])"
                        }
                    } else {
                        $script:settings[$key] = $registrySettings.$key
                    }
                }
            }
            
            Update-Info "Einstellungen aus Registry geladen"
            return $true
        } else {
            Update-Info "Keine gespeicherten Einstellungen gefunden, verwende Standardwerte"
            # Sicherstellen, dass alle erforderlichen Settings gesetzt sind
            $script:settings = $defaultSettings.Clone()
            return $false
        }
    }
    catch {
        Update-Info "Fehler beim Laden der Einstellungen: $($_.Exception.Message)"
        # Bei Fehlern Standardwerte verwenden
        $script:settings = $defaultSettings.Clone()
        return $false
    }
}

function Reset-SettingsToDefault {
    $script:settings = $defaultSettings.Clone()
    Update-Info "Einstellungen auf Standardwerte zurueckgesetzt"
}

# Administrator-Privilegien pruefen
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Self-Elevation wenn nicht als Administrator gestartet
if (-not $isAdmin -and $args -notcontains "-NoElevate") {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "PhinIT CERTUM Code Signing Tool benoetigt Administrator-Rechte fuer die Trust-Verwaltung.`n`nMoechten Sie das Tool als Administrator neu starten?`n`nEmpfohlen: Ermoeglicht vollstaendige Trust-Installation`nOhne Admin: Nur Signierung moeglich",
        "Administrator-Rechte erforderlich",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq "Yes") {
        try {
            Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -NoElevate" -Verb RunAs -Wait
            exit 0
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Starten als Administrator:`n`n$($_.Exception.Message)",
                "Elevation Fehler",
                "OK",
                "Error"
            )
        }
    }
}

# =============================================================================
# HAUPTFENSTER ERSTELLEN
# =============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "PhinIT CERTUM Code Signing & EXE Creator V0.2"
$form.Size = New-Object System.Drawing.Size(1600, 925)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(1200, 750)
$form.Icon = [System.Drawing.SystemIcons]::Shield
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Padding = New-Object System.Windows.Forms.Padding(0)
$form.FormBorderStyle = "Sizable"

# Header-Bereich
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Size = New-Object System.Drawing.Size(1584, 70)
$headerPanel.Location = New-Object System.Drawing.Point(8, 8)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$headerPanel.BorderStyle = "None"
$form.Controls.Add($headerPanel)

# Titel
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "CERTUM Code Signing and EXE Creator"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16, [System.Drawing.FontStyle]::Regular)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$titleLabel.Location = New-Object System.Drawing.Point(24, 18)
$titleLabel.Size = New-Object System.Drawing.Size(600, 35)
$headerPanel.Controls.Add($titleLabel)


# Options Button im Header
$headerOptionsButton = New-Object System.Windows.Forms.Button
$headerOptionsButton.Text = "Einstellungen"
$headerOptionsButton.Location = New-Object System.Drawing.Point(1450, 20)
$headerOptionsButton.Size = New-Object System.Drawing.Size(100, 35)
$headerOptionsButton.FlatStyle = "Flat"
$headerOptionsButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$headerOptionsButton.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
$headerOptionsButton.ForeColor = [System.Drawing.Color]::White
$headerOptionsButton.FlatAppearance.BorderSize = 0
$headerOptionsButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(79, 150, 246)
$headerOptionsButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(39, 110, 226)
$headerOptionsButton.Add_Click({
    Show-OptionsDialog
})
$headerPanel.Controls.Add($headerOptionsButton)

# Haupt-Container mit Splitter
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Size = New-Object System.Drawing.Size(1584, 750)
$splitContainer.Location = New-Object System.Drawing.Point(8, 86)
$splitContainer.Anchor = "Top,Bottom,Left,Right"
$splitContainer.SplitterDistance = 700
$splitContainer.BorderStyle = "None"
$splitContainer.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$splitContainer.SplitterWidth = 8
$form.Controls.Add($splitContainer)

# =============================================================================
# LINKER PANEL: DATEI-EXPLORER
# =============================================================================

# Explorer-Header
$explorerHeader = New-Object System.Windows.Forms.Label
$explorerHeader.Text = "PowerShell Datei-Explorer"
$explorerHeader.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12, [System.Drawing.FontStyle]::Regular)
$explorerHeader.Location = New-Object System.Drawing.Point(16, 16)
$explorerHeader.Size = New-Object System.Drawing.Size(300, 24)
$explorerHeader.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$splitContainer.Panel1.Controls.Add($explorerHeader)

# Verzeichnis-Navigation
$directoryPanel = New-Object System.Windows.Forms.Panel
$directoryPanel.Location = New-Object System.Drawing.Point(16, 48)
$directoryPanel.Size = New-Object System.Drawing.Size(700, 40)
$directoryPanel.BorderStyle = "FixedSingle"
$directoryPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$directoryPanel.BorderStyle = "None"
$directoryPanel.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$splitContainer.Panel1.Controls.Add($directoryPanel)

$currentPathLabel = New-Object System.Windows.Forms.Label
$currentPathLabel.Text = $script:currentDirectory
$currentPathLabel.Location = New-Object System.Drawing.Point(12, 10)
$currentPathLabel.Size = New-Object System.Drawing.Size(400, 20)
$currentPathLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$currentPathLabel.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$directoryPanel.Controls.Add($currentPathLabel)

$upButton = New-Object System.Windows.Forms.Button
$upButton.Text = "UP"
$upButton.Location = New-Object System.Drawing.Point(500, 6)
$upButton.Size = New-Object System.Drawing.Size(32, 28)
$upButton.FlatStyle = "Flat"
$upButton.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$upButton.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$upButton.FlatAppearance.BorderSize = 0
$upButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$directoryPanel.Controls.Add($upButton)

$homeButton = New-Object System.Windows.Forms.Button
$homeButton.Text = "HOME"
$homeButton.Location = New-Object System.Drawing.Point(535, 6)
$homeButton.Size = New-Object System.Drawing.Size(55, 28)
$homeButton.FlatStyle = "Flat"
$homeButton.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$homeButton.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$homeButton.FlatAppearance.BorderSize = 0
$homeButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$directoryPanel.Controls.Add($homeButton)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "BROWSE"
$browseButton.Location = New-Object System.Drawing.Point(610, 6)
$browseButton.Size = New-Object System.Drawing.Size(70, 28)
$browseButton.FlatStyle = "Flat"
$browseButton.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
$browseButton.ForeColor = [System.Drawing.Color]::White
$browseButton.FlatAppearance.BorderSize = 0
$browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$directoryPanel.Controls.Add($browseButton)

# Datei-Liste
$fileListView = New-Object System.Windows.Forms.ListView
$fileListView.Location = New-Object System.Drawing.Point(16, 100)
$fileListView.Size = New-Object System.Drawing.Size(695, 650)
$fileListView.Anchor = "Top,Bottom,Left,Right"
$fileListView.View = "Details"
$fileListView.FullRowSelect = $true
$fileListView.GridLines = $false
$fileListView.MultiSelect = $false
$fileListView.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$fileListView.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$fileListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$fileListView.BorderStyle = "None"

# ListView Spalten
$fileListView.Columns.Add("Name", 407) | Out-Null
$fileListView.Columns.Add("Typ", 75) | Out-Null
$fileListView.Columns.Add("Signatur", 125) | Out-Null
$fileListView.Columns.Add("Groesse", 75) | Out-Null

$splitContainer.Panel1.Controls.Add($fileListView)

# =============================================================================
# RECHTER PANEL: SIGNIERUNG UND EXE-ERSTELLUNG
# =============================================================================

# Signierung-Header
$signingHeader = New-Object System.Windows.Forms.Label
$signingHeader.Text = ">> Code Signierung | EXE-Erstellung"
$signingHeader.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Regular)
$signingHeader.Location = New-Object System.Drawing.Point(16, 16)
$signingHeader.Size = New-Object System.Drawing.Size(400, 28)
$signingHeader.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$splitContainer.Panel2.Controls.Add($signingHeader)

# Ausgewaehlte Datei
$selectedFileLabel = New-Object System.Windows.Forms.Label
$selectedFileLabel.Text = "Ausgewaehlte Datei:"
$selectedFileLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11, [System.Drawing.FontStyle]::Regular)
$selectedFileLabel.Location = New-Object System.Drawing.Point(16, 56)
$selectedFileLabel.Size = New-Object System.Drawing.Size(150, 24)
$splitContainer.Panel2.Controls.Add($selectedFileLabel)

$selectedFileDisplay = New-Object System.Windows.Forms.Label
$selectedFileDisplay.Text = "Keine Datei ausgewaehlt"
$selectedFileDisplay.Location = New-Object System.Drawing.Point(16, 84)
$selectedFileDisplay.Size = New-Object System.Drawing.Size(750, 24)
$selectedFileDisplay.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$selectedFileDisplay.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$selectedFileDisplay.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$selectedFileDisplay.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$splitContainer.Panel2.Controls.Add($selectedFileDisplay)

# Trennlinie
$separator1 = New-Object System.Windows.Forms.Label
$separator1.BorderStyle = "None"
$separator1.Height = 1
$separator1.Location = New-Object System.Drawing.Point(16, 120)
$separator1.Size = New-Object System.Drawing.Size(750, 1)
$separator1.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$splitContainer.Panel2.Controls.Add($separator1)

# Zertifikat-Auswahl
$certSectionLabel = New-Object System.Windows.Forms.Label
$certSectionLabel.Text = "CERTUM Code Signing Zertifikat"
$certSectionLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12, [System.Drawing.FontStyle]::Regular)
$certSectionLabel.Location = New-Object System.Drawing.Point(16, 136)
$certSectionLabel.Size = New-Object System.Drawing.Size(300, 24)
$certSectionLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$splitContainer.Panel2.Controls.Add($certSectionLabel)

$certComboBox = New-Object System.Windows.Forms.ComboBox
$certComboBox.Location = New-Object System.Drawing.Point(16, 168)
$certComboBox.Size = New-Object System.Drawing.Size(650, 32)
$certComboBox.DropDownStyle = "DropDownList"
$certComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$certComboBox.BackColor = [System.Drawing.Color]::White
$certComboBox.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$certComboBox.FlatStyle = "Flat"
$splitContainer.Panel2.Controls.Add($certComboBox)

$refreshCertButton = New-Object System.Windows.Forms.Button
$refreshCertButton.Text = "Aktualisieren"
$refreshCertButton.Location = New-Object System.Drawing.Point(672, 168)
$refreshCertButton.Size = New-Object System.Drawing.Size(100, 32)
$refreshCertButton.FlatStyle = "Flat"
$refreshCertButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$refreshCertButton.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
$refreshCertButton.ForeColor = [System.Drawing.Color]::White
$refreshCertButton.FlatAppearance.BorderSize = 0
$splitContainer.Panel2.Controls.Add($refreshCertButton)

# Signierungsoptionen
$timestampCheckBox = New-Object System.Windows.Forms.CheckBox
$timestampCheckBox.Text = "Timestamp Server verwenden"
$timestampCheckBox.Location = New-Object System.Drawing.Point(16, 216)
$timestampCheckBox.Size = New-Object System.Drawing.Size(200, 24)
$timestampCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$timestampCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$timestampCheckBox.Checked = $true
$splitContainer.Panel2.Controls.Add($timestampCheckBox)

# Trennlinie 2
$separator2 = New-Object System.Windows.Forms.Label
$separator2.BorderStyle = "None"
$separator2.Height = 1
$separator2.Location = New-Object System.Drawing.Point(16, 252)
$separator2.Size = New-Object System.Drawing.Size(750, 1)
$separator2.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$splitContainer.Panel2.Controls.Add($separator2)

# Action Button Panel
$actionButtonLabel = New-Object System.Windows.Forms.Label
$actionButtonLabel.Text = ">> Aktionen"
$actionButtonLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Regular)
$actionButtonLabel.Location = New-Object System.Drawing.Point(16, 268)
$actionButtonLabel.Size = New-Object System.Drawing.Size(150, 28)
$actionButtonLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$splitContainer.Panel2.Controls.Add($actionButtonLabel)

# PS1 SIGNIEREN Button
$signPS1Button = New-Object System.Windows.Forms.Button
$signPS1Button.Text = "PS1 SIGNIEREN"
$signPS1Button.Location = New-Object System.Drawing.Point(15, 295)
$signPS1Button.Size = New-Object System.Drawing.Size(200, 50)
$signPS1Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$signPS1Button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$signPS1Button.ForeColor = [System.Drawing.Color]::White
$signPS1Button.FlatStyle = "Flat"
$signPS1Button.FlatAppearance.BorderSize = 0
$signPS1Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(90, 150, 200)
$signPS1Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(50, 110, 160)
$signPS1Button.Enabled = $false
$splitContainer.Panel2.Controls.Add($signPS1Button)

# PS1 ? EXE Button
$convertToEXEButton = New-Object System.Windows.Forms.Button
$convertToEXEButton.Text = "PS1 zu EXE"
$convertToEXEButton.Location = New-Object System.Drawing.Point(230, 295)
$convertToEXEButton.Size = New-Object System.Drawing.Size(200, 50)
$convertToEXEButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$convertToEXEButton.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$convertToEXEButton.ForeColor = [System.Drawing.Color]::White
$convertToEXEButton.FlatStyle = "Flat"
$convertToEXEButton.FlatAppearance.BorderSize = 0
$convertToEXEButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 160, 20)
$convertToEXEButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(235, 120, 0)
$convertToEXEButton.Enabled = $false
$splitContainer.Panel2.Controls.Add($convertToEXEButton)

# EXE SIGNIEREN Button
$signEXEButton = New-Object System.Windows.Forms.Button
$signEXEButton.Text = "EXE SIGNIEREN"
$signEXEButton.Location = New-Object System.Drawing.Point(445, 295)
$signEXEButton.Size = New-Object System.Drawing.Size(200, 50)
$signEXEButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$signEXEButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$signEXEButton.ForeColor = [System.Drawing.Color]::White
$signEXEButton.FlatStyle = "Flat"
$signEXEButton.FlatAppearance.BorderSize = 0
$signEXEButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(54, 159, 54)
$signEXEButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(14, 119, 14)
$signEXEButton.Enabled = $false
$splitContainer.Panel2.Controls.Add($signEXEButton)

# SimplySign Button
$simplySignButton = New-Object System.Windows.Forms.Button
$simplySignButton.Text = "SimplySign"
$simplySignButton.Location = New-Object System.Drawing.Point(660, 295)
$simplySignButton.Size = New-Object System.Drawing.Size(180, 50)
$simplySignButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$simplySignButton.BackColor = [System.Drawing.Color]::FromArgb(138, 43, 226)
$simplySignButton.ForeColor = [System.Drawing.Color]::White
$simplySignButton.FlatStyle = "Flat"
$simplySignButton.FlatAppearance.BorderSize = 0
$simplySignButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(158, 63, 246)
$simplySignButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(118, 23, 206)
$splitContainer.Panel2.Controls.Add($simplySignButton)

# Trennlinie 3
$separator3 = New-Object System.Windows.Forms.Label
$separator3.BorderStyle = "Fixed3D"
$separator3.Height = 2
$separator3.Location = New-Object System.Drawing.Point(15, 365)
$separator3.Size = New-Object System.Drawing.Size(850, 2)
$splitContainer.Panel2.Controls.Add($separator3)

# Info-Bereich
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Text = "Status & Information"
$infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$infoLabel.Location = New-Object System.Drawing.Point(15, 380)
$infoLabel.Size = New-Object System.Drawing.Size(200, 25)
$infoLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$splitContainer.Panel2.Controls.Add($infoLabel)

$infoTextBox = New-Object System.Windows.Forms.TextBox
$infoTextBox.Location = New-Object System.Drawing.Point(15, 410)
$infoTextBox.Size = New-Object System.Drawing.Size(850, 340)
$infoTextBox.Multiline = $true
$infoTextBox.WordWrap = $true
$infoTextBox.ScrollBars = "Vertical"
$infoTextBox.ReadOnly = $true
$infoTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$infoTextBox.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
$infoTextBox.Text = "PhinIT CERTUM Code Signing & EXE Creation Tool

EXE-Konvertierung - Losung fur PowerShell ExecutionPolicy:
================================================================================

EXE-Dateien umgehen PowerShell ExecutionPolicy komplett
Windows behandelt signierte EXE-Dateien normal
Keine 'Trust Publisher' Installation erforderlich
PS2EXE Integration fur nahtlose Konvertierung

Empfohlener Workflow:
1. PS1-Datei auswahlen (Doppelklick im Explorer)
2. 'PS1 -> EXE' klicken (konvertiert zu ausfuhrbarer EXE)
3. EXE-Datei auswahlen (Doppelklick im Explorer)
4. 'EXE SIGNIEREN' klicken (digital signieren)

Navigation: Doppelklick auf Ordner/Datei zur Auswahl
SimplySign fur CERTUM Cloud-Integration verfugbar
Zertifikate werden automatisch beim Start geladen

Bereit fur EXE-Workflow!"
$infoTextBox.Anchor = "Top,Bottom,Left,Right"
$splitContainer.Panel2.Controls.Add($infoTextBox)

# Status-Panel (am unteren Rand)
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Size = New-Object System.Drawing.Size(1584, 40)
$statusPanel.Location = New-Object System.Drawing.Point(8, 952)
$statusPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$statusPanel.BorderStyle = "FixedSingle"
$statusPanel.Anchor = "Bottom,Left,Right"
$form.Controls.Add($statusPanel)

# Status-Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Bereit - Waehlen Sie eine PowerShell-Datei aus dem Explorer..."
$statusLabel.Location = New-Object System.Drawing.Point(15, 10)
$statusLabel.Size = New-Object System.Drawing.Size(600, 20)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$statusLabel.Anchor = "Bottom,Left"
$statusPanel.Controls.Add($statusLabel)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(650, 8)
$progressBar.Size = New-Object System.Drawing.Size(120, 24)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$progressBar.Anchor = "Bottom,Right"
$statusPanel.Controls.Add($progressBar)

# =============================================================================
# OPTIONS FENSTER
# =============================================================================

function Show-OptionsDialog {
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.Text = "Einstellungen - PhinIT PS2EXE Tool"
    $optionsForm.Size = New-Object System.Drawing.Size(600, 700)
    $optionsForm.StartPosition = "CenterParent"
    $optionsForm.FormBorderStyle = "FixedDialog"
    $optionsForm.MaximizeBox = $false
    $optionsForm.MinimizeBox = $false
    $optionsForm.Icon = [System.Drawing.SystemIcons]::Settings

    # TabControl fuer verschiedene Kategorien
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(565, 600)
    $optionsForm.Controls.Add($tabControl)

    # Tab 1: Pfade
    $tabPaths = New-Object System.Windows.Forms.TabPage
    $tabPaths.Text = "Pfade"
    $tabControl.Controls.Add($tabPaths)

    # PS2EXE Pfad
    $ps2exeLabel = New-Object System.Windows.Forms.Label
    $ps2exeLabel.Text = "PS2EXE Ordner:"
    $ps2exeLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ps2exeLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabPaths.Controls.Add($ps2exeLabel)

    $ps2exeTextBox = New-Object System.Windows.Forms.TextBox
    $ps2exeTextBox.Text = $script:settings.PS2EXEPath
    $ps2exeTextBox.Location = New-Object System.Drawing.Point(20, 45)
    $ps2exeTextBox.Size = New-Object System.Drawing.Size(400, 20)
    $tabPaths.Controls.Add($ps2exeTextBox)

    $ps2exeBrowseButton = New-Object System.Windows.Forms.Button
    $ps2exeBrowseButton.Text = "..."
    $ps2exeBrowseButton.Location = New-Object System.Drawing.Point(430, 43)
    $ps2exeBrowseButton.Size = New-Object System.Drawing.Size(30, 26)
    $ps2exeBrowseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "PS2EXE Ordner auswaehlen"
        $folderBrowser.SelectedPath = $ps2exeTextBox.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $ps2exeTextBox.Text = $folderBrowser.SelectedPath
        }
    })
    $tabPaths.Controls.Add($ps2exeBrowseButton)

    # Standard-Ordner
    $defaultFolderLabel = New-Object System.Windows.Forms.Label
    $defaultFolderLabel.Text = "Standard-Ordner beim Start:"
    $defaultFolderLabel.Location = New-Object System.Drawing.Point(20, 80)
    $defaultFolderLabel.Size = New-Object System.Drawing.Size(150, 20)
    $tabPaths.Controls.Add($defaultFolderLabel)

    $defaultFolderTextBox = New-Object System.Windows.Forms.TextBox
    $defaultFolderTextBox.Text = $script:settings.DefaultFolder
    $defaultFolderTextBox.Location = New-Object System.Drawing.Point(20, 105)
    $defaultFolderTextBox.Size = New-Object System.Drawing.Size(400, 20)
    $tabPaths.Controls.Add($defaultFolderTextBox)

    $defaultFolderBrowseButton = New-Object System.Windows.Forms.Button
    $defaultFolderBrowseButton.Text = "..."
    $defaultFolderBrowseButton.Location = New-Object System.Drawing.Point(430, 103)
    $defaultFolderBrowseButton.Size = New-Object System.Drawing.Size(30, 26)
    $defaultFolderBrowseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Standard-Ordner auswaehlen"
        $folderBrowser.SelectedPath = $defaultFolderTextBox.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $defaultFolderTextBox.Text = $folderBrowser.SelectedPath
        }
    })
    $tabPaths.Controls.Add($defaultFolderBrowseButton)

    # Icon-Pfad
    $iconLabel = New-Object System.Windows.Forms.Label
    $iconLabel.Text = "Icon-Datei:"
    $iconLabel.Location = New-Object System.Drawing.Point(20, 140)
    $iconLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabPaths.Controls.Add($iconLabel)

    $iconTextBox = New-Object System.Windows.Forms.TextBox
    $iconTextBox.Text = $script:settings.IconPath
    $iconTextBox.Location = New-Object System.Drawing.Point(20, 165)
    $iconTextBox.Size = New-Object System.Drawing.Size(400, 20)
    $tabPaths.Controls.Add($iconTextBox)

    $iconBrowseButton = New-Object System.Windows.Forms.Button
    $iconBrowseButton.Text = "..."
    $iconBrowseButton.Location = New-Object System.Drawing.Point(430, 163)
    $iconBrowseButton.Size = New-Object System.Drawing.Size(30, 26)
    $iconBrowseButton.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "Icon-Dateien (*.ico)|*.ico|Alle Dateien (*.*)|*.*"
        $fileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($iconTextBox.Text)
        if ($fileDialog.ShowDialog() -eq "OK") {
            $iconTextBox.Text = $fileDialog.FileName
        }
    })
    $tabPaths.Controls.Add($iconBrowseButton)

    # Tab 2: Anwendung
    $tabApp = New-Object System.Windows.Forms.TabPage
    $tabApp.Text = "Anwendung"
    $tabControl.Controls.Add($tabApp)

    # App-Autor
    $authorLabel = New-Object System.Windows.Forms.Label
    $authorLabel.Text = "Autor:"
    $authorLabel.Location = New-Object System.Drawing.Point(20, 20)
    $authorLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabApp.Controls.Add($authorLabel)

    $authorTextBox = New-Object System.Windows.Forms.TextBox
    $authorTextBox.Text = $script:settings.AppAuthor
    $authorTextBox.Location = New-Object System.Drawing.Point(20, 45)
    $authorTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $tabApp.Controls.Add($authorTextBox)

    # Firma
    $companyLabel = New-Object System.Windows.Forms.Label
    $companyLabel.Text = "Firma:"
    $companyLabel.Location = New-Object System.Drawing.Point(20, 80)
    $companyLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabApp.Controls.Add($companyLabel)

    $companyTextBox = New-Object System.Windows.Forms.TextBox
    $companyTextBox.Text = $script:settings.AppCompany
    $companyTextBox.Location = New-Object System.Drawing.Point(20, 105)
    $companyTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $tabApp.Controls.Add($companyTextBox)

    # Produkt
    $productLabel = New-Object System.Windows.Forms.Label
    $productLabel.Text = "Produkt:"
    $productLabel.Location = New-Object System.Drawing.Point(20, 140)
    $productLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabApp.Controls.Add($productLabel)

    $productTextBox = New-Object System.Windows.Forms.TextBox
    $productTextBox.Text = $script:settings.AppProduct
    $productTextBox.Location = New-Object System.Drawing.Point(20, 165)
    $productTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $tabApp.Controls.Add($productTextBox)

    # Copyright
    $copyrightLabel = New-Object System.Windows.Forms.Label
    $copyrightLabel.Text = "Copyright:"
    $copyrightLabel.Location = New-Object System.Drawing.Point(20, 200)
    $copyrightLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabApp.Controls.Add($copyrightLabel)

    $copyrightTextBox = New-Object System.Windows.Forms.TextBox
    $copyrightTextBox.Text = $script:settings.AppCopyright
    $copyrightTextBox.Location = New-Object System.Drawing.Point(20, 225)
    $copyrightTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $tabApp.Controls.Add($copyrightTextBox)

    # Version
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version:"
    $versionLabel.Location = New-Object System.Drawing.Point(20, 260)
    $versionLabel.Size = New-Object System.Drawing.Size(100, 20)
    $tabApp.Controls.Add($versionLabel)

    $versionTextBox = New-Object System.Windows.Forms.TextBox
    $versionTextBox.Text = $script:settings.AppVersion
    $versionTextBox.Location = New-Object System.Drawing.Point(20, 285)
    $versionTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $tabApp.Controls.Add($versionTextBox)

    # Tab 3: Optionen
    $tabOptions = New-Object System.Windows.Forms.TabPage
    $tabOptions.Text = "Optionen"
    $tabControl.Controls.Add($tabOptions)

    # Administrator-Rechte erforderlich
    $requireAdminCheckBox = New-Object System.Windows.Forms.CheckBox
    $requireAdminCheckBox.Text = "Administrator-Rechte fuer EXE erforderlich"
    $requireAdminCheckBox.Location = New-Object System.Drawing.Point(20, 20)
    $requireAdminCheckBox.Size = New-Object System.Drawing.Size(250, 20)
    $requireAdminCheckBox.Checked = $script:settings.RequireAdmin
    $tabOptions.Controls.Add($requireAdminCheckBox)

    # Konsolen-Fenster ausblenden
    $noConsoleCheckBox = New-Object System.Windows.Forms.CheckBox
    $noConsoleCheckBox.Text = "Konsolen-Fenster ausblenden (GUI-Modus)"
    $noConsoleCheckBox.Location = New-Object System.Drawing.Point(20, 50)
    $noConsoleCheckBox.Size = New-Object System.Drawing.Size(250, 20)
    $noConsoleCheckBox.Checked = $script:settings.NoConsole
    $tabOptions.Controls.Add($noConsoleCheckBox)

    # CPU-Architektur
    $cpuArchLabel = New-Object System.Windows.Forms.Label
    $cpuArchLabel.Text = "CPU-Architektur:"
    $cpuArchLabel.Location = New-Object System.Drawing.Point(20, 140)
    $cpuArchLabel.Size = New-Object System.Drawing.Size(120, 20)
    $tabOptions.Controls.Add($cpuArchLabel)

    $cpuArchComboBox = New-Object System.Windows.Forms.ComboBox
    $cpuArchComboBox.Location = New-Object System.Drawing.Point(20, 165)
    $cpuArchComboBox.Size = New-Object System.Drawing.Size(150, 32)
    $cpuArchComboBox.DropDownStyle = "DropDownList"
    $cpuArchComboBox.Items.AddRange(@("AnyCPU", "x86", "x64"))
    $cpuArchComboBox.Text = $script:settings.CPUArch
    $tabOptions.Controls.Add($cpuArchComboBox)

    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Speichern"
    $saveButton.Location = New-Object System.Drawing.Point(150, 620)
    $saveButton.Size = New-Object System.Drawing.Size(100, 35)
    $saveButton.DialogResult = "OK"
    $saveButton.Add_Click({
        # Einstellungen aktualisieren
        $script:settings.PS2EXEPath = $ps2exeTextBox.Text
        $script:settings.DefaultFolder = $defaultFolderTextBox.Text
        $script:settings.IconPath = $iconTextBox.Text
        $script:settings.AppAuthor = $authorTextBox.Text
        $script:settings.AppCompany = $companyTextBox.Text
        $script:settings.AppProduct = $productTextBox.Text
        $script:settings.AppCopyright = $copyrightTextBox.Text
        $script:settings.AppVersion = $versionTextBox.Text
        $script:settings.RequireAdmin = $requireAdminCheckBox.Checked
        $script:settings.NoConsole = $noConsoleCheckBox.Checked
        $script:settings.CPUArch = $cpuArchComboBox.Text
        $script:settings.TimestampServer = $timestampTextBox.Text
        
        # In Registry speichern
        if (Save-SettingsToRegistry -Settings $script:settings) {
            [System.Windows.Forms.MessageBox]::Show("Einstellungen erfolgreich gespeichert!", "Speichern", "OK", "Information")
        }
    })
    $optionsForm.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Abbrechen"
    $cancelButton.Location = New-Object System.Drawing.Point(260, 620)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.DialogResult = "Cancel"
    $optionsForm.Controls.Add($cancelButton)

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Zurucksetzen"
    $resetButton.Location = New-Object System.Drawing.Point(370, 620)
    $resetButton.Size = New-Object System.Drawing.Size(100, 35)
    $resetButton.Add_Click({
        Reset-SettingsToDefault
        # Dialog neu laden
        $ps2exeTextBox.Text = $script:settings.PS2EXEPath
        $defaultFolderTextBox.Text = $script:settings.DefaultFolder
        $iconTextBox.Text = $script:settings.IconPath
        $authorTextBox.Text = $script:settings.AppAuthor
        $companyTextBox.Text = $script:settings.AppCompany
        $productTextBox.Text = $script:settings.AppProduct
        $copyrightTextBox.Text = $script:settings.AppCopyright
        $versionTextBox.Text = $script:settings.AppVersion
        $requireAdminCheckBox.Checked = $script:settings.RequireAdmin
        $noConsoleCheckBox.Checked = $script:settings.NoConsole
        $cpuArchComboBox.Text = $script:settings.CPUArch
        $timestampTextBox.Text = $script:settings.TimestampServer
    })
    $optionsForm.Controls.Add($resetButton)

    # Dialog anzeigen
    $result = $optionsForm.ShowDialog()
    
    if ($result -eq "OK") {
        # Einstellungen wurden bereits gespeichert
        Update-Info "Einstellungen aktualisiert - Aenderungen werden beim naechsten Neustart wirksam"
    }
}

function Update-Status {
    param([string]$Message, [string]$Color = "FromArgb", [int]$R = 70, [int]$G = 70, [int]$B = 70)
    if ($statusLabel) {
        $statusLabel.Text = $Message
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb($R, $G, $B)
    } else {
        Write-DebugInfo "DEBUG Update-Status: statusLabel ist NULL - Message: $Message"
    }
    
    if ($form) {
        $form.Refresh()
    } else {
        Write-DebugInfo "DEBUG Update-Status: form ist NULL"
    }
}

function Show-Progress {
    param([bool]$Show)
    if ($progressBar) {
        $progressBar.Visible = $Show
    } else {
        Write-DebugInfo "DEBUG Show-Progress: progressBar ist NULL - Show: $Show"
    }
    
    if ($form) {
        $form.Refresh()
    } else {
        Write-DebugInfo "DEBUG Show-Progress: form ist NULL"
    }
}

function Update-Info {
    param([string]$Message)
    # Ersetze alle \n durch echte Zeilenumbrueche
    $formattedMessage = $Message -replace "`n", [Environment]::NewLine
    
    if ($infoTextBox) {
        $infoTextBox.Text = $formattedMessage
        # Scrolle zum Ende
        $infoTextBox.SelectionStart = $infoTextBox.Text.Length
        $infoTextBox.ScrollToCaret()
    } else {
        Write-DebugInfo "DEBUG Update-Info: infoTextBox ist NULL - Message: $Message"
    }
}

function Load-Directory {
    param([string]$Path)
    
    try {
        Write-DebugInfo "DEBUG Load-Directory: Eingabe-Pfad = '$Path'"
        
        # Robuste Validierung des Eingabe-Pfads
        if (-not $Path -or $Path.Trim() -eq "" -or $Path -eq $null) {
            Write-DebugInfo "DEBUG Load-Directory: Eingabe-Pfad ist leer, verwende Fallback"
            if ($script:settings -and $script:settings.DefaultFolder -and $script:settings.DefaultFolder.Trim() -ne "") {
                $Path = $script:settings.DefaultFolder
                Update-Info "Verwende DefaultFolder aus Settings: $Path"
            } else {
                $Path = [Environment]::GetFolderPath("MyDocuments")
                Update-Info "Verwende System-Default (MyDocuments): $Path"
            }
        }
        
        Write-DebugInfo "DEBUG Load-Directory: Verwende Pfad = '$Path'"
        
        # Zusätzliche Validierung: Prüfen ob Pfad existiert
        if (-not (Test-Path $Path)) {
            Write-DebugInfo "DEBUG Load-Directory: Pfad existiert nicht, erstelle ihn: '$Path'"
            try {
                New-Item -ItemType Directory -Path $Path -Force | Out-Null
                Update-Info "Verzeichnis erstellt: $Path"
            }
            catch {
                Write-DebugInfo "DEBUG Load-Directory: Fehler beim Erstellen des Verzeichnisses: $($_.Exception.Message)"
                # Fallback zu MyDocuments
                $Path = [Environment]::GetFolderPath("MyDocuments")
                Update-Info "Fallback zu MyDocuments: $Path"
                if (-not (Test-Path $Path)) {
                    New-Item -ItemType Directory -Path $Path -Force | Out-Null
                }
            }
        }
        
        $script:currentDirectory = $Path
        if ($currentPathLabel) {
            $currentPathLabel.Text = $Path
        } else {
            Write-DebugInfo "DEBUG Load-Directory: currentPathLabel ist NULL"
        }
        
        if ($fileListView) {
            $fileListView.Items.Clear()
        } else {
            Write-DebugInfo "DEBUG Load-Directory: fileListView ist NULL"
            return
        }
        
        # Verzeichnisse hinzufuegen
        $directories = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | Sort-Object Name
        foreach ($dir in $directories) {
            $item = $fileListView.Items.Add($dir.Name)
            $item.SubItems.Add("Ordner")
            $item.SubItems.Add("")
            $item.SubItems.Add("")
            $item.Tag = $dir.FullName
            $item.ImageIndex = 0
        }
        
        # PowerShell-Dateien hinzufuegen
        $psFiles = Get-ChildItem -Path $Path -Filter "*.ps1" -File -ErrorAction SilentlyContinue | Sort-Object Name
        foreach ($file in $psFiles) {
            $item = $fileListView.Items.Add($file.Name)
            $item.SubItems.Add("PS1")
            
            # Signatur-Status pruefen
            try {
                $signature = Get-AuthenticodeSignature -FilePath $file.FullName
                $sigStatus = switch ($signature.Status) {
                    "Valid" { "-> Gueltig" }
                    "NotSigned" { "Nicht signiert" }
                    "UnknownError" { "Fehler" }
                    "NotTrusted" { "Nicht vertrauenswuerdig" }
                    default { "Unbekannt" }
                }
                $item.SubItems.Add($sigStatus)
            }
            catch {
                $item.SubItems.Add("Fehler")
            }
            
            # Dateigroesse hinzufuegen
            $sizeKB = [math]::Round($file.Length / 1KB, 1)
            $item.SubItems.Add("$sizeKB KB")
            
            $item.Tag = $file.FullName
        }
        
        # EXE-Dateien hinzufuegen
        $exeFiles = Get-ChildItem -Path $Path -Filter "*.exe" -File -ErrorAction SilentlyContinue | Sort-Object Name
        foreach ($file in $exeFiles) {
            $item = $fileListView.Items.Add($file.Name)
            $item.SubItems.Add("EXE")
            
            # Signatur-Status pruefen
            try {
                $signature = Get-AuthenticodeSignature -FilePath $file.FullName
                $sigStatus = switch ($signature.Status) {
                    "Valid" { "-> Gueltig" }
                    "NotSigned" { "Nicht signiert" }
                    "UnknownError" { "Fehler" }
                    "NotTrusted" { "Nicht vertrauenswuerdig" }
                    default { "Unbekannt" }
                }
                $item.SubItems.Add($sigStatus)
            }
            catch {
                $item.SubItems.Add("Fehler")
            }
            
            # Dateigroesse hinzufuegen
            $sizeKB = [math]::Round($file.Length / 1KB, 1)
            $item.SubItems.Add("$sizeKB KB")
            
            $item.Tag = $file.FullName
        }
        
        Update-Status "Verzeichnis geladen: $(Split-Path $Path -Leaf) ($($psFiles.Count) PowerShell-Dateien, $($exeFiles.Count) EXE-Dateien)" "FromArgb" 0 100 0
    }
    catch {
        Update-Status "Fehler beim Laden des Verzeichnisses: $($_.Exception.Message)" "FromArgb" 200 0 0
        Write-DebugInfo "DEBUG Load-Directory: Ausnahme = $($_.Exception.Message)"
        Write-DebugInfo "DEBUG Load-Directory: StackTrace = $($_.Exception.StackTrace)"
    }
}

function Connect-SimplySign {
    param([string]$UserId, [string]$OtpUri, [string]$ExePath = "C:\Program Files\SimplySign Desktop\SimplySign.exe")
    
    try {
        if (-not (Test-Path $ExePath)) {
            throw "SimplySign Desktop nicht gefunden: $ExePath"
        }
        
        # TOTP Code generieren (vereinfachte Version)
        $uri = [Uri]$OtpUri
        $query = @{}
        foreach ($part in $uri.Query.TrimStart('?') -split '&') {
            $kv = $part -split '=', 2
            if ($kv.Count -eq 2) { 
                $query[$kv[0]] = [Uri]::UnescapeDataString($kv[1]) 
            }
        }
        
        $secret = $query['secret']
        if (-not $secret) {
            throw "Kein Secret in OTP URI gefunden"
        }
        
        # Einfacher TOTP Generator (nur fuer Demonstrationszwecke)
        $unixTime = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $timeStep = [Math]::Floor($unixTime / 30)
        $otp = ($timeStep % 1000000).ToString("000000")  # Vereinfacht
        
        Update-Info "Verbinde mit SimplySign Desktop...`n`nGenerierter TOTP: $otp`nStarte SimplySign Desktop..."
        
        # SimplySign Desktop starten
        $proc = Start-Process -FilePath $ExePath -PassThru
        Start-Sleep -Seconds 3
        
        # Fenster aktivieren und Credentials senden
        $wshell = New-Object -ComObject WScript.Shell
        $focused = $wshell.AppActivate($proc.Id)
        
        if ($focused) {
            Start-Sleep -Milliseconds 500
            $wshell.SendKeys("$UserId{TAB}$otp{ENTER}")
            Update-Info "SimplySign Desktop Verbindung hergestellt!`n`nCredentials gesendet:`n- User ID: $UserId`n- TOTP Code: $otp`n`nCloud Smart-Card sollte in wenigen Sekunden verfuegbar sein."
            return $true
        } else {
            throw "SimplySign Desktop Fenster konnte nicht aktiviert werden"
        }
    }
    catch {
        Update-Info "SimplySign Verbindungsfehler:`n`n$($_.Exception.Message)`n`nBitte ueberpruefen Sie:`n- SimplySign Desktop installiert?`n- Korrekte OTP URI konfiguriert?`n- User ID korrekt?"
        return $false
    }
}

function Install-TrustedPublisher {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)
    
    if (-not $isAdmin) {
        [System.Windows.Forms.MessageBox]::Show("Administratorrechte erforderlich!", "Fehler", "OK", "Error")
        return $false
    }
    
    try {
        # Code Signing Zertifikat in TrustedPublisher Store installieren
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        
        # Pruefen ob bereits vorhanden
        $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $Certificate.Thumbprint }
        if ($existing) {
            $store.Close()
            Update-Info "Zertifikat bereits als vertrauenswuerdiger Publisher installiert!`n`nName: $($Certificate.Subject)`nThumbprint: $($Certificate.Thumbprint)`n`nSignierte PowerShell-Skripte sollten jetzt ohne Warnung ausgefuehrt werden."
            return $true
        }
        
        # Zertifikat hinzufuegen
        $store.Add($Certificate)
        $store.Close()
        
        Update-Info "Zertifikat erfolgreich als vertrauenswuerdiger Publisher installiert!`n`nName: $($Certificate.Subject)`nThumbprint: $($Certificate.Thumbprint)`n`nSignierte PowerShell-Skripte werden jetzt ohne Warnung ausgefuehrt.`n`nHinweis: Diese Aenderung wirkt sich systemweit aus."
        return $true
    }
    catch {
        Update-Info "Fehler beim Installieren des Trusted Publishers:`n`n$($_.Exception.Message)`n`nMMoegliche Ursachen:`n- Keine Administratorrechte`n- Zertifikat bereits vorhanden`n- Systemrichtlinien verhindern Installation"
        return $false
    }
}

# =============================================================================
# EXE CONVERSION FUNCTIONS
# =============================================================================

# Funktion zum Validieren von Icon-Dateien
function Test-IconFile {
    param([string]$IconPath)
    
    try {
        if (-not (Test-Path $IconPath)) {
            return $false
        }
        
        # Versuche das Icon zu laden um zu pruefen ob es gueltig ist
        $icon = New-Object System.Drawing.Icon($IconPath)
        if ($icon) {
            $icon.Dispose()
            return $true
        }
    }
    catch {
        return $false
    }
    
    return $false
}

# PS2EXE Konvertierung Funktion
function Convert-PS1ToEXE {
    param(
        [string]$PS1FilePath,
        [string]$OutputPath = "",
        [string]$IconPath = "",
        [switch]$NoConsole
    )
    
    try {
        # Parameter validieren
        if (-not $PS1FilePath -or -not (Test-Path $PS1FilePath)) {
            throw "PS1-Datei nicht gefunden: '$PS1FilePath'"
        }
        
        # Output-Pfad generieren falls nicht angegeben
        if (-not $OutputPath -or $OutputPath.Trim() -eq "") {
            $OutputPath = $PS1FilePath -replace '\.ps1$', '.exe'
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: OutputPath war leer, generiert: '$OutputPath'"
        }
        
        # Pruefen ob PS2EXE als Modul verfuegbar ist
        $ps2exeModule = Get-Module -Name ps2exe -ListAvailable | Select-Object -First 1
        
        if ($ps2exeModule) {
            # PS2EXE als Modul verwenden
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: Verwende PS2EXE Modul $($ps2exeModule.Version)"
            
            # Modul importieren falls nicht bereits geladen
            if (-not (Get-Module -Name ps2exe)) {
                Import-Module ps2exe
            }
            
            # Parameter fuer Invoke-ps2exe aufbauen
            $invokeParams = @{
                InputFile = $PS1FilePath
                OutputFile = $OutputPath
            }
            
            # Icon aus Registry verwenden falls kein spezifischer Pfad angegeben
            if (-not $IconPath -and $script:settings.IconPath -and $script:settings.IconPath -ne "" -and (Test-IconFile $script:settings.IconPath)) {
                $IconPath = $script:settings.IconPath
            }
            
            if ($IconPath -and (Test-IconFile $IconPath) -and $IconPath -ne "") {
                $invokeParams.IconFile = $IconPath
            } elseif ($IconPath -and $IconPath -ne "") {
                Update-Info "Icon-Datei nicht gueltig oder nicht gefunden: $IconPath - Verwende Standard-Icon"
            }
            
            # NoConsole aus Registry verwenden falls nicht explizit angegeben
            if (-not $NoConsole -and $script:settings.NoConsole) {
                $invokeParams.NoConsole = $true
            } elseif ($NoConsole) {
                $invokeParams.NoConsole = $true
            }
            
            # Zusaetzliche Parameter aus Registry
            if ($script:settings.RequireAdmin) {
                $invokeParams.RequireAdmin = $true
            }
            
            # CPU-Architektur aus Registry
            if ($script:settings.CPUArch -and $script:settings.CPUArch -ne "AnyCPU" -and $script:settings.CPUArch -ne "") {
                $invokeParams.Architecture = $script:settings.CPUArch
            }
            
            # App-Informationen hinzufuegen (immer wenn Company gesetzt ist)
            if ($script:settings.AppCompany -and $script:settings.AppCompany -ne "") {
                $invokeParams.Title = if ($script:settings.AppProduct -and $script:settings.AppProduct -ne "") { $script:settings.AppProduct } else { "PowerShell Tool" }
                $invokeParams.Description = "Erstellt mit PhinIT CERTUM Code Signing Tool"
                $invokeParams.Company = $script:settings.AppCompany
                $invokeParams.Product = if ($script:settings.AppProduct -and $script:settings.AppProduct -ne "") { $script:settings.AppProduct } else { "PowerShell Tool" }
                $invokeParams.Copyright = if ($script:settings.AppCopyright -and $script:settings.AppCopyright -ne "") { $script:settings.AppCopyright } else { "2025 PhinIT" }
                $invokeParams.Version = if ($script:settings.AppVersion -and $script:settings.AppVersion -ne "") { $script:settings.AppVersion } else { "1.0.0.0" }
            }
            
            # Debug-Ausgabe
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: Invoke-ps2exe Parameter = $($invokeParams | Out-String)"
            
            # Zusätzliche Validierung: Prüfe auf leere Pfade in Invoke-ps2exe Parametern
            foreach ($key in $invokeParams.Keys) {
                $value = $invokeParams[$key]
                if ($value -eq "" -or $value -eq $null) {
                    Write-DebugInfo "DEBUG Convert-PS1ToEXE: LEERER PARAMETER GEFUNDEN - Key: '$key', Value: '$value'"
                    throw "PS2EXE Parameter '$key' ist leer oder nicht definiert"
                }
                Write-DebugInfo "DEBUG Convert-PS1ToEXE: Parameter '$key' = '$value'"
            }
            
            # PS2EXE als Modul ausfuehren
            Invoke-ps2exe @invokeParams
            
        } else {
            # Fallback: PS2EXE als Script verwenden
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: PS2EXE Modul nicht gefunden, verwende Script-Modus"
            
            # PS2EXE Module Pfad aus Registry verwenden (mit Fallback)
            $ps2exePath = $script:settings.PS2EXEPath
            if (-not $ps2exePath) {
                # Fallback-Pfad verwenden
                $ps2exePath = Join-Path $PSScriptRoot "ps2exe"
                Write-DebugInfo "DEBUG Convert-PS1ToEXE: Verwende Fallback-Pfad: $ps2exePath"
            }
            
            if (-not $ps2exePath) {
                throw "PS2EXE Pfad ist nicht konfiguriert. Bitte installieren Sie PS2EXE oder konfigurieren Sie den Pfad in den Optionen."
            }
            
            $ps2exeScript = Join-Path $ps2exePath "ps2exe.ps1"
            
            # PS2EXE Script Pfad validieren
            if (-not $ps2exeScript -or $ps2exeScript -eq "") {
                throw "PS2EXE Script-Pfad ist leer oder nicht definiert"
            }
            
            if (-not (Test-Path $ps2exeScript)) {
                throw "PS2EXE Script nicht gefunden: '$ps2exeScript'. Bitte installieren Sie PS2EXE oder konfigurieren Sie den Pfad in den Optionen."
            }
            
            # PS2EXE Parameter als Array aufbauen
            $params = @(
                "-inputFile", "`"$PS1FilePath`"",
                "-outputFile", "`"$OutputPath`""
            )
            
            # Icon aus Registry verwenden falls kein spezifischer Pfad angegeben
            if (-not $IconPath -and $script:settings.IconPath -and $script:settings.IconPath -ne "" -and (Test-IconFile $script:settings.IconPath)) {
                $IconPath = $script:settings.IconPath
            }
            
            if ($IconPath -and (Test-IconFile $IconPath) -and $IconPath -ne "") {
                $params += "-iconFile", "`"$IconPath`""
            } elseif ($IconPath -and $IconPath -ne "") {
                Update-Info "Icon-Datei nicht gueltig oder nicht gefunden: $IconPath - Verwende Standard-Icon"
            }
            
            # NoConsole aus Registry verwenden falls nicht explizit angegeben
            if (-not $NoConsole -and $script:settings.NoConsole) {
                $params += "-noConsole"
            } elseif ($NoConsole) {
                $params += "-noConsole"
            }
            
            # Zusaetzliche Parameter aus Registry
            if ($script:settings.RequireAdmin) {
                $params += "-requireAdmin"
            }
            
            # CPU-Architektur aus Registry
            if ($script:settings.CPUArch -and $script:settings.CPUArch -ne "AnyCPU" -and $script:settings.CPUArch -ne "") {
                $params += "-architecture", $script:settings.CPUArch
            }
            
            # App-Informationen hinzufuegen (immer wenn Company gesetzt ist)
            if ($script:settings.AppCompany) {
                $params += "-title", "`"$($script:settings.AppProduct)`""
                $params += "-description", "`"Erstellt mit PhinIT CERTUM Code Signing Tool`""
                $params += "-company", "`"$($script:settings.AppCompany)`""
                $params += "-product", "`"$($script:settings.AppProduct)`""
                $params += "-copyright", "`"$($script:settings.AppCopyright)`""
                $params += "-version", "`"$($script:settings.AppVersion)`""
            }
            
            # Debug-Ausgabe
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: PS2EXE Script = $ps2exeScript"
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: Parameter = $($params -join ' ')"
            
            # Zusätzliche Validierung: Prüfe auf leere Pfade in Script-Parametern
            for ($i = 0; $i -lt $params.Count; $i++) {
                $param = $params[$i]
                if ($param -eq "" -or $param -eq $null) {
                    Write-DebugInfo "DEBUG Convert-PS1ToEXE: LEERER PARAMETER GEFUNDEN - Index: $i, Value: '$param'"
                    throw "PS2EXE Parameter an Index $i ist leer oder nicht definiert"
                }
                Write-DebugInfo "DEBUG Convert-PS1ToEXE: Parameter[$i] = '$param'"
            }
            
            # Zusätzliche Validierung: Prüfe auf leere Pfade in Parametern
            foreach ($param in $params) {
                if ($param -eq "" -or $param -eq $null) {
                    throw "PS2EXE Parameter enthält leeren Pfad. Parameter-Array: $($params -join ', ')"
                }
            }
            
            # Spezifische Validierung der kritischen Pfad-Parameter
            if (-not $PS1FilePath -or $PS1FilePath -eq "") {
                throw "PS1FilePath ist leer oder nicht definiert: '$PS1FilePath'"
            }
            if (-not $OutputPath -or $OutputPath -eq "") {
                throw "OutputPath ist leer oder nicht definiert: '$OutputPath'"
            }
            if (-not $ps2exeScript -or $ps2exeScript -eq "") {
                throw "PS2EXE Script-Pfad ist leer oder nicht definiert: '$ps2exeScript'"
            }
            
            # PS2EXE als Script ausführen
            $result = & $ps2exeScript @params
            
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: PS2EXE Result = $result"
        }
        
        if (Test-Path $OutputPath) {
            Write-DebugInfo "DEBUG Convert-PS1ToEXE: EXE erfolgreich erstellt: $OutputPath"
            return $OutputPath
        } else {
            throw "EXE-Datei wurde nicht erstellt: '$OutputPath'"
        }
    }
    catch {
        Write-DebugInfo "DEBUG Convert-PS1ToEXE: Fehler = $($_.Exception.Message)"
        throw "PS2EXE Konvertierung fehlgeschlagen: $($_.Exception.Message)"
    }
}

# Button-Status aktualisieren basierend auf Auswahl
function Update-ButtonStates {
    try {
        $hasFile = ($script:selectedScriptPath -and $script:selectedScriptPath -ne "" -and (Test-Path $script:selectedScriptPath))
        $hasCert = ($null -ne $script:selectedCertificate)
        $isPS1 = $hasFile -and ([System.IO.Path]::GetExtension($script:selectedScriptPath) -eq ".ps1")
        $isEXE = $hasFile -and ([System.IO.Path]::GetExtension($script:selectedScriptPath) -eq ".exe")
        
        Update-Info "DEBUG: Update-ButtonStates - hasFile: $hasFile, hasCert: $hasCert, isPS1: $isPS1, isEXE: $isEXE"
        
        # PS1 SIGNIEREN: PS1-Datei + Zertifikat
        if ($signPS1Button) {
            $signPS1Button.Enabled = ($isPS1 -and $hasCert)
        } else {
            Write-DebugInfo "DEBUG: signPS1Button ist NULL"
        }
        
        # PS1 ? EXE: PS1-Datei
        if ($convertToEXEButton) {
            $convertToEXEButton.Enabled = $isPS1
        } else {
            Write-DebugInfo "DEBUG: convertToEXEButton ist NULL"
        }
        
        # EXE SIGNIEREN: EXE-Datei + Zertifikat  
        if ($signEXEButton) {
            $signEXEButton.Enabled = ($isEXE -and $hasCert)
        } else {
            Write-DebugInfo "DEBUG: signEXEButton ist NULL"
        }

        # SimplySign: immer verfuegbar
        if ($simplySignButton) {
            $simplySignButton.Enabled = $true
        } else {
            Write-DebugInfo "DEBUG: simplySignButton ist NULL"
        }
        
        Update-Info "DEBUG: Button-Status aktualisiert - PS1 Sign: $($signPS1Button.Enabled), PS1->EXE: $($convertToEXEButton.Enabled), EXE Sign: $($signEXEButton.Enabled)"
    }
    catch {
        Write-DebugInfo "DEBUG: Fehler in Update-ButtonStates: $($_.Exception.Message)"
        # Bei Fehlern alle Buttons deaktivieren
        if ($signPS1Button) { $signPS1Button.Enabled = $false }
        if ($convertToEXEButton) { $convertToEXEButton.Enabled = $false }
        if ($signEXEButton) { $signEXEButton.Enabled = $false }
        if ($simplySignButton) { $simplySignButton.Enabled = $false }
    }
}

# Prueft ob eine Datei digital signiert ist
function Test-FileSignature {
    param([string]$FilePath)
    
    try {
        $signature = Get-AuthenticodeSignature -FilePath $FilePath
        return ($signature.Status -eq "Valid")
    }
    catch {
        return $false
    }
}

# =============================================================================
# EVENT HANDLERS
# =============================================================================

# Verzeichnis-Navigation
$upButton.Add_Click({
    $parent = Split-Path $script:currentDirectory -Parent
    if ($parent) {
        Load-Directory $parent
    }
})

$homeButton.Add_Click({
    Load-Directory ([Environment]::GetFolderPath("MyDocuments"))
})

# Ordner-Browser
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "PowerShell-Ordner auswaehlen"
    $folderBrowser.SelectedPath = $script:currentDirectory
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        Load-Directory $folderBrowser.SelectedPath
    }
})

# Datei-Auswahl
$fileListView.Add_DoubleClick({
    try {
        Update-Info "DEBUG: Datei-Explorer DoubleClick ausgefuehrt"
        
        # NULL-Check für fileListView
        if (-not $fileListView) {
            Update-Info "DEBUG: fileListView ist NULL"
            return
        }
        
        if ($fileListView.SelectedItems.Count -eq 0) { 
            Update-Info "DEBUG: Kein Element ausgewaehlt"
            return 
        }
        
        $selectedItem = $fileListView.SelectedItems[0]
        
        # NULL-Check für selectedItem
        if (-not $selectedItem) {
            Update-Info "DEBUG: selectedItem ist NULL"
            return
        }
        
        $path = $selectedItem.Tag
        Update-Info "DEBUG: Ausgewaehltes Element: $path"
        
        # NULL-Check für path
        if (-not $path -or $path.Trim() -eq "") {
            Update-Info "DEBUG: Pfad ist leer oder NULL"
            return
        }
        
        # Pruefen ob Pfad existiert
        if (-not (Test-Path $path)) {
            Update-Info "DEBUG: Pfad existiert nicht: $path"
            [System.Windows.Forms.MessageBox]::Show("Die ausgewaehlte Datei/der Ordner existiert nicht mehr.", "Pfad nicht gefunden", "OK", "Warning")
            return
        }

        if ((Get-Item $path).PSIsContainer) {
            Update-Info "DEBUG: Verzeichnis oeffnen: $path"
            # Verzeichnis oeffnen
            Load-Directory $path
        } else {
            Update-Info "DEBUG: PowerShell-Datei auswaehlen: $path"
            # PowerShell-Datei auswaehlen
            $script:selectedScriptPath = $path
            
            # NULL-Check für selectedFileDisplay
            if ($selectedFileDisplay) {
                $selectedFileDisplay.Text = Split-Path $path -Leaf
                $selectedFileDisplay.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
            } else {
                Update-Info "DEBUG: selectedFileDisplay ist NULL"
            }
            
            Update-ButtonStates
            Update-Status "Datei ausgewaehlt: $(Split-Path $path -Leaf)"
            
            # Datei-Info anzeigen
            try {
                $signature = Get-AuthenticodeSignature -FilePath $path
                $info = "Ausgewaehlte Datei: $(Split-Path $path -Leaf)`n`n"
                $info += "Pfad: $path`n"
                $info += "Groesse: $([math]::Round((Get-Item $path).Length / 1KB, 2)) KB`n"
                $info += "Geaendert: $((Get-Item $path).LastWriteTime.ToString('dd.MM.yyyy HH:mm'))`n`n"
                $info += "Signatur-Status: $($signature.Status)`n"
                
                if ($signature.SignerCertificate) {
                    $info += "Signiert von: $($signature.SignerCertificate.Subject)`n"
                    $info += "Zeitstempel: $($signature.TimeStamperCertificate.Subject)"
                } else {
                    $info += "Datei ist nicht digital signiert."
                }
                
                Update-Info $info
            }
            catch {
                Update-Info "Fehler beim Lesen der Datei-Informationen: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Update-Info "DEBUG: Fehler im DoubleClick Event: $($_.Exception.Message)"
        Write-DebugInfo "DEBUG: StackTrace: $($_.Exception.StackTrace)"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Oeffnen der Datei/des Ordners:`n`n$($_.Exception.Message)", "Auswahlfhler", "OK", "Error")
    }
})

# Zertifikat-Verwaltung
$refreshCertButton.Add_Click({
    Update-Info "DEBUG: Zertifikat-Refresh Button wurde geklickt"
    Show-Progress $true
    Update-Status "Lade CERTUM Zertifikate..." "FromArgb" 0 0 200
    $certComboBox.Items.Clear()
    
    try {
        Update-Info "DEBUG: Suche nach Code Signing Zertifikaten..."
        $certs = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Where-Object {
            $_.EnhancedKeyUsageList -match "Code Signing" -or
            $_.Issuer -match "CERTUM"
        } | Sort-Object NotAfter -Descending
        
        Update-Info "DEBUG: Gefundene Zertifikate: $($certs.Count)"
        
        if ($certs.Count -eq 0) {
            $certComboBox.Items.Add("Kein CERTUM Code Signing Zertifikat gefunden")
            Update-Status "Keine CERTUM Zertifikate gefunden - Bitte installieren" "FromArgb" 200 0 0
            Update-Info "DEBUG: Keine Zertifikate gefunden"
        } else {
            foreach ($cert in $certs) {
                $subject = ($cert.Subject -split ',')[0] -replace 'CN=', ''
                $validUntil = $cert.NotAfter.ToString('dd.MM.yyyy')
                $isExpired = $cert.NotAfter -lt (Get-Date)

                $status = if ($isExpired) { "ABGELAUFEN" } else { "GUELTIG" }
                $item = "$status | $subject | Gueltig bis: $validUntil"
                
                $certComboBox.Items.Add($item)
                $certComboBox.Tag = $certs
            }
            
            if ($certComboBox.Items.Count -gt 0) {
                $certComboBox.SelectedIndex = 0
                $script:selectedCertificate = $certs[0]
                Update-ButtonStates
                Update-Status "$($certs.Count) CERTUM Zertifikat(e) geladen" "FromArgb" 0 100 0
                Update-Info "DEBUG: Zertifikate erfolgreich geladen, Button-Status aktualisiert"
            }
        }
    }
    catch {
        $certComboBox.Items.Add("Fehler beim Laden der Zertifikate")
        Update-Status "Fehler beim Laden der Zertifikate: $($_.Exception.Message)" "FromArgb" 200 0 0
        Update-Info "DEBUG: Fehler beim Laden der Zertifikate: $($_.Exception.Message)"
    }
    finally {
        Show-Progress $false
    }
})

$certComboBox.Add_SelectedIndexChanged({
    if ($certComboBox.SelectedIndex -ge 0 -and $certComboBox.Tag) {
        $certs = $certComboBox.Tag
        if ($certComboBox.SelectedIndex -lt $certs.Count) {
            $script:selectedCertificate = $certs[$certComboBox.SelectedIndex]
            Update-ButtonStates
            
            $certName = ($script:selectedCertificate.Subject -split ',')[0] -replace 'CN=', ''
            Update-Status "Zertifikat ausgewaehlt: $certName"
        }
    }
})

# Alte Event Handler entfernt - siehe neue EXE-fokussierte Handler oben

# SimplySign Integration
$simplySignButton.Add_Click({
    Update-Info "DEBUG: SimplySign Button wurde geklickt"
    
    # Pruefen ob SimplySignDesktop.exe bereits laeuft
    $simplySignProcess = Get-Process -Name "SimplySignDesktop" -ErrorAction SilentlyContinue
    
    if ($simplySignProcess) {
        Update-Info "SimplySign Desktop laeuft bereits (PID: $($simplySignProcess.Id))`n`nDer Prozess ist aktiv und bereit zur Verwendung."
        Update-Status "SimplySign Desktop bereits aktiv" "FromArgb" 0 150 0
        return
    }
    
    # Pfad zu SimplySign Desktop
    $simplySignPath = "C:\Program Files\Certum\SimplySign Desktop\SimplySignDesktop.exe"
    
    if (-not (Test-Path $simplySignPath)) {
        Update-Info "SimplySign Desktop nicht gefunden:`n$simplySignPath`n`nBitte ueberpruefen Sie, ob SimplySign Desktop korrekt installiert ist."
        Update-Status "SimplySign Desktop nicht gefunden" "FromArgb" 200 0 0
        [System.Windows.Forms.MessageBox]::Show("SimplySign Desktop wurde nicht gefunden.`n`nPfad: $simplySignPath`n`nBitte stellen Sie sicher, dass SimplySign Desktop installiert ist.", "SimplySign nicht gefunden", "OK", "Warning")
        return
    }
    
    Update-Info "SimplySign Desktop wird gestartet...`nPfad: $simplySignPath"
    Update-Status "Starte SimplySign Desktop..." "FromArgb" 0 0 200
    Show-Progress $true
    
    try {
        # SimplySign Desktop starten
        $process = Start-Process -FilePath $simplySignPath -PassThru
        
        # Kurz warten und dann pruefen ob der Prozess laeuft
        Start-Sleep -Seconds 3
        
        $runningProcess = Get-Process -Name "SimplySignDesktop" -ErrorAction SilentlyContinue
        if ($runningProcess) {
            Update-Info "SimplySign Desktop erfolgreich gestartet!`nPID: $($runningProcess.Id)`n`nDer Prozess ist nun aktiv und bereit zur Verwendung."
            Update-Status "SimplySign Desktop gestartet" "FromArgb" 0 150 0
        } else {
            throw "SimplySign Desktop konnte nicht gestartet werden"
        }
    }
    catch {
        Update-Info "Fehler beim Starten von SimplySign Desktop:`n$($_.Exception.Message)"
        Update-Status "Fehler beim Starten" "FromArgb" 200 0 0
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Starten von SimplySign Desktop:`n`n$($_.Exception.Message)", "Startfehler", "OK", "Error")
    }
    finally {
        Show-Progress $false
    }
})

# Alte Trust Installation Event Handler entfernt - siehe EXE-Workflow oben

# PS1 SIGNIEREN Event Handler
$signPS1Button.Add_Click({
    Update-Info "DEBUG: PS1 SIGNIEREN Button wurde geklickt"
    
    if (-not $script:selectedScriptPath -or -not (Test-Path $script:selectedScriptPath)) {
        Update-Info "DEBUG: Keine Datei ausgewaehlt oder Datei existiert nicht"
        [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zunaechst eine PowerShell-Datei aus dem Explorer aus.", "Keine Datei ausgewaehlt", "OK", "Warning")
        return
    }
    
    if ([System.IO.Path]::GetExtension($script:selectedScriptPath) -ne ".ps1") {
        Update-Info "DEBUG: Ausgewaehlte Datei ist keine .ps1 Datei"
        [System.Windows.Forms.MessageBox]::Show("Die ausgewaehlte Datei muss eine PowerShell-Datei (.ps1) sein.", "Falscher Dateityp", "OK", "Warning")
        return
    }
    
    if (-not $script:selectedCertificate) {
        Update-Info "DEBUG: Kein Zertifikat ausgewaehlt"
        [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zunaechst ein gueltiges CERTUM Zertifikat aus.", "Kein Zertifikat ausgewaehlt", "OK", "Warning")
        return
    }

    Update-Info "DEBUG: Alle Pruefungen bestanden, starte Signierung"
    
    try {
        $fileName = Split-Path $script:selectedScriptPath -Leaf
        Update-Status "Signiere PowerShell-Script '$fileName'..." "FromArgb" 0 100 0
        Update-Info "PS1-Signierung gestartet...`n`nDatei: $fileName`nZertifikat: $($script:selectedCertificate.Subject)`nTimestamp: $(if ($timestampCheckBox.Checked) { "Aktiviert" } else { "Deaktiviert" })`nStatus: Signierung laeuft...`n`nBitte warten..."
        Show-Progress $true
        $signPS1Button.Enabled = $false
        
        # Signierung durchfuehren
        $params = @{
            FilePath = $script:selectedScriptPath
            Certificate = $script:selectedCertificate
            HashAlgorithm = "SHA256"
        }
        
        if ($timestampCheckBox.Checked) {
            $params.TimestampServer = $script:settings.TimestampServer
        }
        
        Set-AuthenticodeSignature @params | Out-Null
        
        # Signatur pruefen
        $signature = Get-AuthenticodeSignature -FilePath $script:selectedScriptPath
        
        if ($signature.Status -eq "Valid") {
            Update-Status "PowerShell-Script erfolgreich signiert!" "FromArgb" 0 150 0
            Update-Info "PS1-Signierung erfolgreich!`n`nDatei: $fileName`nStatus: Digital signiert`nTimestamp: $(if ($timestampCheckBox.Checked) { "Hinzugefuegt" } else { "Nicht verwendet" })`n`nNaechster Schritt: 'PS1 zu EXE' fuer ExecutionPolicy-freie Ausfuehrung"
            
            # Kurze Verzoegerung um sicherzustellen, dass die Signatur vollstaendig geschrieben wurde
            Start-Sleep -Milliseconds 300
            
            # Explorer aktualisieren - mehrmals um sicherzustellen
            Load-Directory $script:currentDirectory
            Start-Sleep -Milliseconds 200
            Load-Directory $script:currentDirectory
        } else {
            throw "Signaturpruefung fehlgeschlagen: $($signature.Status)"
        }
        
    } catch {
        Update-Status "Fehler beim Signieren: $($_.Exception.Message)" "Red"
        Update-Info "Signierungsfehler:`n`n Fehlerdetails:`n$($_.Exception.Message)`n`n Moegliche Ursachen:`n Zertifikat ist abgelaufen`n Datei ist schreibgeschuetzt`n SimplySign nicht aktiviert`n Netzwerkprobleme (Timestamp)"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der PS1-Signierung:`n`n$($_.Exception.Message)", "Signierungsfehler", "OK", "Error")
    } finally {
        Show-Progress $false
        $signPS1Button.Enabled = $true
    }
})

# PS1 ? EXE Konvertierung Event Handler
$convertToEXEButton.Add_Click({
    Update-Info "DEBUG: PS1 - EXE Button wurde geklickt`nDEBUG: selectedScriptPath = '$script:selectedScriptPath'"
    
    if (-not $script:selectedScriptPath -or -not (Test-Path $script:selectedScriptPath)) {
        Update-Info "DEBUG: Keine Datei ausgewaehlt oder Datei existiert nicht`nDEBUG: selectedScriptPath = '$script:selectedScriptPath'`nDEBUG: Test-Path Ergebnis = $(if ($script:selectedScriptPath) { Test-Path $script:selectedScriptPath } else { 'N/A' })"
        [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zunaechst eine PowerShell-Datei (.ps1) aus dem Explorer aus.", "Keine Datei ausgewaehlt", "OK", "Warning")
        return
    }
    
    if ([System.IO.Path]::GetExtension($script:selectedScriptPath) -ne ".ps1") {
        Update-Info "DEBUG: Ausgewaehlte Datei ist keine .ps1 Datei`nDEBUG: Extension = $([System.IO.Path]::GetExtension($script:selectedScriptPath))"
        [System.Windows.Forms.MessageBox]::Show("Die ausgewaehlte Datei muss eine PowerShell-Datei (.ps1) sein.", "Falscher Dateityp", "OK", "Warning")
        return
    }

    Update-Info "DEBUG: Alle Pruefungen bestanden, starte EXE-Konvertierung`nDEBUG: Final selectedScriptPath = '$script:selectedScriptPath'"
    
    try {
        $fileName = Split-Path $script:selectedScriptPath -Leaf
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        Update-Status "Konvertiere '$fileName' zu EXE..." "FromArgb" 0 100 200
        Update-Info "- PS1 - EXE Konvertierung gestartet...`n`n Datei: $fileName`n Ziel: $baseName.exe`n  Tool: PS2EXE Modul`n Status: Konvertierung laeuft...`n`nBitte warten..."
        Show-Progress $true
        $convertToEXEButton.Enabled = $false

        # EXE-Konvertierung durchfuehren - mit minimalen Parametern
        $exePath = Convert-PS1ToEXE -PS1FilePath $script:selectedScriptPath
        Update-Info "DEBUG: Convert-PS1ToEXE abgeschlossen, EXE-Pfad = '$exePath'"
        
        if (Test-Path $exePath) {
            # Zusätzliche Validierung des EXE-Pfads
            Write-DebugInfo "DEBUG: EXE-Pfad vor Korrektur = '$exePath'"
            
            # Sicherstellen, dass der Pfad mit .exe endet
            if ($exePath -notmatch '\.exe$') {
                $correctedPath = $exePath -replace '\.ex$', '.exe'
                Write-DebugInfo "DEBUG: Korrigiere EXE-Pfad von '$exePath' zu '$correctedPath'"
                $exePath = $correctedPath
            }
            
            # Finale Validierung
            if (-not (Test-Path $exePath)) {
                Write-DebugInfo "DEBUG: Korrigierter Pfad existiert nicht: '$exePath'"
                throw "EXE-Datei wurde nicht am erwarteten Pfad gefunden: '$exePath'"
            }
            
            Write-DebugInfo "DEBUG: Finale EXE-Pfad validiert: '$exePath'"
            Update-Status "EXE-Konvertierung erfolgreich!" "FromArgb" 0 150 0
            Update-Info "- PS1 - EXE Konvertierung erfolgreich!`n`n Original: $fileName`n EXE-Datei: $([System.IO.Path]::GetFileName($exePath))`n Groesse: $([Math]::Round((Get-Item $exePath).Length / 1KB, 1)) KB`n Icon: ico-app.ico`n`n Vorteil: EXE umgeht PowerShell ExecutionPolicy`n Naechster Schritt: 'EXE SIGNIEREN' fuer digitale Signatur"

            # Kurze Verzoegerung um sicherzustellen, dass die Datei vollstaendig geschrieben wurde
            Start-Sleep -Milliseconds 500
            
            # DEBUG: Verzeichnis-Pfad pruefen
            Write-DebugInfo "DEBUG: currentDirectory vor Explorer-Update = '$script:currentDirectory'"
            
            # Explorer aktualisieren - mit Fallback falls currentDirectory leer ist
            if ($script:currentDirectory -and $script:currentDirectory -ne "") {
                Write-DebugInfo "DEBUG: Verwende currentDirectory: '$script:currentDirectory'"
                Load-Directory $script:currentDirectory
                Start-Sleep -Milliseconds 200
                Load-Directory $script:currentDirectory
            } else {
                # Fallback: Verzeichnis aus EXE-Pfad verwenden
                $exeDir = Split-Path $exePath -Parent
                Write-DebugInfo "DEBUG: currentDirectory leer, verwende EXE-Verzeichnis: '$exeDir'"
                if ($exeDir -and $exeDir -ne "") {
                    Load-Directory $exeDir
                    Start-Sleep -Milliseconds 200
                    Load-Directory $exeDir
                } else {
                    Write-DebugInfo "DEBUG: Auch EXE-Verzeichnis ist leer - ueberspringe Explorer-Update"
                }
            }
            
            # Neue EXE-Datei automatisch auswaehlen - mit Validierung
            if ($exePath -and $exePath -ne "" -and (Test-Path $exePath)) {
                Write-DebugInfo "DEBUG: Setze selectedScriptPath auf: '$exePath'"
                $script:selectedScriptPath = $exePath
                $selectedFileDisplay.Text = Split-Path $exePath -Leaf
                $selectedFileDisplay.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
                Update-ButtonStates
            } else {
                Write-DebugInfo "DEBUG: exePath ist ungueltig oder existiert nicht: '$exePath'"
                Update-Info "WARNUNG: EXE-Datei wurde erstellt, konnte aber nicht automatisch ausgewaehlt werden"
            }
        } else {
            throw "EXE-Datei wurde nicht erstellt"
        }
        
    } catch {
        Update-Status "EXE-Konvertierung fehlgeschlagen" "Red"
        Update-Info "- PS1 - EXE Konvertierung fehlgeschlagen:`n`n Fehlerdetails:`n$($_.Exception.Message)`n`n Moegliche Ursachen:`n PS2EXE Modul nicht verfuegbar`n PowerShell-Datei syntaktisch fehlerhaft`n Unzureichende Schreibrechte`n Ungueltiger Icon-Pfad`n`n Debug-Info:`nPfad: $script:selectedScriptPath`nIcon: $(Join-Path $PSScriptRoot 'assets\ico-app.ico')"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der EXE-Konvertierung:`n`n$($_.Exception.Message)", "Konvertierungsfehler", "OK", "Error")
    } finally {
        Show-Progress $false
        $convertToEXEButton.Enabled = $true
    }
})

# EXE SIGNIEREN Event Handler
$signEXEButton.Add_Click({
    Update-Info "DEBUG: EXE SIGNIEREN Button wurde geklickt"
    
    if (-not $script:selectedScriptPath -or -not (Test-Path $script:selectedScriptPath)) {
        Update-Info "DEBUG: Keine Datei ausgewaehlt oder Datei existiert nicht"
        [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zunaechst eine EXE-Datei aus dem Explorer aus.", "Keine Datei ausgewaehlt", "OK", "Warning")
        return
    }
    
    if ([System.IO.Path]::GetExtension($script:selectedScriptPath) -ne ".exe") {
        Update-Info "DEBUG: Ausgewaehlte Datei ist keine .exe Datei"
        [System.Windows.Forms.MessageBox]::Show("Die ausgewaehlte Datei muss eine EXE-Datei sein.", "Falscher Dateityp", "OK", "Warning")
        return
    }
    
    if (-not $script:selectedCertificate) {
        Update-Info "DEBUG: Kein Zertifikat ausgewaehlt"
        [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zunaechst ein gueltiges CERTUM Zertifikat aus.", "Kein Zertifikat ausgewaehlt", "OK", "Warning")
        return
    }
    
    Update-Info "DEBUG: Alle Pruefungen bestanden, starte EXE-Signierung"

    try {
        $fileName = Split-Path $script:selectedScriptPath -Leaf
        Update-Status "Signiere EXE-Datei '$fileName'..." "FromArgb" 0 150 0
        Update-Info "EXE-Signierung gestartet...`n`nDatei: $fileName`nZertifikat: $($script:selectedCertificate.Subject)`nTimestamp: $(if ($timestampCheckBox.Checked) { "Aktiviert" } else { "Deaktiviert" })`nStatus: Signierung laeuft...`n`nBitte warten..."
        Show-Progress $true
        $signEXEButton.Enabled = $false
        
        # EXE signieren
        $params = @{
            FilePath = $script:selectedScriptPath
            Certificate = $script:selectedCertificate
            HashAlgorithm = "SHA256"
        }
        
        if ($timestampCheckBox.Checked) {
            $params.TimestampServer = $script:settings.TimestampServer
        }
        
        Set-AuthenticodeSignature @params | Out-Null
        
        # Signatur pruefen
        $signature = Get-AuthenticodeSignature -FilePath $script:selectedScriptPath
        
        if ($signature.Status -eq "Valid") {
            Update-Status "EXE erfolgreich signiert!" "FromArgb" 0 150 0
            Update-Info "EXE-Signierung erfolgreich!`n`n Datei: $fileName`n Status: Digital signiert`n Zertifikat: $($script:selectedCertificate.Subject)`n Timestamp: $(if ($timestampCheckBox.Checked) { "Hinzugefuegt" } else { "Nicht verwendet" })`n`n Fertig! Die EXE kann jetzt ohne PowerShell ExecutionPolicy-Probleme ausgefuehrt werden."

            # Kurze Verzoegerung um sicherzustellen, dass die Signatur vollstaendig geschrieben wurde
            Start-Sleep -Milliseconds 300
            
            # Explorer aktualisieren - mehrmals um sicherzustellen
            Load-Directory $script:currentDirectory
            Start-Sleep -Milliseconds 200
            Load-Directory $script:currentDirectory
        } else {
            throw "Signaturpruefung fehlgeschlagen: $($signature.Status)"
        }
        
    } catch {
        Update-Status "EXE-Signierung fehlgeschlagen" "Red"
        Update-Info "EXE-Signierungsfehler:`n`n Fehlerdetails:`n$($_.Exception.Message)`n`n Moegliche Ursachen:`n Zertifikat ist abgelaufen`n EXE-Datei ist in Verwendung`n SimplySign nicht aktiviert`n Unzureichende Schreibrechte"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der EXE-Signierung:`n`n$($_.Exception.Message)", "Signierungsfehler", "OK", "Error")
    } finally {
        Show-Progress $false
        $signEXEButton.Enabled = $true
    }
})

# Form Load
$form.Add_Load({
    try {
        Write-DebugInfo "DEBUG: Form Load gestartet"
        
        # Einstellungen aus Registry laden
        Load-SettingsFromRegistry
        
        Write-DebugInfo "DEBUG: Settings geladen, pruefe DefaultFolder"
        
        # Sicherstellen, dass settings Hashtable existiert
        if (-not $script:settings) {
            $script:settings = $defaultSettings.Clone()
            Write-DebugInfo "DEBUG: settings war NULL, initialisiert mit DefaultSettings"
        }
        
        # Sicherstellen, dass DefaultFolder immer einen gueltigen Wert hat
        if (-not $script:settings.DefaultFolder -or $script:settings.DefaultFolder.Trim() -eq "") {
            $script:settings.DefaultFolder = [Environment]::GetFolderPath("MyDocuments")
            Write-DebugInfo "DEBUG: DefaultFolder war leer, gesetzt auf: $($script:settings.DefaultFolder)"
        }
        
        Write-DebugInfo "DEBUG: DefaultFolder validiert: $($script:settings.DefaultFolder)"
        
        # Sicherstellen, dass currentDirectory initialisiert ist
        if (-not $script:currentDirectory -or $script:currentDirectory.Trim() -eq "") {
            $script:currentDirectory = $script:settings.DefaultFolder
            Write-DebugInfo "DEBUG: currentDirectory war leer, gesetzt auf: $($script:currentDirectory)"
        }
        
        Write-DebugInfo "DEBUG: currentDirectory validiert: $($script:currentDirectory)"
        
        # Zertifikate und Verzeichnis beim Start laden
        Write-DebugInfo "DEBUG: Lade Zertifikate..."
        if ($refreshCertButton) {
            $refreshCertButton.PerformClick()
        } else {
            Write-DebugInfo "DEBUG: refreshCertButton ist NULL"
        }
        
        Write-DebugInfo "DEBUG: Lade Verzeichnis..."
        Load-Directory $script:settings.DefaultFolder
        
        # Button-Status initialisieren
        Write-DebugInfo "DEBUG: Initialisiere Button-Status..."
        Update-ButtonStates
        
        Write-DebugInfo "DEBUG: Form Load erfolgreich abgeschlossen"
        
        # EXE-Workflow Info anzeigen
        Update-Info "PhinIT CERTUM Code Signing & EXE Creation Tool`n`nEXE-Konvertierung - Die Losung fur PowerShell ExecutionPolicy:`n================================================================================`n`nEXE-Dateien umgehen PowerShell ExecutionPolicy komplett`nWindows behandelt signierte EXE-Dateien normal`nKeine 'Trust Publisher' Installation erforderlich`nPS2EXE Integration fur nahtlose Konvertierung`n`nEmpfohlener Workflow:`n1. PS1-Datei auswahlen (Doppelklick im Explorer)`n2. 'PS1 -> EXE' klicken (konvertiert zu ausfuhrbarer EXE)`n3. EXE-Datei auswahlen (Doppelklick im Explorer)`n4. 'EXE SIGNIEREN' klicken (digital signieren)`n`nNavigation: Doppelklick auf Ordner/Datei zur Auswahl`nSimplySign fur CERTUM Cloud-Integration verfugbar`nZertifikate werden automatisch beim Start geladen`n`nUnterstutzte Dateitypen: .ps1 (PowerShell-Skripte) und .exe (ausfuhrbare Dateien)`n`nBereit fur EXE-Workflow!"
    }
    catch {
        Write-DebugInfo "DEBUG: Fehler beim Form Load: $($_.Exception.Message)"
        Write-DebugInfo "DEBUG: StackTrace: $($_.Exception.StackTrace)"
        
        # Fallback bei Fehlern
        $script:settings = $defaultSettings.Clone()
        $script:currentDirectory = [Environment]::GetFolderPath("MyDocuments")
        
        try {
            Load-Directory $script:currentDirectory
            Update-ButtonStates
            Update-Info "Fehler beim Laden behoben - Fallback-Modus aktiviert`n`nFehlerdetails: $($_.Exception.Message)"
        }
        catch {
            Update-Info "Kritischer Fehler beim Laden - bitte kontaktieren Sie den Support`n`nFehler: $($_.Exception.Message)"
        }
    }
})

# Form anzeigen
Update-Status "Anwendung gestartet - Wahlen Sie eine PowerShell-Datei aus"
$form.ShowDialog() | Out-Null


# SIG # Begin signature block
# MIIRcAYJKoZIhvcNAQcCoIIRYTCCEV0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA/koouWrskVGly
# QuaxvKwCHW90mHIsDBj2jfmMAh9JEKCCDaowgga5MIIEoaADAgECAhEAmaOACiZV
# O2Wr3G6EprPqOTANBgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNV
# BAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBD
# ZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQg
# TmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjEL
# MAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEk
# MCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34
# vuTmflN4mSAfgLKTvggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSo
# McbvIKck6+hI1shsylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+
# njvMk2xqbNTIPsnWtw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHI
# ohzuC8KNqfcYf7Z4/iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRda
# CtVD2bSlbfsq7BiqljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837
# Q4QOZgYqVWQ4x6cM7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8m
# FSkcSkAExzd4prHwYjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI
# 5XDYO962WZayx7ACFf5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5bou
# fUnq1UiYPIAHlezf4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplK
# ShBSND36E/ENVv8urPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n8
# 58JSqPECAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10
# XUwA23ufoHTKsW73PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbR
# Og79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8E
# KTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsG
# AQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVt
# LmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0
# bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6
# Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCia
# U58Q7EP89DttyZqGYn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L9
# 4C9LGF3vjzzH8Jq3iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rG
# DxLUUAsO0eqeLNhLVsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f
# 8pXdd3x2mbJCKKtl2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0Cd
# Y9rNLqyA3ahe8WlxVWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B
# 7eMcZNhpn8zJ+6MTyE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloM
# U+vUBfSouCReZwSLo8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1t
# VxTpV2Na3PR7nxYVlPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2
# DaGgK2W1eEJfo2qyrBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNj
# t2ywzOr+tKyEVAotnyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCr
# W602NCmvO1nm+/80nLy5r0AZvCQxaQ4wggbpMIIE0aADAgECAhBiOsZKIV2oSfsf
# 25d4iu6HMA0GCSqGSIb3DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNp
# Z25pbmcgMjAyMSBDQTAeFw0yNTA3MzExMTM4MDhaFw0yNjA3MzExMTM4MDdaMIGO
# MQswCQYDVQQGEwJERTEbMBkGA1UECAwSQmFkZW4tV8O8cnR0ZW1iZXJnMRQwEgYD
# VQQHDAtCYWllcnNicm9ubjEeMBwGA1UECgwVT3BlbiBTb3VyY2UgRGV2ZWxvcGVy
# MSwwKgYDVQQDDCNPcGVuIFNvdXJjZSBEZXZlbG9wZXIsIEhlcHAgQW5kcmVhczCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOt2txKXx2UtfBNIw2kVihIA
# cgPkK3lp7np/qE0evLq2J/L5kx8m6dUY4WrrcXPSn1+W2/PVs/XBFV4fDfwczZnQ
# /hYzc8Ot5YxPKLx6hZxKC5v8LjNIZ3SRJvMbOpjzWoQH7MLIIj64n8mou+V0CMk8
# UElmU2d0nxBQyau1njQPCLvlfInu4tDndyp3P87V5bIdWw6MkZFhWDkILTYInYic
# YEkut5dN9hT02t/3rXu230DEZ6S1OQtm9loo8wzvwjRoVX3IxnfpCHGW8Z9ie9I9
# naMAOG2YpvpoUbLG3fL/B6JVNNR1mm/AYaqVMtAXJpRlqvbIZyepcG0YGB+kOQLd
# oQCWlIp3a14Z4kg6bU9CU1KNR4ueA+SqLNu0QGtgBAdTfqoWvyiaeyEogstBHglr
# Z39y/RW8OOa50pSleSRxSXiGW+yH+Ps5yrOopTQpKHy0kRincuJpYXgxGdGxxKHw
# uVJHKXL0nWScEku0C38pM9sYanIKncuF0Ed7RvyNqmPP5pt+p/0ZG+zLNu/Rce0L
# E5FjAIRtW2hFxmYMyohkafzyjCCCG0p2KFFT23CoUfXx59nCU+lyWx/iyDMV4sqr
# cvmZdPZF7lkaIb5B4PYPvFFE7enApz4Niycj1gPUFlx4qTcXHIbFLJDp0ry6MYel
# X+SiMHV7yDH/rnWXm5d3AgMBAAGjggF4MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1Ud
# HwQ2MDQwMqAwoC6GLGh0dHA6Ly9jY3NjYTIwMjEuY3JsLmNlcnR1bS5wbC9jY3Nj
# YTIwMjEuY3JsMHMGCCsGAQUFBwEBBGcwZTAsBggrBgEFBQcwAYYgaHR0cDovL2Nj
# c2NhMjAyMS5vY3NwLWNlcnR1bS5jb20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBv
# c2l0b3J5LmNlcnR1bS5wbC9jY3NjYTIwMjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA
# 23ufoHTKsW73PMAywHDNMB0GA1UdDgQWBBQYl6R41hwxInb9JVvqbCTp9ILCcTBL
# BgNVHSAERDBCMAgGBmeBDAEEATA2BgsqhGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIB
# FhlodHRwczovL3d3dy5jZXJ0dW0ucGwvQ1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAQ4guyo7zysB7MHMB
# OVKKY72rdY5hrlxPci8u1RgBZ9ZDGFzhnUM7iIivieAeAYLVxP922V3ag9sDVNR+
# mzCmu1pWCgZyBbNXykueKJwOfE8VdpmC/F7637i8a7Pyq6qPbcfvLSqiXtVrT4NX
# 4NIvODW3kIqf4nGwd0h31tuJVHLkdpGmT0q4TW0gAxnNoQ+lO8uNzCrtOBk+4e1/
# 3CZXSDnjR8SUsHrHdhnmqkAnYb40vf69dfDR148tToUj872yYeBUEGUsQUDgJ6HS
# kMVpLQz/Nb3xy9qkY33M7CBWKuBVwEcbGig/yj7CABhIrY1XwRddYQhEyozUS4mX
# NqXydAD6Ylt143qrECD2s3MDQBgP2sbRHdhVgzr9+n1iztXkPHpIlnnXPkZrt89E
# 5iGL+1PtjETrhTkr7nxjyMFjrbmJ8W/XglwopUTCGfopDFPlzaoFf5rH/v3uzS24
# yb6+dwQrvCwFA9Y9ZHy2ITJx7/Ll6AxWt7Lz9JCJ5xRyYeRUHs6ycB8EuMPAKyGp
# zdGtjWv2rkTXbkIYUjklFTpquXJBc/kO5L+Quu0a0uKn4ea16SkABy052XHQqd87
# cSJg3rGxsagi0IAfxGM608oupufSS/q9mpQPgkDuMJ8/zdre0st8OduAoG131W+X
# J7mm0gIuh2zNmSIet5RDoa8THmwxggMcMIIDGAIBATBqMFYxCzAJBgNVBAYTAlBM
# MSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0Nl
# cnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQQIQYjrGSiFdqEn7H9uXeIruhzANBglg
# hkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MC8GCSqGSIb3DQEJBDEiBCDl3pUhzZolC/5fSb5lzBPiII+f7ESI2wDIFoY3B9jx
# LzANBgkqhkiG9w0BAQEFAASCAgAxaXrcvCWCD7EDYBDZ2c7L/6tgHAf4u2BHt5Uy
# 2Oadx9Ma2ooji5Fx7uiKjEnEEbQfR+KeNqyPXewvrmj2SvoC62KwpQd7BVUp+1qo
# ZQZ7XuTRXjYcEViVljwoU7STS8cFRxhNeezspH+pTTzKlhxZ1jbk6uHulh/onL5t
# XGXc3XyiFytBbuE3MO6XQlOEcZr/k6+t1Y+b+O5CkkBjBa+Eplf/x/tpZyJrt3Vs
# Iko2SjLBgxRrjrxQ1K4+Hgm8/4YUU4sVT243KdlqOtukb69untklD6TaYkCdj8oQ
# Vd4ZD7XFnyyuqPtepSHkxF2NhtjihY44ooUeIgZ1c2pmcq0hm7dvabjn9bdDAu/G
# PZ8yWxo3u9DZ6PQzfKavjIKc+iIXhSpDEnUhup1/QgZ4XTsrzTTdQh9IS4sNdRfg
# RtMkPCIkflfpOr5KQdM+c/zpmnuANqRF1xAGK5kBKDAgBr/pcjyj4Xl37NbFQvwb
# ++at1CTOb98ZnfMyVFBmSJtaMk0HX5TE5rkduaZBr9Po9CL2V0CE8HEMKSLjwPd3
# 69F2LAKQzMVSTLXKb4D3eMGZN0ZrFf3w216TTOFbgssWpuuiw+5PS+f1h9G1OAbh
# J7SmrKGKwZvJ6kijq+8DIQ5UhNJjykTtHo25T9J9ZmbIRY6rLg2uLnjLqhnZLtX/
# nI1BDQ==
# SIG # End signature block
