# Build Script for RAM Optimizer
$ErrorActionPreference = 'Stop'

# Configuration
$config = @{
    AppName = "RAM Optimizer"
    Version = "3.2.0"
    OutputDir = ".\build"
    DistDir = ".\dist"
}

# Create necessary directories
Write-Host "Creating build directories..." -ForegroundColor Cyan
$null = New-Item -ItemType Directory -Force -Path $config.OutputDir
$null = New-Item -ItemType Directory -Force -Path $config.DistDir

# Function to check and install required modules
function Install-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing $ModuleName module..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
    }
    Import-Module $ModuleName -Force
}

# Install required modules
Install-RequiredModule "PS2EXE"

# Build the executable
Write-Host "Building RAM Optimizer executable..." -ForegroundColor Cyan
try {
    $exeParams = @{
        InputFile = "RAMOptimizer.ps1"
        OutputFile = "$($config.OutputDir)\RAMOptimizer.exe"
        NoConsole = $true  # Hide console window since we're using system tray
        RequireAdmin = $true
        Title = $config.AppName
        Version = $config.Version
        WindowStyle = "Hidden"
        IconFile = "$PSScriptRoot\ram_icon.ico"
    }
    
    Invoke-PS2EXE @exeParams
    
    Write-Host "Executable created successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error creating executable: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Create installer script
$installerScript = @'
; RAM Optimizer Installer Script
#define MyAppName "RAM Optimizer"
#define MyAppVersion "3.2.0"
#define MyAppPublisher "System Utilities"
#define MyAppExeName "RAMOptimizer.exe"

[Setup]
AppId={{B89F4599-9C9B-4E8F-BA73-D1F5A16E8C9A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=dist
OutputBaseFilename=RAMOptimizer_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "startupicon"; Description: "Start with Windows"; GroupDescription: "Windows Startup"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "build\RAMOptimizer.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent runascurrentuser
'@

# Save installer script
$installerScript | Out-File -FilePath "$($config.OutputDir)\installer.iss" -Encoding UTF8

Write-Host "`nBuild completed successfully!" -ForegroundColor Green
Write-Host "Executable location: $($config.OutputDir)\RAMOptimizer.exe" -ForegroundColor Cyan
Write-Host "Installer script: $($config.OutputDir)\installer.iss" -ForegroundColor Cyan

Write-Host "`nTo create the installer:" -ForegroundColor Yellow
Write-Host "1. Download and install Inno Setup from: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
Write-Host "2. Open the installer.iss file with Inno Setup" -ForegroundColor Yellow
Write-Host "3. Click Build > Compile" -ForegroundColor Yellow 