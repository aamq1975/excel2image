; Inno Setup script for Excel Image Extractor
; -------------------------------------------
; Open this file in Inno Setup Compiler (after installing it) and hit
; Build > Compile, or run it from the command line with ISCC.exe.
; See README.md, Step 5, for full instructions.

#define MyAppName "Excel Image Extractor"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Muhammad Aamir"
#define MyAppExeName "ExcelImageExtractor.exe"

[Setup]
AppId={{B3B6B4B0-2E3A-4C1F-9E2A-EXCELIMGEXT01}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; The installer itself will be created inside an "Output" folder next
; to this script.
OutputBaseFilename=ExcelImageExtractor_Setup
Compression=lzma2/max
SolidCompression=yes
SetupIconFile=icon.ico
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
; This expects the PyInstaller build to already exist at dist\ExcelImageExtractor.exe
; (built with --onefile), relative to this .iss file's location.
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
