; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "DOF Creator"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "DW Hirst & Associates INC"
#define MyAppExeName "display.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{1F85AF7E-5CC4-4E6D-B7C1-9438D590F621}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project
OutputBaseFilename=DOF_setup_v1.0.0_x64
SetupIconFile=C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project\src\*"; DestDir: "{app}/src"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Alexander Hastings\OneDrive\Desktop\_test_project\icon.ico"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

