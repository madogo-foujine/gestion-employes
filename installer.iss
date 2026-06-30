; Inno Setup script — Gestion des Employes & Paie
; Build first:  pyinstaller --onefile --windowed --name GestionEmployes --icon icon.ico employee_manager.py
; Then compile: ISCC.exe installer.iss   (output: installer\GestionEmployes_Setup.exe)

[Setup]
AppId={{B7E4A2C1-9F3D-4E58-A1B2-EMP-PAIE-MAROC}}
AppName=Gestion des Employes
AppVersion=1.1
AppPublisher=Ma Societe
DefaultDirName={autopf}\GestionEmployes
DefaultGroupName=Gestion des Employes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=installer
OutputBaseFilename=GestionEmployes_Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\GestionEmployes.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Creer un raccourci sur le bureau"; GroupDescription: "Raccourcis :"

[Files]
Source: "dist\GestionEmployes.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gestion des Employes"; Filename: "{app}\GestionEmployes.exe"
Name: "{group}\Desinstaller Gestion des Employes"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Gestion des Employes"; Filename: "{app}\GestionEmployes.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\GestionEmployes.exe"; Description: "Lancer l'application maintenant"; Flags: nowait postinstall skipifsilent
