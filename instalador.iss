#ifndef Version
  #define Version "1.0.0"
#endif
#define Icono "assets\logo.ico"

[Setup]
AppId={{B370BDA2-4B29-4184-8DC8-CCED3B817D12}
AppName=Extractos PDF
AppVersion={#Version}
AppPublisher=Asesores Contables & Tributarios
DefaultDirName={autopf}\ExtractosPDF
DefaultGroupName=Extractos PDF
UninstallDisplayIcon={app}\ExtractosPDF.exe
OutputDir=instaladores
OutputBaseFilename=ExtractosPDF_Setup_{#Version}
SetupIconFile={#Icono}
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
CloseApplications=yes
RestartApplications=no
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "dist\app_gui.dist\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[InstallDelete]
Type: filesandordirs; Name: "{app}\*"

[Icons]
Name: "{group}\Extractos PDF"; Filename: "{app}\ExtractosPDF.exe"
Name: "{group}\Desinstalar Extractos PDF"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Extractos PDF"; Filename: "{app}\ExtractosPDF.exe"; Tasks: escritorio

[Run]
Filename: "{app}\ExtractosPDF.exe"; Description: "Abrir Extractos PDF"; Flags: nowait postinstall skipifsilent
Filename: "{app}\ExtractosPDF.exe"; Flags: nowait runasoriginaluser; Check: WizardSilent
