[Setup]
; Información básica de la aplicación
AppId={{EASYDOC SUITE-Tony2908-2026}}
AppName=EASYDOC SUITE
AppVersion=1.0
AppPublisher=Tony2908
DefaultDirName={autopf}\EASYDOC SUITE
DefaultGroupName=EASYDOC SUITE
AllowNoIcons=yes
; Configuramos el ícono del instalador
SetupIconFile=icono.ico
; Nombre del archivo final
OutputBaseFilename=EASYDOC SUITE
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Tomamos el ejecutable de la carpeta dist generada por tu Compilador.bat[cite: 3]
Source: "dist\EASYDOC SUITE.exe"; DestDir: "{app}"; Flags: ignoreversion
; Incluimos el ícono para los accesos directos
Source: "icono.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Acceso directo en el Menú Inicio
Name: "{group}\EASYDOC SUITE"; Filename: "{app}\EASYDOC SUITE.exe"; IconFilename: "{app}\icono.ico"
; Acceso directo opcional en el Escritorio
Name: "{autodesktop}\EASYDOC SUITE"; Filename: "{app}\EASYDOC SUITE.exe"; IconFilename: "{app}\icono.ico"; Tasks: desktopicon

[Run]
; Opción para ejecutar la aplicación al terminar la instalación
Filename: "{app}\EASYDOC SUITE.exe"; Description: "{cm:LaunchProgram,EASYDOC SUITE}"; Flags: nowait postinstall skipifsilent