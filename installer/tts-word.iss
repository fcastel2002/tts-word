; Instalador de TTS Word (Inno Setup 6). Se compila con build.ps1, que pasa /DAppVersion.
;
; Qué hace, en orden:
;  1. Copia el helper empaquetado (PyInstaller) a Program Files.
;  2. Crea una CA propia de esta PC y el certificado de https://localhost:8766, y confía en la CA.
;  3. Hace que el helper arranque con Windows (para todos los usuarios) y lo arranca ya.
;     Al arrancar, el helper registra el complemento en Word para ese usuario.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "TTS Word"
#define HelperExe "tts-helper.exe"
#define CaName "TTS Word - CA local"

[Setup]
AppId={{8C1E3A52-6D4B-4F0E-9B7A-3E2F1D5C9A61}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Francisco Castel
DefaultDirName={autopf}\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=Output
OutputBaseFilename=TTSWord-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupIconFile=tts-word.ico
UninstallDisplayIcon={app}\{#HelperExe}
; El helper se cierra en PrepareToInstall; no hace falta que Restart Manager pregunte.
CloseApplications=no

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Messages]
es.FinishedLabelNoIcons=Listo. Abrí Word (si ya estaba abierto, cerralo y volvé a abrirlo) y hacé clic en "Leer en voz alta", en la pestaña Inicio.

[Files]
Source: "build\dist\tts-helper\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#AppName}"; ValueData: """{app}\{#HelperExe}"""; Flags: uninsdeletevalue

[Run]
; Si es una reinstalación, la CA anterior se reemplaza por una nueva.
Filename: "{sys}\certutil.exe"; Parameters: "-delstore Root ""{#CaName}"""; Flags: runhidden waituntilterminated; StatusMsg: "Preparando la conexión segura..."
Filename: "{app}\{#HelperExe}"; Parameters: "--generate-cert"; Flags: waituntilterminated; StatusMsg: "Preparando la conexión segura..."
Filename: "{sys}\certutil.exe"; Parameters: "-addstore Root ""{commonappdata}\{#AppName}\ca.cer"""; Flags: runhidden waituntilterminated; StatusMsg: "Preparando la conexión segura..."
Filename: "{app}\{#HelperExe}"; Flags: nowait runasoriginaluser; StatusMsg: "Iniciando el servicio de voz..."

[UninstallRun]
Filename: "{app}\{#HelperExe}"; Parameters: "--unregister"; Flags: waituntilterminated; RunOnceId: "Unregister"
Filename: "{sys}\certutil.exe"; Parameters: "-delstore Root ""{#CaName}"""; Flags: runhidden waituntilterminated; RunOnceId: "DeleteCA"

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\{#AppName}"

[Code]
procedure StopHelper();
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#HelperExe}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  StopHelper();  // si no, los archivos están en uso al actualizar
  Result := '';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  // Después de --unregister (UninstallRun) y antes de borrar los archivos.
  if CurUninstallStep = usUninstall then
    StopHelper();
end;
