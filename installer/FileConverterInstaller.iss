[Setup]
AppName=FileConverter
AppVersion=1.0
DefaultDirName={commonpf}\FileConverter
DefaultGroupName=FileConverter
OutputBaseFilename=FileConverterInstaller
Compression=lzma
SolidCompression=yes

[Files]
Source: "C:\Users\pavan\Desktop\FileConverter\app\dist\launcher.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\FileConverter"; Filename: "{app}\launcher.exe"
Name: "{commondesktop}\FileConverter"; Filename: "{app}\launcher.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Registry]
; Register context menu for all files (*)
Root: HKCR; Subkey: "*\shell\ConvertToPDF"; ValueType: string; ValueName: ""; ValueData: "Convert to PDF"
Root: HKCR; Subkey: "*\shell\ConvertToPDF\command"; ValueType: string; ValueName: ""; ValueData: """{app}\launcher.exe"" ""%1"" ""topdf"""

Root: HKCR; Subkey: "*\shell\ConvertToJPG"; ValueType: string; ValueName: ""; ValueData: "Convert to JPG"
Root: HKCR; Subkey: "*\shell\ConvertToJPG\command"; ValueType: string; ValueName: ""; ValueData: """{app}\launcher.exe"" ""%1"" ""tojpg"""

Root: HKCR; Subkey: "*\shell\ConvertToDOCX"; ValueType: string; ValueName: ""; ValueData: "Convert to DOCX"
Root: HKCR; Subkey: "*\shell\ConvertToDOCX\command"; ValueType: string; ValueName: ""; ValueData: """{app}\launcher.exe"" ""%1"" ""todocx"""

Root: HKCR; Subkey: "*\shell\ConvertToPPTX"; ValueType: string; ValueName: ""; ValueData: "Convert to PPTX"
Root: HKCR; Subkey: "*\shell\ConvertToPPTX\command"; ValueType: string; ValueName: ""; ValueData: """{app}\launcher.exe"" ""%1"" ""topptx"""

Root: HKCR; Subkey: "*\shell\CompressFile"; ValueType: string; ValueName: ""; ValueData: "Compress"
Root: HKCR; Subkey: "*\shell\CompressFile\command"; ValueType: string; ValueName: ""; ValueData: """{app}\launcher.exe"" ""%1"" ""compress"""

[Run]
Filename: "{app}\launcher.exe"; Description: "Launch FileConverter"; Flags: nowait postinstall skipifsilent
