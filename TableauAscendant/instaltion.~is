; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=Tableau d'ascendance
AppVersion=1.5
DefaultDirName={pf}\TableauAscendance
DefaultGroupName=TableauAscendance
UninstallDisplayIcon={app}\TableauAscendant.exe
Compression=lzma2
SolidCompression=yes
OutputDir=userdocs:Inno Setup Examples Output

[Files]
Source: "C:\Users\dapam\OneDrive\Csharp\source\repos\TableauAscendant\TableauAscendant\bin\Release\TableauAscendant.exe"; DestDir: "{app}"
Source: "litmoi.html"; DestDir: "{app}"; Flags: isreadme
Source: "C:\Users\dapam\OneDrive\Csharp\source\repos\TableauAscendant\packages\PDFsharp.1.32.3057.0\lib\net20\PdfSharp.dll"; DestDir: "{app}";

[Icons]
Name: "{group}\TableauAscendant"; Filename: "{app}\TableauAscendant.exe"
