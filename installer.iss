; Inno Setup script for Docx_Translator (Windows)

[Setup]
AppName=Docx Translator
AppVersion=1.0.0
DefaultDirName={pf}\Docx Translator
DefaultGroupName=Docx Translator
OutputBaseFilename=Docx_Translator_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "dist\Docx_Translator.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "settings.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "Project_env.env"; DestDir: "{app}"; Flags: ignoreversion
Source: "Files\quality_rules.json"; DestDir: "{app}\Files"; Flags: ignoreversion
Source: "Files\protected_words.json"; DestDir: "{app}\Files"; Flags: ignoreversion
Source: "translation_memory.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Docx Translator"; Filename: "{app}\Docx_Translator.exe"
Name: "{commondesktop}\Docx Translator"; Filename: "{app}\Docx_Translator.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; Flags: unchecked

[Run]
Filename: "{app}\Docx_Translator.exe"; Description: "Launch Docx Translator"; Flags: nowait postinstall skipifsilent
