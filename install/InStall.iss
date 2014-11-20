; 脚本由 Inno Setup 脚本向导 生成！
; 有关创建 Inno Setup 脚本文件的详细资料请查阅帮助文档！

[Setup]
; 注意: AppId 值用于唯一识别该应用程序。
; 禁止对其他应用程序的安装器使用相同的 AppId 值！
; (若要生成一个新的 GUID，请选择“工具 | 生成 GUID”。)
AppId={{13779961-0058-4591-A51D-15FB410FAE3A}}
AppName=AgentOCX
AppVersion=2.0.0.0
AppCopyright=Copyright (C)qiwei Inc.
AppComments=qiwei
VersionInfoVersion=2.0.0.0
VersionInfoDescription=AgentOCX安装包
AppPublisher=qiwei
DefaultDirName={pf32}\CTI\AgentOCX
DefaultGroupName=CTI\AgentOCX
AllowNoIcons=yes
OutputDir=.\
OutputBaseFilename=AgentOCXSetup
Compression=lzma2/ultra
SolidCompression=yes
PrivilegesRequired=admin

ArchitecturesInstallIn64BitMode=x64 ia64

[Languages]
Name: "chinese"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
;Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
;Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: ".\DLL\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\msvcrt.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\mfc42.dll"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\Msvbvm50.dll"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\LiteZip.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\LiteUnzip.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\vbaliml6.ocx"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\vbalgrid.ocx"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\SSubTmr6.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\ssubtmr.dll"; DestDir: "{app}\DLL";  Flags: 32bit regserver
Source: ".\DLL\SmartUI.ocx"; DestDir: "{app}\DLL";  Flags: 32bit regserverSource: ".\DLL\SmartUI.oca"; DestDir: "{app}\DLL";  Flags: 32bit  Source: ".\DLL\Shdocvw.dll"; DestDir: "{app}\DLL";  Flags: 32bit 
Source: ".\DLL\scrrun.dll"; DestDir: "{sys}";   Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall  Source: ".\DLL\RecordListnerOCX.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\prjIVRTree.ocx"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\prjAgentInterpretor.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\pophandler.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\NetMeetingCOM.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\Mswinsck.ocx"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\msscript.ocx"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\Mscomct2.ocx"; DestDir: "{sys}"; Flags: 32bit regserver onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\cnewmenu.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\AutoItX.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\AgentOCX.dll"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\Actbar2.ocx"; DestDir: "{app}\DLL"; Flags: 32bit regserver
Source: ".\DLL\G723ToWave.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\djcvt.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist 
Source: ".\DLL\DJ_TIF.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist
Source: ".\tools\ClientWindow.exe"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\DLL\atl.dll"; DestDir: "{sys}"; Flags: 32bit onlyifdoesntexist uninsneveruninstall
Source: ".\Uniland\ItemString.txt"; DestDir: "{app}\Uniland";
Source: ".\Uniland\Prj_Uniland_AgentOCX.ocx"; DestDir: "{app}\Uniland"; Flags: 32bit regserver
Source: ".\Uniland\Project1.exe"; DestDir: "{app}\Uniland";
Source: ".\Uniland\Yeling.wav"; DestDir: "{app}\Uniland";
Source: ".\Sample\VB\*"; DestDir:"{app}\Sample\VB";Source: ".\Sample\VB\GingeAPI.dll"; DestDir:"{app}\Sample\VB"; Flags: 32bit regserver
Source: ".\Reg\regset.ini"; DestDir: "{app}\Reg"; 
Source: ".\Reg\reg.bat"; DestDir: "{app}\Reg"; 
Source: ".\Doc\*"; DestDir: "{app}\Doc"; 
Source: ".\*.vox"; DestDir: "{app}";
Source: ".\Bin\*"; DestDir: "{app}\Bin";

; 注意: 不要在任何共享系统文件上使用“Flags: ignoreversion”
                                            
[Icons]
Name: "{group}\Project1"; Filename: "{app}\Uniland\Project1.exe"          
Name: "{group}\SoftPhone使用手册"; Filename: "{app}\Doc\SoftPhone使用手册.pdf"
Name: "{group}\Sample\VB\Project1"; Filename: "{app}\Sample\VB\Project1.exe"
;Name: "{group}\{cm:UninstallProgram,AgentOCX}"; Filename: "{uninstallexe}"
;Name: "{commondesktop}\CloopenClientOCX"; Filename: "{app}\ClientOCX.htm"; Tasks: desktopicon
;Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\CloopenClientOCX"; Filename: "{app}\ClientOCX.htm"; Tasks: quicklaunchicon

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "AgentRunPath"; ValueData:"{app}\Bin\AgentRun.abs"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "BinPath"; ValueData:"{app}\Bin"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "ConnectionString"; ValueData:"Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=TF_CMS;Data Source=192.168.1.6"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "StatePath"; ValueData:"{app}\Bin\State.xml"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "DeviceName"; ValueData:"User"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword;  ValueName: "SetDeviceName"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "AutoNodeID"; ValueData:"E188FDDF884145CFAA133075BDE78A44"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "AutoSbsPath"; ValueData:"C:\123.sbs"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "OffHook_Always"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "EnableEmailFunc"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "EnableTextChatFunc"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "EnableTextChatNotify"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "POP3IPAddress"; ValueData:"192.168.1.6"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "POP3Port"; ValueData:"111"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: "TextChatServerIP"; ValueData:"192.168.1.2"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "NotifyAcwDur"; ValueData:"0"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "NotifyAcwDur_Rep"; ValueData:"1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: dword; ValueName: "NotifyDur"; ValueData:"3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Samwoo\AA\PBXSoftPhone"; ValueType: string; ValueName: ""; ValueData:""; Flags: uninsdeletekey

[UninstallDelete]
Type:files; Name:"{app}\*.log";
Type:filesandordirs;Name:{app};
Type:dirifempty;Name:"{pf32}\CTI"
Type:dirifempty;Name:{group};

