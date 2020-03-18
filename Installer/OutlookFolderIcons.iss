[Setup]
LicenseFile=MIT.txt
VersionInfoDescription=Outlook Folder Icons Add-in Setup
VersionInfoCompany=T800 Productions
VersionInfoCopyright=T800 Productions © 2019
VersionInfoVersion=1.0.7163.22045

AppCopyright=T800 Productions © 2019
AppName=Outlook Folder Icons Add-in

; TODO sync with dll fileversion
AppVersion=1.0.0.0
AppVerName=Outlook Folder Icons Add-in 1.0
VersionInfoProductVersion=1.0.0.0

AppPublisher=T800 Productions
AppPublisherURL=https://t800g.github.io
UninstallDisplayName=Outlook Folder Icons Add-in
DisableWelcomePage=no
DefaultGroupName=T800 Productions\Outlook Folder Icons
DefaultDirName={localappdata}\Outlook Folder Icons
DisableDirPage=true
DirExistsWarning=no
WizardImageFile=compiler:wizmodernimage-is.bmp
WizardSmallImageFile=compiler:wizmodernsmallimage-is.bmp
OutputDir=.
OutputBaseFilename=OutlookFolderIconsSetup
AlwaysUsePersonalGroup=true
AlwaysShowDirOnReadyPage=true
AlwaysShowGroupOnReadyPage=true

SolidCompression=true
AllowCancelDuringInstall=false

;always per-user install
PrivilegesRequired=lowest
MinVersion=0,5.0
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: ..\OutlookFolderIcons\bin\Release\OutlookFolderIcons.dll; DestDir: {app}; Flags: restartreplace uninsrestartdelete ignoreversion
Source: ..\Help\*.*; DestDir: {app}\Help; Flags: restartreplace uninsrestartdelete ignoreversion recursesubdirs
Source: ..\Icons\*.*; DestDir: {userappdata}\Outlook Folder Icons\Library; Flags: uninsneveruninstall onlyifdoesntexist recursesubdirs


[Registry]
Root: HKCU; Subkey: Software\Microsoft\Office\Outlook\Addins\OutlookFolderIcons.Connect; Flags: uninsdeletekey; ValueType: string; ValueName: FriendlyName; ValueData: Outlook Folder Icons Add-in
Root: HKCU; Subkey: Software\Microsoft\Office\Outlook\Addins\OutlookFolderIcons.Connect; Flags: uninsdeletekey; ValueType: string; ValueName: Description; ValueData: Outlook Folder Icons Add-in
Root: HKCU; Subkey: Software\Microsoft\Office\Outlook\Addins\OutlookFolderIcons.Connect; Flags: uninsdeletekey; ValueType: dword; ValueName: LoadBehavior; ValueData: 00000003
Root: HKCU; Subkey: Software\Microsoft\Office\Outlook\Addins\OutlookFolderIcons.Connect; Flags: uninsdeletekey; ValueType: dword; ValueName: CommandLineSafe; ValueData: 00000000

Root: HKCU; Subkey: Software\Classes\OutlookFolderIcons.Connect\CLSID; Flags: uninsdeletekey; ValueType: string; ValueData: {{AC92B228-C86E-4500-B1DE-D6E78D4CD094}

Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}; Flags: uninsdeletekey; ValueType: string; ValueData: OutlookFolderIcons.Connect
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32; Flags: uninsdeletekey; ValueType: string; ValueData: mscoree.dll
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32; Flags: uninsdeletekey; ValueType: string; ValueName: ThreadingModel; ValueData: Both
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Flags: uninsdeletekey; ValueType: string; ValueName: Assembly; ValueData: OutlookFolderIcons, Version=1.0.0.0, Culture=neutral, PublicKeyToken=c2bccb689fe74e7c
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Flags: uninsdeletekey; ValueType: string; ValueName: Class; ValueData: OutlookFolderIcons.Connect
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Flags: uninsdeletekey; ValueType: string; ValueName: Codebase; ValueData: file:///{code:GetAppURI}/OutlookFolderIcons.DLL
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Flags: uninsdeletekey; ValueType: string; ValueName: RuntimeVersion; ValueData: v4.0.30319.42000
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\ProgId; Flags: uninsdeletekey; ValueType: string; ValueData: OutlookFolderIcons.Connect
Root: HKCU; Subkey: Software\Classes\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\Version; Flags: uninsdeletekey; ValueType: string; ValueData: 1.0.0.0

Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueData: OutlookFolderIcons.Connect
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueData: mscoree.dll
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueName: ThreadingModel; ValueData: Both
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueName: Assembly; ValueData: OutlookFolderIcons, Version=1.0.0.0, Culture=neutral, PublicKeyToken=c2bccb689fe74e7c
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueName: Class; ValueData: OutlookFolderIcons.Connect
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueName: Codebase; ValueData: file:///{code:GetAppURI}/OutlookFolderIcons.DLL
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\InProcServer32\1.0.0.0; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueName: RuntimeVersion; ValueData: v4.0.30319.42000
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\ProgId; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueData: OutlookFolderIcons.Connect
Root: HKCU64; Subkey: Software\Classes\Wow6432Node\CLSID\{{AC92B228-C86E-4500-B1DE-D6E78D4CD094}\Version; Check: IsWin64; Flags: uninsdeletekey; ValueType: string; ValueData: 1.0.0.0

[Icons]
Name: {group}\Folder Icons Help; Filename: {app}\Help\index.html; Comment: Open help file

[Run]
Filename: {app}\Help\index.html; Description: Open user manual; Flags: postinstall shellexec skipifsilent

[Code]
function InitializeSetup(): Boolean;
Begin
	Result:=false;
	try
	  ExpandConstant('{dotnet40}')
	  Result:=true;
	except
	  MsgBox('Outlook Folder Icons Add-in requires .NET Framework 4.0 or higher', mbCriticalError, mb_Ok);
	end;
end;

function GetAppURI(Param: String): string;
begin
  Result:=ExpandConstant('{app}');
  StringChangeEx(Result, '\', '/', True);
end;


[_ISToolPreCompile]
Name: autover.exe; Parameters: OutlookFolderIcons.iss /p
