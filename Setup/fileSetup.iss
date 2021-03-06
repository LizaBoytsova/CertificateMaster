; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define Name "������ �������������"
#define ExeName "������ �������������.exe"
#define Publisher "BoytsovaProject"
#define SourceExeName "CertificateMaster.exe"
#define SourceExeFolder  "..\CertificateMaster\bin\Release\"
#define DocFolder  "..\CertificateMaster\Doc\"
#define TemplateDir "�������"

#define Version GetFileVersion(SourceExeFolder + "\" + SourceExeName)


[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{Guid("5C3F0E74-A039-4BE8-BB07-22FC710F8CA3}
AppName={#Name}
AppVersion={#Version}
AppVerName={#Name} ���. {#Version}
AppPublisher={#Publisher}
DefaultDirName={pf}\{#Name}
DefaultGroupName={#Name}
UninstallDisplayIcon = {app}\{#ExeName}
OutputDir= .
OutputBaseFilename=Setup {#Name} ���. {#Version} 
Compression=lzma
SolidCompression=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: {#SourceExeFolder}{#SourceExeName}; DestDir: {app}; Flags: ignoreversion; DestName:  {#ExeName};
Source: {#DocFolder}\Certificate.doc; DestDir: {localappdata}\{#Name}\{#TemplateDir}; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

; ��������� NET 4.5.1
Source: "Net 4.5\dotnetfx45_full_x86_x64.exe";     DestDir: {app}; Components: dotnet; Check: isNET45_Installed; AfterInstall: InstallNET_45; Flags: replacesameversion nocompression

[Icons]
Name: "{group}\{#Name}"; Filename: "{app}\{#Name}"
Name: {group}\{cm:UninstallProgram,{#NAME}}; Filename: {uninstallexe};
Name: "{commondesktop}\{#Name}"; Filename: "{app}\{#Name}"; Tasks: desktopicon

[Components]
;�������� ������ ���������
Name: "dotnet"; Description: '.NET 4.5'; Types: full compact custom;
;Flags: fixed; 

[Run]
Filename: "{app}\{#Name}"; Description: "{cm:LaunchProgram,{#StringChange(Name, '&', '&&')}}"; Flags: nowait postinstall skipifsilent


[Code]
// ��������� .NET 4.5
procedure InstallNET_45();
var
  ResultCode: Integer;
begin
   Log('��������� .NET 4.5');
   Exec(ExpandConstant('{app}\dotnetfx45_full_x86_x64.exe'), ExpandConstant('/passive /norestart'), '', SW_SHOW, ewWaitUntilTerminated, ResultCode);   

   Log('��������� ������� ����������� .NET: ' + IntToStr(ResultCode));
   Log('�������� ������: ' + SysErrorMessage(ResultCode));
end;


function IsDotNetDetected(version: string; service: cardinal): Boolean;

// Indicates whether the specified version and service pack of //
// version -- Specify one of these strings for the required .NET
// 'v1.1.4322' .NET Framework 1.1
// 'v2.0.50727' .NET Framework 2.0
// 'v3.0' .NET Framework 3.0
// 'v3.5' .NET Framework 3.5
// 'v4\Client' .NET Framework 4.0 Client Profile
// 'v4\Full' .NET Framework 4.0 Full Installation
// 'v4.5' .NET Framework 4.5
//
// service -- Specify any non-negative integer for the required // 0 No service packs required
// 1, 2, etc. Service pack 1, 2, etc. required
var
key: string;
install, release, serviceCount: cardinal;
check45, check451, success: boolean;
begin
	// .NET 4.5 installs as update to .NET 4.0 Full
	if version = 'v4.5' then 
	begin
		version := 'v4\Full';
		check45 := true;
	end 
	else
		check45 := false;

  if version = 'v4.5.1' then 
	begin
		version := 'v4\Full';
		check451 := true;
	end 
	else
		check451 := false;
		
		// installation key group for all .NET versions
		key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\' + version;
		// .NET 3.0 uses value InstallSuccess in subkey Setup
		if Pos('v3.0', version) = 1 then 
		begin
			success := RegQueryDWordValue(HKLM, key + '\Setup', 'InstallSuccess', install);
		end 
		else 
		begin
			success := RegQueryDWordValue(HKLM, key, 'Install', install);
		end;
		
		// .NET 4.0/4.5 uses value Servicing instead of SP
		if Pos('v4', version) = 1 then 
		begin
			success := success and RegQueryDWordValue(HKLM, key, 'Servicing', serviceCount);
    end
    else
    begin
			success := success and RegQueryDWordValue(HKLM, key, 'SP', serviceCount);
    end;

		// .NET 4.5 uses additional value Release
		if check45 then 
		begin
			success := success and RegQueryDWordValue(HKLM, key, 'Release', release); 
      success := success and (release >= 378389);
		end;

    if check451 then 
		begin
			success := success and RegQueryDWordValue(HKLM, key, 'Release', release); 
      success := success and (release >= 378758);
		end;
		
		result := success and (install = 1) and (serviceCount >= service);
end;



function isNET45_Installed(): Boolean;
begin
    result := not IsDotNetDetected('v4.5', 0);
end;


