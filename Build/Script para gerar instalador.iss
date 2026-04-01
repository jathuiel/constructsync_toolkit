[Setup]
AppName=Set Atributes Toolkit
AppVersion=1.0
OutputBaseFilename=SetAtributesToolkit_Installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "SetAtributesToolkit.dll"; DestDir: "{code:GetSelectedDir}"; Flags: ignoreversion
Source: "SetAtributesToolkit.addin"; DestDir: "{code:GetSelectedDir}"; Flags: ignoreversion
Source: "SetAtributesToolkit.xaml"; DestDir: "{code:GetSelectedDir}"; Flags: ignoreversion
Source: "Resources\*.png"; DestDir: "{code:GetSelectedDir}\Resources"; Flags: ignoreversion

[Code]
var
  NavisPage: TInputOptionWizardPage;
  NavisPaths: array of string;

procedure InitializeWizard;
var
  BasePath: string;
  FindRec: TFindRec;
  Index: Integer;
begin
  BasePath := ExpandConstant('{commonappdata}\Autodesk\');

  NavisPage := CreateInputOptionPage(
    wpSelectDir,
    'Selecione o Navisworks',
    'Escolha onde instalar o plugin',
    'Selecione a(s) versão(ões) do Navisworks disponíveis:',
    True,
    False
  );

  Index := 0;

  if FindFirst(BasePath + 'Navisworks*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          if (Pos('Manage', FindRec.Name) > 0) or (Pos('Simulate', FindRec.Name) > 0) then
          begin
            SetArrayLength(NavisPaths, Index + 1);
            NavisPaths[Index] := BasePath + FindRec.Name + '\Plugins\SetAtributesToolkit';

            NavisPage.Add(FindRec.Name);
            Index := Index + 1;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function GetSelectedDir(Value: string): string;
var
  i: Integer;
begin
  for i := 0 to GetArrayLength(NavisPaths) - 1 do
  begin
    if NavisPage.Values[i] then
    begin
      Result := NavisPaths[i];
      Exit;
    end;
  end;

  Result := ExpandConstant('{commonappdata}\Autodesk\Navisworks Manage 2023\Plugins\SetAtributesToolkit');
end;