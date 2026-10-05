[Code]
var
  UpdateInstall: Boolean;

procedure DetectUpdateInstall(const AppPath: String);
begin
  { Check before files are copied, including manual upgrades from older versions. }
  UpdateInstall := (ExpandConstant('{param:UPDATE|0}') = '1') or FileExists(AppPath);
end;

function GetAppLaunchParameters(Param: String): String;
begin
  if UpdateInstall then
    Result := '--background'
  else
    Result := '';
end;
