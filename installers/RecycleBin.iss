[Code]
{ This cleanup runs without the .NET runtime, including after an abnormal app exit. }
procedure RestoreRecycleBinOpenInView(Root: Integer; const Page: String);
var
  Shell, Verb, Command, Owner, Recorded, Actual, Delegate, Previous, CurrentDefault: String;
  HadDefault, PreviousKind, HadPage, HadShell: Cardinal;
begin
  Shell := Page + '\shell';
  Verb := Shell + '\open';
  Command := Verb + '\command';
  if not RegQueryStringValue(Root, Verb, 'WinTab.Owner', Owner) then exit;
  if Owner <> 'WinTab.RecycleBinOpen.v1' then exit;
  if not RegQueryStringValue(Root, Verb, 'WinTab.Command', Recorded) then exit;
  if RegQueryStringValue(Root, Command, '', Actual) then
  begin
    if Actual <> Recorded then exit;
    if not RegQueryStringValue(Root, Command, 'DelegateExecute', Delegate) then exit;
    if Delegate <> '' then exit;
  end;
  HadPage := 1;
  HadShell := 1;
  RegQueryDWordValue(Root, Verb, 'WinTab.HadPage', HadPage);
  RegQueryDWordValue(Root, Verb, 'WinTab.HadShell', HadShell);
  if RegQueryStringValue(Root, Shell, '', CurrentDefault) and (CurrentDefault = 'open') then
  begin
    if not RegQueryDWordValue(Root, Verb, 'WinTab.HadDefault', HadDefault) then exit;
    if HadDefault <> 0 then
    begin
      if not RegQueryStringValue(Root, Verb, 'WinTab.Default', Previous) then exit;
      if not RegQueryDWordValue(Root, Verb, 'WinTab.DefaultKind', PreviousKind) then exit;
      if PreviousKind = 2 then
      begin
        if not RegWriteExpandStringValue(Root, Shell, '', Previous) then exit;
      end
      else if not RegWriteStringValue(Root, Shell, '', Previous) then exit;
    end
    else if not RegDeleteValue(Root, Shell, '') then exit;
  end;
  RegDeleteValue(Root, Command, '');
  RegDeleteValue(Root, Command, 'DelegateExecute');
  RegDeleteKeyIfEmpty(Root, Command);
  RegDeleteValue(Root, Verb, 'WinTab.Owner');
  RegDeleteValue(Root, Verb, 'WinTab.Command');
  RegDeleteValue(Root, Verb, 'WinTab.HadDefault');
  RegDeleteValue(Root, Verb, 'WinTab.Default');
  RegDeleteValue(Root, Verb, 'WinTab.DefaultKind');
  RegDeleteValue(Root, Verb, 'WinTab.HadPage');
  RegDeleteValue(Root, Verb, 'WinTab.HadShell');
  RegDeleteKeyIfEmpty(Root, Verb);
  if HadShell = 0 then RegDeleteKeyIfEmpty(Root, Shell);
  if HadPage = 0 then RegDeleteKeyIfEmpty(Root, Page);
end;
