unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus, ComCtrls;

type
  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    Start1: TMenuItem;
    Stop1: TMenuItem;
    Timer1: TTimer;
    Pause1: TMenuItem;
    Continue1: TMenuItem;
    ListView1: TListView;
    Refresh1: TMenuItem;
    Memo1: TMemo;
    Recreate1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ChgEtatService(Sender: TObject);
    procedure Rafraichir(Sender: TObject);
    procedure Recreate1Click(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    SC : THandle;
  end;

var
  Form1: TForm1;

implementation  uses WinSvc;

{$R *.dfm}

const SrvAccess = SERVICE_INTERROGATE or SERVICE_PAUSE_CONTINUE or
                  SERVICE_START or SERVICE_STOP or SERVICE_QUERY_STATUS;

function CurrState(State : dword) : string;
Begin
  Case State Of
    SERVICE_STOPPED: Result := 'STOP';
    SERVICE_START_PENDING: Result := 'starting';
    SERVICE_STOP_PENDING: Result := 'stopping';
    SERVICE_RUNNING: Result := 'RUN';
    SERVICE_CONTINUE_PENDING: Result := 'continuing...';
    SERVICE_PAUSE_PENDING: Result := 'pausing...';
    SERVICE_PAUSED: Result := 'PAUSE';
  End;
End;

// Inutilisé : ErrSt et StartSt
function ErrSt(Err : dword) : string;
Begin
  Case Err of
    SERVICE_ERROR_IGNORE: Result := 'ignore';
    SERVICE_ERROR_NORMAL: Result := 'normal';
    SERVICE_ERROR_SEVERE: Result := 'SEVERE';
    SERVICE_ERROR_CRITICAL: Result := 'CRITICAL';
   else  Result := '?';
  End;
End;

function StartSt(St : dword) : string;
Begin
  Case St of
    SERVICE_BOOT_START   : Result := 'BOOT';
    SERVICE_SYSTEM_START : Result := 'SYSTEM';
    SERVICE_AUTO_START   : Result := ' auto ';
    SERVICE_DEMAND_START : Result := 'Manual';
    SERVICE_DISABLED     : Result := 'Disabled';
   else  Result := '?';
  End;
End;

{  =========================================================================== }

procedure TForm1.FormCreate(Sender: TObject);
begin
  SC := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  Recreate1Click(Sender);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  CloseServiceHandle(SC);
end;

procedure TForm1.ChgEtatService(Sender: TObject);
var x : Integer;
    OS : THandle;
    SrvName : string;
    lpServiceStatus : _SERVICE_STATUS;

begin
  Timer1.Enabled := False;
  With ListView1, Items do
  If Count > 0 Then
    for x := 0 to Count-1 do
    With Item[x] do
      If Selected Then
      Begin
        SrvName := SubItems.Strings[0]; // ServiceInternalName
        OS := OpenService(SC, PChar(SrvName), SrvAccess);
        If Sender = Start1 Then
          StartService(OS, 0, PChar(nil^))
        else if Sender = Stop1 Then
          ControlService(OS, SERVICE_CONTROL_STOP, lpServiceStatus)
        else if Sender = Continue1 Then
          ControlService(OS, SERVICE_CONTROL_CONTINUE, lpServiceStatus)
        else if Sender = Pause1 Then
          ControlService(OS, SERVICE_CONTROL_PAUSE, lpServiceStatus);
        CloseServiceHandle(OS);
      End;

  Timer1.Enabled := True;
end;

procedure TForm1.Rafraichir(Sender: TObject);
var Tbl : array[1..500] of TEnumServiceStatus;
    card, card2, nbsvc, x : cardinal;

  function TrouverItem(Caption : string) : TListItem;
  var n : Integer;
  begin
    n := Memo1.Lines.IndexOf(Caption);
    With ListView1 do
      TrouverItem := Items.Item[n];
  end;

begin
  card2 := 0;
  nbsvc := 0;
  EnumServicesStatus(SC, SERVICE_WIN32, SERVICE_STATE_ALL,
    Tbl[1], SizeOf(Tbl), card, nbsvc, card2);
  If nbsvc > 0 Then
  for x := 1 to nbsvc do
    With Tbl[x], ServiceStatus do
    try
      TrouverItem(lpServiceName).SubItems.Strings[1] :=
        CurrState(dwCurrentState);
    except
      showmessage('Impossible de mettre à jour '+
          lpServiceName+'. Recréez la liste !');
    end;
end;

procedure TForm1.Recreate1Click(Sender: TObject);
var Tbl : array[1..500] of TEnumServiceStatus;
    card, card2, nbsvc, x : cardinal;
begin
  Timer1.Enabled := False;
  // Enumération
  ListView1.Clear;
  Memo1.Clear;
  card2 := 0;
  nbsvc := 0;
  EnumServicesStatus(SC, SERVICE_WIN32, SERVICE_STATE_ALL,
    Tbl[1], SizeOf(Tbl), card, nbsvc, card2);
  If nbsvc > 0 Then
  for x := 1 to nbsvc do
    With Tbl[x], ServiceStatus, ListView1.Items, Add do
    Begin
        Memo1.Lines.Add(lpServiceName);
        Caption := StrPas(lpDisplayName);
        SubItems.Add(StrPas(lpServiceName));
        SubItems.Add(CurrState(dwCurrentState));
    End;
   Timer1.Enabled := True;
end;

end.
