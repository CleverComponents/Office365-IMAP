unit main;

interface

uses
  System.UITypes, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, clOAuth, clTcpClient,
  clTcpClientTls, clTcpCommandClient, clMC, clImap4, Vcl.ComCtrls, Vcl.ImgList,
  clMailMessage, clSocketUtils, ImapEnvelope;

type
  TMainForm = class(TForm)
    edtUser: TEdit;
    btnLogin: TButton;
    btnLogout: TButton;
    Label6: TLabel;
    tvFolders: TTreeView;
    lvMessages: TListView;
    Label8: TLabel;
    edtFrom: TEdit;
    Label9: TLabel;
    edtSubject: TEdit;
    memBody: TMemo;
    Label1: TLabel;
    clImap: TclImap4;
    clMailMessage: TclMailMessage;
    Images: TImageList;
    clOAuth1: TclOAuth;
    Label2: TLabel;
    procedure tvFoldersChange(Sender: TObject; Node: TTreeNode);
    procedure lvMessagesClick(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure btnLogoutClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    FChanging: Boolean;

    procedure FillFolderList;
    procedure AddFolderToList(AName: string);
    procedure ClearMessage;
    procedure FillMessage(const AResponse: string; AMsgNo: Integer);
    procedure FillMessages(AResponse: TStrings);
    function GetFolderName(Node: TTreeNode): string;
    procedure EnableControls(AEnabled: Boolean);
    procedure Logout;
    function GetMessageId(const AResponseLine: string): Integer;
    function ParseFlags(const AResponse: string; var Index: Integer): string;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.btnLoginClick(Sender: TObject);
begin
  if (FChanging) then Exit;

  if (clImap.Active) then Exit;

  clOAuth1.AuthUrl := 'https://login.live.com/oauth20_authorize.srf';
  clOAuth1.TokenUrl := 'https://login.live.com/oauth20_token.srf';
  clOAuth1.RedirectUrl := 'http://localhost';
  clOAuth1.ClientID := 'a0a907aa-1e38-4bdb-8764-c4f931051018';
  clOAuth1.ClientSecret := '6FYd=PdPS-06UgOdNlFon2TXo*BDyAi-';
  clOAuth1.Scope := 'wl.imap wl.offline_access';

  clImap.Server := 'outlook.office365.com';
  clImap.Port := 993;
  clImap.UseTLS := ctImplicit;

  clImap.UserName := edtUser.Text;

  clImap.Authorization := clOAuth1.GetAuthorization();

  clImap.Open();

  FillFolderList();
end;

procedure TMainForm.FillFolderList;
var
  i: integer;
  list: TStrings;
begin
  list := TStringList.Create();
  try
    tvFolders.Items.BeginUpdate();
    tvFolders.Items.Clear();

    clImap.GetMailBoxes(list);

    for i := 0 to list.Count - 1 do
    begin
      AddFolderToList(list[i]);
    end;
  finally
    tvFolders.Items.EndUpdate();
    list.Free();
  end;
end;

procedure TMainForm.AddFolderToList(AName: string);
var
  Papa, N: TTreeNode;
  S: string;
  i: Integer;
begin
  Papa := nil;
  N := tvFolders.Items.GetFirstNode();
  if AName[1] = clImap.MailBoxSeparator then
  begin
    Delete(AName, 1, 1);
  end;

  while True do
  begin
    i := Pos(clImap.MailBoxSeparator, AName);
    if (i = 0) then
    begin
      Papa := tvFolders.Items.AddChild(Papa, AName);
      Papa.ImageIndex := 0;
      Papa.SelectedIndex := 0;
      Break;
    end else
    begin
      S := Copy(AName, 1, i - 1);
      Delete(AName, 1, i);
      while ((N <> nil) and (N.Text <> S)) do
      begin
        N := N.getNextSibling;
      end;
      if (N = nil) then
      begin
        Papa := tvFolders.Items.AddChild(Papa, S);
      end else
      begin
        Papa := N;
      end;
      N := Papa.GetFirstChild();
    end;
  end;
end;

procedure TMainForm.btnLogoutClick(Sender: TObject);
begin
  Logout();
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := (not clImap.Active) or (MessageDlg('Do you want to exit?', mtConfirmation, [mbYes, mbNo], 0) = mrYes);
  if CanClose then
  begin
    Logout();
  end;
end;

procedure TMainForm.lvMessagesClick(Sender: TObject);
begin
  if (FChanging) then Exit;

  FChanging := True;
  try
    EnableControls(False);

    if clImap.Active and (lvMessages.Selected <> nil) then
    begin
      clImap.RetrieveMessage(Integer(lvMessages.Selected.Data), clMailMessage);

      edtFrom.Text := clMailMessage.From.FullAddress;
      edtSubject.Text := clMailMessage.Subject;
      memBody.Lines := clMailMessage.MessageText;
    end else
    begin
      ClearMessage();
    end;
  finally
    FChanging := False;
    EnableControls(True);
  end;
end;

procedure TMainForm.tvFoldersChange(Sender: TObject; Node: TTreeNode);
begin
  if (FChanging) then Exit;

  FChanging := True;
  try
    EnableControls(False);

    lvMessages.Items.Clear();
    ClearMessage();

    if clImap.Active and Assigned(tvFolders.Selected) then
    begin
      clImap.SelectMailBox(GetFolderName(tvFolders.Selected));

      if (clImap.CurrentMailBox.ExistsMessages > 0) then
      begin
        clImap.SendTaggedCommand('FETCH 1:* (ENVELOPE FLAGS)', [IMAP_OK]);
        FillMessages(clImap.Response);
      end;
    end;
  finally
    FChanging := False;
    EnableControls(True);
  end;
end;

function TMainForm.GetFolderName(Node: TTreeNode): string;
begin
  if (Node = nil) then
  begin
    Result := ''
  end else
  begin
    Result := Node.Text;
    while (Node.Parent <> nil) do
    begin
      Node := Node.Parent;
      Result := Node.Text + clImap.MailBoxSeparator + Result;
    end;
  end;
end;

function TMainForm.GetMessageId(const AResponseLine: string): Integer;
var
  ind: Integer;
begin
  ind := System.Pos(' FETCH', UpperCase(AResponseLine));
  Result := 0;
  if (ind > 3) then
  begin
    Result := StrToIntDef(Trim(System.Copy(AResponseLine, 2, ind - 1)), 0);
  end;
end;

procedure TMainForm.FillMessages(AResponse: TStrings);
var
  i, msgNo: Integer;
  response: string;
begin
  lvMessages.Items.Clear();
  ClearMessage();

  msgNo := 1;
  response := '';
  for i := 0 to AResponse.Count - 1 do
  begin
    if (GetMessageId(AResponse[i]) = msgNo) then
    begin
      FillMessage(response, msgNo - 1);
      response := AResponse[i] + #13#10;
      Inc(msgNo);
    end else
    begin
      response := response + AResponse[i] + #13#10;
    end;
  end;
  FillMessage(response, msgNo - 1);
end;

function TMainForm.ParseFlags(const AResponse: string; var Index: Integer): string;
const
  FlagsLexem = 'FLAGS';
var
  ind, indEnd: Integer;
begin
  Result := '';
  ind := system.Pos(FlagsLexem, UpperCase(AResponse), Index);
  if (ind > 0) then
  begin
    ind := system.Pos('(', AResponse, ind + Length(FlagsLexem));
    if (ind > 0) then
    begin
      indEnd := system.Pos(')', AResponse, ind);
      if (indEnd > 0) then
      begin
        Result := system.Copy(AResponse, ind + 1, indEnd - ind - 1);
        Index := indEnd + 1;
      end;
    end;
  end;
end;

procedure TMainForm.FillMessage(const AResponse: string; AMsgNo: Integer);
var
  item: TListItem;
  envelope: TImapEnvelope;
  ind: Integer;
begin
  if (AResponse = '') then Exit;

  envelope := TImapEnvelope.Create();
  try
    ind := 1;
    envelope.Parse(AResponse, ind);

    item := lvMessages.Items.Insert(0);
    item.Data := Pointer(AMsgNo);

    item.Caption := envelope.Subject;
    item.SubItems.Clear();

    item.SubItems.Add(envelope.From);
    item.SubItems.Add(DateTimeToStr(envelope.Date));
    item.SubItems.Add(ParseFlags(AResponse, ind));
  finally
    envelope.Free();
  end;
end;

procedure TMainForm.ClearMessage;
begin
  edtFrom.Text := '';
  edtSubject.Text := '';
  memBody.Lines.Clear();
end;

procedure TMainForm.Logout;
begin
  try
    clImap.Close();
  except
    on EclSocketError do;
  end;

  try
    clOAuth1.Close();
  except
    on EclSocketError do;
  end;

  tvFolders.Items.Clear();
  lvMessages.Clear();
  ClearMessage();
end;

procedure TMainForm.EnableControls(AEnabled: Boolean);
begin
  btnLogin.Enabled := AEnabled;
  btnLogout.Enabled := AEnabled;

  if (AEnabled) then
  begin
    Cursor := crArrow;
  end else
  begin
    Cursor := crHourGlass;
  end;
end;

end.
