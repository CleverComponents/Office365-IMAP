program Office365IMAP;

uses
  Vcl.Forms,
  main in 'main.pas' {MainForm},
  ImapEnvelope in 'ImapEnvelope.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
