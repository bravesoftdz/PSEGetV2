program PSEGetGui;

uses
  Forms,
  uMain in 'uMain.pas' {frmMain},
  uExNotice in 'uExNotice.pas' {frmNotice},
  uPseParser in 'uPseParser.pas',
  DataAccessU in 'DataAccessU.pas',
  uPSEDataProvider in 'uPSEDataProvider.pas',
  SQLite3 in 'SQLite3.pas',
  SQLiteTable3 in 'SQLiteTable3.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmNotice, frmNotice);
  Application.Run;
end.
