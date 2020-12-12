program Spreadsheet;

uses
  Vcl.Forms,
  MainScreen in 'Forms\MainScreen.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TSpreadsheet, Sheet);
  Application.Run;
end.
