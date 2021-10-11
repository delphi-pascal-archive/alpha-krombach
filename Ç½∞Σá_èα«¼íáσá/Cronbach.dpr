Program Cronbach;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1};

{$R *.res}

Begin
  Application.Initialize;
  Application.Title:='Alpha de Cronbach';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
End.
