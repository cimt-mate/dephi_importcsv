program TK10IMP100;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Form1},
  ImportSetting in 'ImportSetting.pas' {Form2};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
