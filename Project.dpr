program Project;

uses
  Vcl.Forms,
  Envoi_donnees in 'Envoi_donnees.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
