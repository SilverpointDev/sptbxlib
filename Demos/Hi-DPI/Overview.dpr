program Overview;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  VCL.Styles.DPIAware in 'VCL.Styles.DPIAware.pas',
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  TStyleManager.TrySetStyle('Windows10 Dark');
  Application.Initialize;
  TStyleManager.TrySetStyle('Windows10 Blue');
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
