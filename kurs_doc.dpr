program kurs_doc;

uses
  Forms,
  main in 'main.pas' {FormInput};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormInput, FormInput);
  Application.Run;
end.
