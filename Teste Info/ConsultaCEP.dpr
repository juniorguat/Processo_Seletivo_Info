program ConsultaCEP;

uses
  Vcl.Forms,
  frm_cadastro_cliente in 'frm_cadastro_cliente.pas' {f_cadastro_cliente};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(Tf_cadastro_cliente, f_cadastro_cliente);
  Application.Run;
end.
