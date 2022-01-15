unit frm_cadastro_cliente;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Mask, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.Buttons, Vcl.ToolWin, REST.Client, System.JSON, XMLDoc, XMLIntf;

type
  TEndereco = class
    private
      FLogradouro : String;
      FBairro     : String;
      FCep        : String;
      FCidade     : String;
      FEstado     : String;
    published
      property Logradouro : String read FLogradouro write FLogradouro;
      property Bairro     : String read FBairro     write FBairro    ;
      property Cep        : String read FCep        write FCep       ;
      property Cidade     : String read FCidade     write FCidade    ;
      property Estado     : String read FEstado     write FEstado    ;
  end;

  Tf_cadastro_cliente = class(TForm)
    ToolBar2: TToolBar;
    sbtnovo: TSpeedButton;
    sbtsalva: TSpeedButton;
    sbtsair: TSpeedButton;
    StatusBar1: TStatusBar;
    edt_nome: TLabeledEdit;
    edt_Identidade: TLabeledEdit;
    LinkLabel1: TLinkLabel;
    msk_telefone: TMaskEdit;
    edt_email: TLabeledEdit;
    LinkLabel2: TLinkLabel;
    msk_CEP: TMaskEdit;
    edt_logradouro: TLabeledEdit;
    edt_complemento: TLabeledEdit;
    edt_bairro: TLabeledEdit;
    edt_cidade: TLabeledEdit;
    edt_estado: TLabeledEdit;
    sbtBuscaCEP: TSpeedButton;
    edt_numero: TLabeledEdit;
    msk_CPF: TMaskEdit;
    LinkLabel3: TLinkLabel;
    Memo1: TMemo;
    procedure sbtsairClick(Sender: TObject);
    procedure sbtBuscaCEPClick(Sender: TObject);
    procedure msk_CEPKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure sbtnovoClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure sbtsalvaClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    function getCEP(cep: string): TEndereco;
    function EnviarEmail(const AAssunto, ADestino, AAnexo: String; ACorpo: TStrings): Boolean;
  public
    { Public declarations }
  end;

var
  f_cadastro_cliente: Tf_cadastro_cliente;

implementation

uses
IniFiles,
IdComponent,
IdTCPConnection,
IdTCPClient,
IdHTTP,
IdBaseComponent,
IdMessage,
IdExplicitTLSClientServerBase,
IdMessageClient,
IdSMTPBase,
IdSMTP,
IdIOHandler,
IdIOHandlerSocket,
IdIOHandlerStack,
IdSSL,
IdSSLOpenSSL,
IdAttachmentFile,
IdText;

{$R *.dfm}
function getEmailValido(email : string) : string;
var
  aux : string;
begin
    aux    := '';
    if pos('@',email) > 0 then
      begin aux    := lowerCase(email); end;

    result := aux;
end;

function Tf_cadastro_cliente.EnviarEmail(const AAssunto, ADestino, AAnexo: String; ACorpo: TStrings): Boolean;
var
  IniFile              : TIniFile;
  sFrom                : String;
  sBccList             : String;
  sHost                : String;
  iPort                : Integer;
  sUserName            : String;
  sPassword            : String;

  idMsg                : TIdMessage;
  IdText               : TIdText;
  idSMTP               : TIdSMTP;
  IdSSLIOHandlerSocket : TIdSSLIOHandlerSocketOpenSSL;
begin
  try
    try
      //Criação e leitura do arquivo INI com as configurações
      IniFile                          := TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'Config.ini');
      sFrom                            := IniFile.ReadString('Email' , 'From'     , sFrom);
      sBccList                         := IniFile.ReadString('Email' , 'BccList'  , sBccList);
      sHost                            := IniFile.ReadString('Email' , 'Host'     , sHost);
      iPort                            := IniFile.ReadInteger('Email', 'Port'     , iPort);
      sUserName                        := IniFile.ReadString('Email' , 'UserName' , sUserName);
      sPassword                        := IniFile.ReadString('Email' , 'Password' , sPassword);

      //Configura os parâmetros necessários para SSL
      IdSSLIOHandlerSocket                   := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
      IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
      IdSSLIOHandlerSocket.SSLOptions.Mode  := sslmClient;

      //Variável referente a mensagem
      idMsg                            := TIdMessage.Create(Self);
      idMsg.CharSet                    := 'utf-8';
      idMsg.Encoding                   := meMIME;
      idMsg.From.Name                  := 'Sistema de cadastro de clientes';
      idMsg.From.Address               := sFrom;
      idMsg.Priority                   := mpNormal;
      idMsg.Subject                    := AAssunto;

      //Add Destinatário(s)
      idMsg.Recipients.Add;
      idMsg.Recipients.EMailAddresses := ADestino;
      idMsg.CCList.EMailAddresses      := 'xxxxx@xxxxxx';
      idMsg.BccList.EMailAddresses    := sBccList;
      idMsg.BccList.EMailAddresses    := 'xxxxx@xxxxxx'; //Cópia Oculta

      //Variável do texto
      idText := TIdText.Create(idMsg.MessageParts);
      idText.Body.Add(ACorpo.Text);
      idText.ContentType := 'text/html; text/plain; charset=iso-8859-1';

      //Prepara o Servidor
      IdSMTP                           := TIdSMTP.Create(Self);
      IdSMTP.IOHandler                 := IdSSLIOHandlerSocket;
      IdSMTP.UseTLS                    := utUseImplicitTLS;
      IdSMTP.AuthType                  := satDefault;
      IdSMTP.Host                      := sHost;
      IdSMTP.AuthType                  := satDefault;
      IdSMTP.Port                      := iPort;
      IdSMTP.Username                  := sUserName;
      IdSMTP.Password                  := sPassword;

      //Conecta e Autentica
      IdSMTP.Connect;
      IdSMTP.Authenticate;

        if FileExists(AAnexo) then
         begin TIdAttachmentFile.Create(idMsg.MessageParts, AAnexo); end;

      //Se a conexão foi bem sucedida, envia a mensagem
      if IdSMTP.Connected then
      begin
        try
          IdSMTP.Send(idMsg);
        except on E:Exception do
          begin
            ShowMessage('Erro ao tentar enviar: ' + E.Message);
          end;
        end;
      end;

      //Depois de tudo pronto, desconecta do servidor SMTP
      if IdSMTP.Connected then
        IdSMTP.Disconnect;

      Result := True;
    finally
      IniFile.Free;

      UnLoadOpenSSLLibrary;

      FreeAndNil(idMsg);
      FreeAndNil(IdSSLIOHandlerSocket);
      FreeAndNil(idSMTP);
    end;
  except on e:Exception do
    begin
      Result := False;
    end;
  end;
end;

procedure Tf_cadastro_cliente.FormActivate(Sender: TObject);
begin
    sbtnovoClick(self);
end;

procedure Tf_cadastro_cliente.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if VK_ESCAPE = key then
    begin sbtsairClick(self); end;

  if vk_f1 = key then
    begin sbtnovoClick(self); end;

  if VK_F2 = key then
    begin sbtsalvaClick(self); end;

end;

function Tf_cadastro_cliente.getCEP(cep: string): TEndereco;
var
  data: TJSONObject;
  RESTClient1: TRESTClient;
  RESTRequest1: TRESTRequest;
  RESTResponse1: TRESTResponse;
  endereco: TEndereco;

begin
  RESTClient1 := TRESTClient.Create(nil);
  RESTRequest1 := TRESTRequest.Create(nil);
  RESTResponse1 := TRESTResponse.Create(nil);
  RESTRequest1.Client := RESTClient1;
  RESTRequest1.Response := RESTResponse1;
  RESTClient1.BaseURL := 'https://viacep.com.br/ws/' + cep + '/json';
  RESTRequest1.Execute;
  data := RESTResponse1.JSONValue as TJSONObject;
  endereco := TEndereco.Create;

  if data <> nil then
    begin
      Try
        endereco := TEndereco.Create;
        if Assigned(data) then
        begin
            if data.Count > 1 then
             begin
              endereco.FLogradouro:=data.Values['logradouro'].Value;
              endereco.FBairro:=data.Values['bairro'].Value;
              endereco.FCidade:=data.Values['localidade'].Value;
              endereco.FEstado:=data.Values['uf'].Value;
              endereco.FCEP:=data.Values['cep'].Value;
             end else MessageBox(0, PChar('CEP não cadastrado'), PChar('Aviso'), MB_ICONINFORMATION or MB_OK);
        end;
      finally
        FreeAndNil(data);
      end;
    end else MessageBox(0, PChar('CEP incorreto'), PChar('Aviso'), MB_ICONINFORMATION or MB_OK);

  FreeAndNil(RESTClient1);
  FreeAndNil(RESTRequest1);

  Result := endereco;
end;


procedure Tf_cadastro_cliente.msk_CEPKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if vk_return = key then
    begin sbtBuscaCEPClick(self); end;
end;

procedure Tf_cadastro_cliente.sbtBuscaCEPClick(Sender: TObject);
var
  endereco : TEndereco;
begin
  endereco := getCEP(msk_CEP.Text);

  edt_logradouro.Text := endereco.FLogradouro;
  edt_bairro.Text := endereco.FBairro;
  edt_cidade.Text := endereco.FCidade;
  edt_estado.Text := endereco.FEstado;

end;

procedure Tf_cadastro_cliente.sbtnovoClick(Sender: TObject);
begin
  edt_nome.Clear;
  edt_Identidade.Clear;
  msk_CPF.Clear;
  msk_telefone.Clear;
  edt_email.Clear;
  msk_CEP.Clear;
  edt_logradouro.Clear;
  edt_numero.Clear;
  edt_complemento.Clear;
  edt_bairro.Clear;
  edt_cidade.Clear;
  edt_estado.Clear;

  edt_nome.SetFocus;
end;

procedure Tf_cadastro_cliente.sbtsairClick(Sender: TObject);
begin
  Close;
end;

procedure Tf_cadastro_cliente.sbtsalvaClick(Sender: TObject);
var
  XML: TXMLDocument;
  NodeEndereco : IXMLNode;
  sequencia : string;
  arquivo : string;
begin

  if getEmailValido(edt_email.Text) = '' then
    begin
      MessageBox(0, PChar('E-mail inválido'), PChar('Erro'), MB_ICONSTOP or MB_OK);
      Abort;
    end;

  XML := TXMLDocument.Create(nil);
  XML.Active := True;
  NodeEndereco := XML.AddChild('Endereço');

  NodeEndereco.ChildValues['Nome'] := edt_nome.Text;
  NodeEndereco.ChildValues['Identidade'] := edt_Identidade.Text;
  NodeEndereco.ChildValues['CPF'] := msk_CPF.Text;
  NodeEndereco.ChildValues['Telefone'] := msk_telefone.Text;
  NodeEndereco.ChildValues['Email'] := edt_email.Text;
  NodeEndereco.ChildValues['CEP'] := msk_CEP.Text;
  NodeEndereco.ChildValues['Logradouro'] := edt_logradouro.Text;
  NodeEndereco.ChildValues['Número'] := edt_numero.Text;
  NodeEndereco.ChildValues['Complemento'] := edt_complemento.Text;
  NodeEndereco.ChildValues['Bairro'] := edt_bairro.Text;
  NodeEndereco.ChildValues['Cidade'] := edt_cidade.Text;
  NodeEndereco.ChildValues['Estado'] := edt_estado.Text;

  arquivo :=  ExtractFileDir(Application.ExeName) + '\registro_'+IntToStr(Random(9999))+'.xml';
  XML.SaveToFile(arquivo);

  memo1.Lines.Text := 'Nome: '+edt_nome.Text;
  memo1.Lines.Add('Identidade: '+ edt_Identidade.Text);
  memo1.Lines.Add('CPF: '+ msk_CPF.Text);
  memo1.Lines.Add('Telefone: '+ msk_telefone.Text);
  memo1.Lines.Add('Email: '+ edt_email.Text);
  memo1.Lines.Add('CEP: '+msk_CEP.Text);
  memo1.Lines.Add('Logradouro: '+edt_logradouro.Text);
  memo1.Lines.Add('Número: '+edt_numero.Text);
  memo1.Lines.Add('Complemento: '+ edt_complemento.Text);
  memo1.Lines.Add('Bairro: '+edt_bairro.Text);
  memo1.Lines.Add('Cidade: '+edt_cidade.Text);
  memo1.Lines.Add('Estado: '+ edt_estado.Text);

  EnviarEmail('Cadastro de Clientes', edt_email.Text, arquivo, memo1.Lines);

  MessageBox(0, PChar('Dados gravados com sucesso'), PChar('Aviso'), MB_ICONINFORMATION or MB_OK);
  sbtnovoClick(self);

end;

end.
