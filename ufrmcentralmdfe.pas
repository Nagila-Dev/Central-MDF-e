unit ufrmcentralmdfe;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldb, db, Forms, Controls, Graphics, Dialogs, ComCtrls,
  ExtCtrls, DBExtCtrls, StdCtrls, DBGrids, DBCtrls, ExtDlgs, EditBtn,
  DateTimePicker, Grids, ufuncoes, uDM, ufrmRelatorioCentralMDF_E, zipper, ufrmRelatorioDAMDFE;

type

  { TfrmCentralMDFe }

  TfrmCentralMDFe = class(TForm)
    btnBuscarMDFe: TButton;
    btnImprimirMDFe: TButton;
    btnCancelarMDFe: TButton;
    btnGerarArquivoMDFe: TButton;
    btnEncerrarMDFe: TButton;
    Button6: TButton;
    btnVoltar: TButton;
    btnExecutarcancelamento: TButton;
    cbModeloNF: TComboBox;
    dsNOTIFICACAO: TDataSource;
    dsMDFE_EVENTOS: TDataSource;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    Edit1: TEdit;
    Label13: TLabel;
    Label19: TLabel;
    lcbFiliais: TDBLookupComboBox;
    dsFiliais: TDataSource;
    edtDataInicial: TDateEdit;
    edtDataFinal: TDateEdit;
    dsMDFE: TDataSource;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    dbgridManifesto: TDBGrid;
    DBGrid2: TDBGrid;
    DBGridMDFe: TDBGrid;
    edtDocumento: TEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    PageControl1: TPageControl;
    Panel1: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    qryFiliais: TSQLQuery;
    qryFiliaisFILIAL: TStringField;
    qryFiliaisPAR_COD: TLongintField;
    qryMDFECHAVE_MDFE: TStringField;
    qryMDFECPF_CONDUTOR: TStringField;
    qryMDFEDATA_AUT_MDFE: TDateField;
    qryMDFEDATA_EMISSAO: TDateField;
    qryMDFEDATA_VIAGEM: TDateField;
    qryMDFEFORMA_EMISSAO: TLongintField;
    qryMDFEHORA_DATA: TStringField;
    qryMDFEHORA_EMISSAO: TTimeField;
    qryMDFEHORA_VIAGEM: TTimeField;
    qryMDFEID_MD5: TStringField;
    qryMDFEMDFE_ID: TLongintField;
    qryMDFEMDV_RNTRC: TStringField;
    qryMDFEMODALIDADE: TLongintField;
    qryMDFEMODELO_DOC: TStringField;
    qryMDFENF_EMPRESA: TLongintField;
    qryMDFENOME_CONDUTOR: TStringField;
    qryMDFENUMERO_DOC: TLongintField;
    qryMDFEOBS: TMemoField;
    qryMDFEPLACA: TStringField;
    qryMDFEPROTOCOLO_MDFE: TStringField;
    qryMDFEPRO_COD: TStringField;
    qryMDFEPRO_PRED_CARGA_CEP: TStringField;
    qryMDFEPRO_PRED_CARGA_LAT: TFMTBCDField;
    qryMDFEPRO_PRED_CARGA_LONG: TFMTBCDField;
    qryMDFEPRO_PRED_DESCARGA_CEP: TStringField;
    qryMDFEPRO_PRED_DESCARGA_LAT: TFMTBCDField;
    qryMDFEPRO_PRED_DESCARGA_LONG: TFMTBCDField;
    qryMDFEPRO_PRED_EAN: TStringField;
    qryMDFEPRO_PRED_NCM: TStringField;
    qryMDFEPRO_PRED_NOME: TStringField;
    qryMDFESERIE_DOC: TStringField;
    qryMDFESTATUS_MDFE: TStringField;
    qryMDFETIPO_CARGA: TStringField;
    qryMDFETIPO_EMITENTE: TLongintField;
    qryMDFETIPO_TRANSPORTADOR: TLongintField;
    qryMDFETOT_CTE: TLongintField;
    qryMDFETOT_MERCADORIAS: TBCDField;
    qryMDFETOT_NF1_NF1A: TLongintField;
    qryMDFETOT_NFE: TLongintField;
    qryMDFETOT_PESO: TFMTBCDField;
    qryMDFEUF_CARREGAMENTO: TStringField;
    qryMDFEUF_DESCARREGAMENTO: TStringField;
    qryMDFEUND_MEDIDA: TLongintField;
    qryMDFE_EVENTOSLOG_MSG: TMemoField;
    qryMDFE_EVENTOSMDFE_COD_TP: TLongintField;
    qryMDFE_EVENTOSMDFE_CORRECAO: TMemoField;
    qryMDFE_EVENTOSMDFE_ID: TLongintField;
    qryMDFE_EVENTOSMDFE_LOTE: TLongintField;
    qryMDFE_EVENTOSMDFE_SEQ: TLongintField;
    qryMDFE_EVENTOSUSR_COD: TLongintField;
    qryMDFE_EVENTOSXML_EVENTO: TMemoField;
    qryNOTIFICACAOCHAVE_MDFE: TStringField;
    qryNOTIFICACAODATA_AUT_MDFE: TDateField;
    qryNOTIFICACAODATA_EMISSAO: TDateField;
    qryNOTIFICACAOLOG_MSG: TMemoField;
    qryNOTIFICACAOMDFE_COD_TP: TLongintField;
    qryNOTIFICACAOMDFE_CORRECAO: TMemoField;
    qryNOTIFICACAOMDFE_DATA: TDateField;
    qryNOTIFICACAOMDFE_HORA: TTimeField;
    qryNOTIFICACAOMDFE_ID: TLongintField;
    qryNOTIFICACAOMDFE_LOG_LIDO: TStringField;
    qryNOTIFICACAOMDFE_LOTE: TLongintField;
    qryNOTIFICACAOMDFE_PROTOCOLO: TStringField;
    qryNOTIFICACAOMDFE_SEQ: TLongintField;
    qryNOTIFICACAONF_EMPRESA: TLongintField;
    qryNOTIFICACAONUMERO_DOC: TLongintField;
    qryNOTIFICACAOPROTOCOLO_MDFE: TStringField;
    qryNOTIFICACAOSERIE_DOC: TStringField;
    qryNOTIFICACAOTOT_MERCADORIAS: TBCDField;
    qryNOTIFICACAOUF_DESCARREGAMENTO: TStringField;
    qryNOTIFICACAOUSR_COD: TLongintField;
    qryNOTIFICACAOXML_EVENTO: TMemoField;
    RadioTodas: TRadioButton;
    RadioDigitadas: TRadioButton;
    RadioTransmitida: TRadioButton;
    RadioAutorizada: TRadioButton;
    RadioCancelada: TRadioButton;
    RadioEncerrada: TRadioButton;
    qryMDFE: TSQLQuery;
    qryMDFE_EVENTOS: TSQLQuery;
    qryNOTIFICACAO: TSQLQuery;
    Shape1: TShape;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    tbAnterior: TToolButton;
    tbApagar: TToolButton;
    btBuscar: TToolButton;
    tbCancelar: TToolButton;
    tbFechar: TToolButton;
    tbGravar: TToolButton;
    tbIncluir: TToolButton;
    tbModificar: TToolButton;
    tbProximo: TToolButton;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    procedure btnBuscarMDFeClick(Sender: TObject);
    procedure btnImprimirMDFeClick(Sender: TObject);
    procedure btnCancelarMDFeClick(Sender: TObject);
    procedure btnGerarArquivoMDFeClick(Sender: TObject);
    procedure btnEncerrarMDFeClick(Sender: TObject);
    procedure btnExecutarcancelamentoClick(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure dbgridManifestoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormCreate(Sender: TObject);
    procedure qryMDFEAfterScroll(DataSet: TDataSet);
    procedure qryNOTIFICACAOAfterScroll(DataSet: TDataSet);
    procedure tbFecharClick(Sender: TObject);
  private

    var
      Texto: TStringStream;
      FZipper : TZipper;

  procedure ControlarBotoes(EmEdicao: Boolean);
  procedure ControlarCancelareEncerrar ();
  procedure MontarArquivoZipado(XML, Chave: String);
  procedure GerarArquivoZipado(Caminho: String);
  procedure CancelarMDFe();
  procedure EncerrarMDFe();
  procedure BuscarMDFe();
  procedure StatusMDFe();

  function TirarAcentoMDFe(str:String):String;

  public
    var select: String;

  end;

var
  frmCentralMDFe: TfrmCentralMDFe;

implementation
                                   //
{$R *.lfm}

{ TfrmCentralMDFe }

procedure TfrmCentralMDFe.FormCreate(Sender: TObject);
begin
  qryFiliais.Open;
  qryNOTIFICACAO.Open;

  Panel10.Visible:= False;

  select := qryMDFE.SQL.Text;

  lcbFiliais.ItemIndex:= 0;

  edtDataFinal.Date:= now;
  edtDataInicial.Date:= now;

  ufuncoes.PermissaoUsuario(DM.Id_usuario,46,DM.empresaSelecionada, frmCentralMDFe);

  ControlarBotoes(False);
end;

procedure TfrmCentralMDFe.qryMDFEAfterScroll(DataSet: TDataSet);
begin
  qryMDFE_EVENTOS.Close;
  qryMDFE_EVENTOS.ParamByName('MDFE_ID').AsInteger:= qryMDFEMDFE_ID.AsInteger;
  qryMDFE_EVENTOS.Open;

  ControlarCancelareEncerrar();
end;

procedure TfrmCentralMDFe.qryNOTIFICACAOAfterScroll(DataSet: TDataSet);
var PalavraTirarACento:String;
begin
  PalavraTirarACento:= TirarAcentoMDFe(qryNotificacao.FieldByName('LOG_MSG').AsString);
  Label19.Caption:= PalavraTirarACento;
end;

procedure TfrmCentralMDFe.tbFecharClick(Sender: TObject);
begin
  FreeAndNil(frmCentralMDFe);
end;

procedure TfrmCentralMDFe.MontarArquivoZipado(XML, Chave: String);
begin
  Texto:= TStringStream.Create(XML);
  FZipper.Entries.AddFileEntry(Texto, Chave + '.xml');
end;

procedure TfrmCentralMDFe.GerarArquivoZipado(Caminho: String);
begin
  FZipper.SaveToFile(Caminho + '.zip');
end;

procedure TfrmCentralMDFe.CancelarMDFe();
var loteCancelar: Integer;
    seguenciacancelar: Integer;
begin
   with qryMDFE_EVENTOS do
   begin
     Append;
     loteCancelar     := qryMDFE_EVENTOSMDFE_LOTE.AsInteger + 1;

     FieldByName('MDFE_ID'      ).AsInteger:= qryMDFEMDFE_ID.AsInteger;
     FieldByName('MDFE_SEQ'     ).AsInteger:= 1;
     FieldByName('MDFE_COD_TP'  ).AsInteger:= 110111;
     //FieldByName('XML_EVENTO' ).AsString := ;  NULL
     FieldByName('MDFE_LOTE'    ).AsInteger:= loteCancelar;
     //FieldByName('LOG_MSG'    ).AsString := ;  NULL
     FieldByName('USR_COD'      ).AsInteger:= DM.usuarioLogado;
     FieldByName('MDFE_CORRECAO').AsString := Edit1.Text;
     post;
  end;

  qryMDFE_EVENTOS.Edit;
  if qryMDFE_EVENTOS.State in dsEditModes then
    qryMDFE_EVENTOS.ApplyUpdates();

  qryMDFE.SQLTransaction.CommitRetaining;

end;

procedure TfrmCentralMDFe.EncerrarMDFe();
  var loteEncerrar: integer;
begin
  with qryMDFE_EVENTOS do
  begin
    Append;
    loteEncerrar := qryMDFE_EVENTOSMDFE_LOTE.AsInteger + 1;

    FieldByName('MDFE_ID'      ).AsInteger:= qryMDFEMDFE_ID.AsInteger;
    FieldByName('MDFE_SEQ'     ).AsInteger:= 1;
    FieldByName('MDFE_COD_TP'  ).AsInteger:= 110112;
    //FieldByName('XML_EVENTO' ).AsString := ;   NULL
    FieldByName('MDFE_LOTE'    ).AsInteger:= loteEncerrar;
    //FieldByName('LOG_MSG'    ).AsString := ;   NULL
    FieldByName('USR_COD'      ).AsInteger:= DM.usuarioLogado;
    FieldByName('MDFE_CORRECAO').AsString := 'ENCERRAMENTO DO DOCUMENTO';
    post;
  end;

  qryMDFE_EVENTOS.Edit;
  if qryMDFE_EVENTOS.State in dsEditModes then
    qryMDFE_EVENTOS.ApplyUpdates();

  qryMDFE.SQLTransaction.CommitRetaining;
end;

procedure TfrmCentralMDFe.BuscarMDFe();
begin
  with qryMDFE do
  begin
    Close;

    if (edtDocumento.Text = '') then
    begin
      SQL.Text := select;
      SQL.Add ('and ((DATA_EMISSAO BETWEEN :dtinicial1 AND :dtinicial2 ) or (DATA_AUT_MDFE BETWEEN :dtfinal1 AND :dtfinal2 ))');
      ParamByName('dtinicial1').AsDateTime := edtDataInicial.Date;
      ParamByName('dtinicial2').AsDateTime := edtDataFinal.Date;
      ParamByName('dtfinal1'  ).AsDateTime := edtDataInicial.Date;
      ParamByName('dtfinal2'  ).AsDateTime := edtDataFinal.Date;
    end else
    begin
      SQL.Text:= select;
      SQL.add ('and (M.NUMERO_DOC = :NUMERO_DOC) ');
      ParamByName('NUMERO_DOC').AsInteger := StrToInt(edtDocumento.Text);
    end;

    ParamByName('filial').AsString :=  VariantParaString(lcbFiliais.KeyValue);

    StatusMDFe();

    Open;
  end;
end;

procedure TfrmCentralMDFe.StatusMDFe();
begin
  with qryMDFE do
  begin
    if RadioTodas.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'TODAS';

    if RadioCancelada.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'C';

    if RadioAutorizada.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'A';

    if RadioDigitadas.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'D';

    if RadioTransmitida.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'T';

    if RadioEncerrada.Checked then
      ParamByName('STATUS_MDFE').AsString:= 'TODAS';
  end;

end;

function TfrmCentralMDFe.TirarAcentoMDFe(str: String): String;
const
  AccentedChars :array[0..25] of string =
  ('á','à','ã','â','ä','é','è','ê','ë','í','ì','ï','î','ó','ò','õ','ô','ö','ú','ù','ü','û','ç','ñ','ý','ÿ');
  NormalChars :array[0..25] of string =
  ('a','a','a','a','a','e','e','e','e','i','i','i','i','o','o','o','o','o','u','u','u','u','c','n','y','y');
var
  i: Integer;
begin
  Result := str;
  for i := 0 to 25 do
    Result := StringReplace(Result, AccentedChars[i], NormalChars[i],
    [rfReplaceAll]);
end;

procedure TfrmCentralMDFe.btnBuscarMDFeClick(Sender: TObject);
begin
  if ('' = VariantParaString(lcbFiliais.KeyValue)) or (cbModeloNF.ItemIndex = -1) then
  begin
    ShowMessage('Selecione uma afilial e modelo para a busca!');
    Exit;
  end;
  BuscarMDFe();
end;

procedure TfrmCentralMDFe.btnImprimirMDFeClick(Sender: TObject);
begin
  //Application.CreateForm(TfrmRelatorioCentralMDF_E,frmRelatorioCentralMDF_E);
  //frmRelatorioCentralMDF_E.RLReport1.PreviewModal;

  Application.CreateForm(TfrmRelatorioDAMDFE, frmRelatorioDAMDFE);
   with frmRelatorioDAMDFE do
   begin
     qryRelatorioMDFE.Close;
     qryRelatorioMDFE.ParamByName('MDFE_ID').AsInteger:= qryMDFEMDFE_ID.AsInteger;
     qryRelatorioMDFE.Open;

     if qryRelatorioMDFE.IsEmpty then
     Exit
     else
     RLReport1.PreviewModal;
   end;
end;

procedure TfrmCentralMDFe.btnCancelarMDFeClick(Sender: TObject);
begin
  Panel10.Visible:= true;
end;

procedure TfrmCentralMDFe.btnGerarArquivoMDFeClick(Sender: TObject);
begin
  SaveDialog1 := TSaveDialog.Create(SaveDialog1);
  SaveDialog1.FileName:= 'MDFe';
  SaveDialog1.Execute;

  try
    FZipper := TZipper.Create;

    if MessageDlg('Deseja Imprimir lista completa?', mtConfirmation, [mbYes , mbNo], 1) = mrNo then
      begin
        with qryMDFE_EVENTOS do
        begin
          First;
          while not EOF do
          begin
            MontarArquivoZipado(FieldByName('XML_EVENTO').AsString,qryMDFE.FieldByName('CHAVE_MDFE').AsString);
            Next;
          end;
        end;
      end else
      begin
        qryMDFE.First;
        while not qryMDFE.EOF do
        begin
          with qryMDFE_EVENTOS do
          begin
            First;
            while not EOF do
            begin
              MontarArquivoZipado(FieldByName('XML_EVENTO').AsString,qryMDFE.FieldByName('CHAVE_MDFE').AsString);
              Next;
            end;
          end;
          qryMDFE.Next;
        end;
      end;
      GerarArquivoZipado(SaveDialog1.FileName);

  finally
     FZipper.free;
  end;

end;

procedure TfrmCentralMDFe.btnEncerrarMDFeClick(Sender: TObject);
begin
  EncerrarMDFe();
end;

procedure TfrmCentralMDFe.btnExecutarcancelamentoClick(Sender: TObject);
begin
  CancelarMDFe();
end;

procedure TfrmCentralMDFe.Button6Click(Sender: TObject);
begin
  With qryNOTIFICACAO do
  begin
    Edit;
    FieldByName('LOG_MSG').AsVariant      := null;
    FieldByName('MDFE_LOG_LIDO').AsVariant:= null;
    Post;

    if State in dsEditModes then
      Post;
    ApplyUpdates(0);

    qryMDFE.SQLTransaction.CommitRetaining;

    DBGridMDFe.DataSource.DataSet.Refresh;

    Label19.Caption:= '';
  end;
end;

procedure TfrmCentralMDFe.dbgridManifestoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if (Column.FieldName = 'STATUS_MDFE') then
  begin
    with qryMDFE do
    begin
      with dbgridManifesto do
      begin
        if (FieldByName('STATUS_MDFE').AsString = 'A') then
        begin
          Canvas.Brush.Color:= clGreen;
          Canvas.Font.Color := clWhite;
          Canvas.Font.Bold  := true;
        end;

        if (FieldByName('STATUS_MDFE').AsString = 'C') then
        begin
          Canvas.Brush.Color:= clRed;
          Canvas.Font.Color := clWhite;
          Canvas.Font.Bold  := true;
        end;

        if (FieldByName('STATUS_MDFE').AsString = 'D') then
        begin
          Canvas.Brush.Color:= clYellow;
          Canvas.Font.Color := clRed;
          Canvas.Font.Bold  := true;
        end;

        if (FieldByName('STATUS_MDFE').AsString = 'T') then
        begin
          Canvas.Brush.Color:= clYellow;
          Canvas.Font.Color := clRed;
          Canvas.Font.Bold  := true;
        end;
      end;
    end;
end;
  if (gdSelected in State) then
    TDBGrid(Sender).Canvas.Font.Color:= clBlack;

   TDBGrid(Sender).Canvas.FillRect(Rect);
   TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol,Column, State);
end;

procedure TfrmCentralMDFe.ControlarBotoes(EmEdicao: Boolean);
begin
  tbIncluir.Enabled   := EmEdicao;
  tbModificar.Enabled := EmEdicao;
  tbGravar.Enabled    := EmEdicao;
  tbCancelar.Enabled  := EmEdicao;
  tbApagar.Enabled    := EmEdicao;
  btBuscar.Enabled    := not EmEdicao;
  tbAnterior.Enabled  := not EmEdicao;
  tbProximo.Enabled   := not EmEdicao;
  tbFechar.Enabled    := not EmEdicao;
end;

procedure TfrmCentralMDFe.ControlarCancelareEncerrar();
begin
  if (qryMDFE_EVENTOS.Locate('MDFE_COD_TP', 110112, [])) or (qryMDFE_EVENTOS.Locate('MDFE_COD_TP', 110111, [])) then
  begin
    btnCancelarMDFe.Enabled:= false;
  end else
    btnCancelarMDFe.Enabled:= True;
end;

end.

