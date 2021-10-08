unit uImportExcel;

interface
uses
  {$IF CompilerVersion > 28} Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, System.Win.comobj, Vcl.StdCtrls,
  Data.DB, Vcl.DBGrids,
  Vcl.ExtCtrls, Vcl.ComCtrls
  {$ELSE}
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, Grids, comobj, StdCtrls,
  DB,
  ExtCtrls, ComCtrls {$IFEND};
type
  TImportExcel = class(TComponent)
  private

    FAutor: string;
    FE_mail: string;
    FVersao: string;
    FCanal: string;
    FExcelFile: string;

  public
    constructor Create(AOwner: TComponent); override; export;
  published
    property ExcelFile: string read FExcelFile write FExcelFile;
    property Autor: string read FAutor;
    property E_Mail: string read FE_mail;
    property Versao: string read FVersao;
    property Canal: string read FCanal;

    function ExcelParaStringGrid(const AGrid: TStringGrid;
      const pProgressBar: TProgressBar): Boolean;
    procedure ExportarParaExcel(const pDataSet: TDataSet; pCaption: string; pFieldsPref: array of string);
    procedure Importar();
  end;
  procedure Register;

implementation

{ TImportExcel }

procedure Register;
begin
  RegisterComponents('Importa Excel', [TImportExcel]);
end;

constructor TImportExcel.Create;
begin
 inherited;
 FAutor  := 'Josué Elias';
 FE_mail := 'aulas.josue@gmail.com';
 FVersao := '1.0.1.8';
 FCanal  := 'https://www.youtube.com/channel/UCFF3H1ZQyE7uQsvGV-IddLA';
end;

function TImportExcel.ExcelParaStringGrid(const AGrid: TStringGrid; const pProgressBar: TProgressBar): Boolean;
const
	xlCellTypeLastCell = $0000000B;
var
	XLApp, Sheet: OLEVariant;
	RangeMatrix: Variant;
	x, y, k, r: Integer;
begin
  Result:=False;

  //meu codigo
  XLApp:=CreateOleObject('Excel.Application');
  try
    //Esconde Excel
    XLApp.Workbooks.Open(FExcelFile);
    Sheet := XLApp.Workbooks[ExtractFileName(FExcelFile)].WorkSheets[1];
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    //Pegar o número da última linha
    x := XLApp.ActiveCell.Row;
    //Pegar o número da última coluna
    y:=XLApp.ActiveCell.Column;
    //Seta Stringgrid linha e coluna
    AGrid.RowCount:=x;
    AGrid.ColCount:=y;
    //Associa a variant WorkSheet com a variant do Delphi
    Application.ProcessMessages;
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

    //Cria o loop para listar os registros no TStringGrid
    k:=1;
    if pProgressBar <> nil then begin
      pProgressBar.Max      := x;
      pProgressBar.Position := 0;
    end;

    repeat
      Application.ProcessMessages;
      //alimentando celulas
      for r := 1 to y do begin
        try
//          AGrid.Cells[r,I] := Trim(Sheet.Cells[I+1, r+1].Text);
          if VarType(RangeMatrix[K, R]) <> varError then begin
            AGrid.Cells[(r - 1),(k - 1)] := RangeMatrix[K, R];
          end;
        except
          on E: Exception do begin
            ShowMessage('Erro na linha e coluna: ' + IntToStr(r)  + '/' + IntToStr(k) + #13#13 + E.Message + #13#13 + E.ClassName);
          end;
        end;
      end;

      Inc(k,1);

      if pProgressBar <> nil then begin
        pProgressBar.Position := pProgressBar.Position + 1;
      end;
    until k > x;
//    Application.ProcessMessages;
//    RangeMatrix := Unassigned;

    if pProgressBar <> nil then begin
      pProgressBar.Max      := AGrid.RowCount;
      pProgressBar.Position := 0;
    end;

  finally
//    Application.ProcessMessages;
    if pProgressBar <> nil then begin
      pProgressBar.Position := 0;
    end;

    //Fecha o Excel
    if not VarIsEmpty(XLApp) then begin
      XLApp.Quit;
      XLAPP:=Unassigned;
      Sheet:=Unassigned;
      Result:=True;
    end;
	end;
end;

procedure TImportExcel.ExportarParaExcel(const pDataSet: TDataSet; pCaption: string; pFieldsPref: array of string);
var linha, coluna : integer;
  I: Integer;
var planilha : variant;
var valorcampo : string;
begin
  planilha:= CreateoleObject('Excel.Application');
  planilha.WorkBooks.add(1);
  planilha.caption := pCaption;
  planilha.visible := False;

  pDataSet.Open;
  pDataSet.First;

  try
    pDataSet.DisableControls;
    //cabeçalho
    for linha := 0 to pDataSet.RecordCount - 1 do begin
      Application.ProcessMessages;
      for coluna := 1 to pDataSet.FieldCount do begin
        valorcampo := pDataSet.Fields[coluna - 1].AsString;
        planilha.cells[linha + 2,coluna] := valorCampo;
      end;
      pDataSet.Next;
    end;

    //linhas
    for coluna := 1 to pDataSet.FieldCount do begin
      if Length(pFieldsPref) > 0 then begin
        for I := Low(pFieldsPref) to High(pFieldsPref) do begin
          if pDataSet.Fields[coluna-1].FieldName = pFieldsPref[I] then begin
            valorcampo := pDataSet.Fields[coluna - 1].DisplayLabel;

            planilha.cells[1,coluna] := valorcampo;
          end;
        end;
      end
      else begin
        valorcampo := pDataSet.Fields[coluna - 1].DisplayLabel;
        planilha.cells[1,coluna] := valorcampo;
      end;
    end;
    planilha.columns.Autofit;
  finally
    pDataSet.EnableControls;

    ShowMessage('Pronto. O arquivo foi criado com sucesso. Por favor verifique o arquivo Excel que está aberto.');
  end;
  planilha.visible := True;
end;

procedure TImportExcel.Importar;
var
  planilha, sheet: OleVariant;
  linha, coluna: Integer;
begin
  //Crio o objeto que gerencia o arquivo excel
  planilha:= CreateOleObject('Excel.Application');

  //Abro o arquivo
  planilha.WorkBooks.open('c:\nome_da_planilha.xls');

  //Pego a primeira planilha do arquivo
  sheet:= planilha.WorkSheets[1];


  //Aqui pego o texto de uma das células
  linha:= 0;
  coluna:= 0;
  ShowMessage(sheet.cells[linha, coluna].Text);


  //Fecho a planilha
  planilha.WorkBooks.Close;
end;

end.
