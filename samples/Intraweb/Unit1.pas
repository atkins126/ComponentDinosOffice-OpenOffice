unit Unit1;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, Vcl.Controls, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompButton,

  uOpenOffice,
  uOpenOffice_calc,
  uOpenOfficeHelper,
  uOpenOfficeCollors, uOpenOffice_writer;

type
  TIWForm1 = class(TIWAppForm)
    iwbtn1: TIWButton;
    OpenOffice_calc1: TOpenOffice_calc;
    iwbtn2: TIWButton;
    iwbtn3: TIWButton;
    iwbtn4: TIWButton;
    iwbtn5: TIWButton;
    iwbtn6: TIWButton;
    OpenOffice_writer1: TOpenOffice_writer;
    procedure iwbtn1Click(Sender: TObject);
    procedure iwbtn2Click(Sender: TObject);
    procedure iwbtn3Click(Sender: TObject);
    procedure iwbtn4Click(Sender: TObject);
    procedure iwbtn6Click(Sender: TObject);
    procedure iwbtn5Click(Sender: TObject);
  public
  end;

implementation

{$R *.dfm}
uses ServerController;

procedure TIWForm1.iwbtn1Click(Sender: TObject);
var SettingsChart: TSettingsChart;
begin
  OpenOffice_calc1.DocVisible := true;
  OpenOffice_calc1.startSheet;

  OpenOffice_calc1.SetValue(1, 'A', 'STATUS')
   .SetBorder([bAll], opBrown)
   .changeJustify(fthRIGHT, ftvTOP)
   .setBold(true)
   .changeFont('Arial', 12)
   .SetUnderline(true)
   .setColor(opWhite, opMagenta);

  OpenOffice_calc1.SetValue(1, 'B', 'VALOR').changeJustify(fthRIGHT, ftvTOP)
   .SetBorder([bAll], opBrown)
   .setBold(true)
   .changeFont('Arial', 12)
   .SetUnderline(true)
   .setColor(opWhite, opMagenta);

  OpenOffice_calc1.SetValue(2, 'B', 109, ftNumeric).SetBorder([bAll], opBrown);
  OpenOffice_calc1.SetValue(2, 'A', 'AGUA').SetBorder([bAll], opBrown);

  OpenOffice_calc1.SetValue(3, 'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
  OpenOffice_calc1.SetValue(3, 'A', 'LUZ').SetBorder([bAll], opBrown);

  OpenOffice_calc1.SetValue(4, 'B', 1005.22, ftNumeric);
  OpenOffice_calc1.SetValue(4, 'A', 'ALUGUEL');

  OpenOffice_calc1.SetValue(6, 'A', 'Total de linhas');
  OpenOffice_calc1.SetValue(6, 'B', OpenOffice_calc1.CountRow, ftNumeric);

  OpenOffice_calc1.SetValue(7, 'A', 'Total de Colunas');
  OpenOffice_calc1.SetValue(7, 'B', OpenOffice_calc1.CountCell, ftNumeric);

  OpenOffice_calc1.addNewSheet('A Receber', 1);

  OpenOffice_calc1.SetValue(1, 'A', 'VALOR')
   .SetBorder([bAll], opBrown)
   .changeJustify(fthRIGHT, ftvTOP)
   .setBold(true);

  OpenOffice_calc1.SetValue(1, 'B', 'DESC')
   .SetBorder([bAll], opBrown)
   .changeJustify(fthRIGHT, ftvTOP)
   .setBold(true)
   .changeFont('Arial', 12)
   .SetUnderline(true)
   .setColor(opWhite, opCiano);

  OpenOffice_calc1.SetValue(1, 'C', 'SOMA')
   .SetBorder([bAll], opBrown)
   .changeJustify(fthRIGHT, ftvTOP)
   .setBold(true)
   .changeFont('Arial', 12)
   .SetUnderline(true)
   .setColor(opWhite, opSoftRed);

  OpenOffice_calc1.SetValue(1, 'H', 'SOMA')
   .SetBorder([bAll], opBrown)
   .changeJustify(fthRIGHT, ftvTOP)
   .setBold(true)
   .changeFont('Arial', 12)
   .SetUnderline(true)
   .setColor(opWhite, opSoftRed);

  OpenOffice_calc1.SetValue(2, 'A', 200, ftNumeric);
  OpenOffice_calc1.SetValue(2, 'B', 'Emprestimo');
  OpenOffice_calc1.SetValue(2, 'C', 0, ftNumeric);

  OpenOffice_calc1.SetValue(3, 'A', 369.55, ftNumeric);
  OpenOffice_calc1.SetValue(3, 'B', 'Dividendos');
  OpenOffice_calc1.SetValue(3, 'C', 0, ftNumeric);

  OpenOffice_calc1.SetValue(4, 'A', 1585.22, ftNumeric);
  OpenOffice_calc1.SetValue(4, 'B', 'ALUGUEL');
  OpenOffice_calc1.SetValue(4, 'C', 0, ftNumeric);

  OpenOffice_calc1.SetValue(8, 'A', 1585.22, ftNumeric);
  OpenOffice_calc1.SetValue(8, 'B', 'Renda extra');
  OpenOffice_calc1.SetValue(8, 'C', 0, ftNumeric);

  OpenOffice_calc1.SetValue(15, 'A', 1585.22, ftNumeric);
  OpenOffice_calc1.SetValue(15, 'B', 'ALUGUEL 2');
  OpenOffice_calc1.SetValue(15, 'C', 0, ftNumeric);

  OpenOffice_calc1.SetValue(17, 'A', 'Total de linhas');
  OpenOffice_calc1.SetValue(17, 'B', OpenOffice_calc1.CountRow, ftNumeric);

  OpenOffice_calc1.SetValue(19, 'A', 'Total de Colunas');
  OpenOffice_calc1.SetValue(19, 'B', OpenOffice_calc1.CountCell, ftNumeric);
  OpenOffice_calc1.setFormula(20, 'A', '=A2+A3+A4+A15')
   .setBold(true);

  OpenOffice_calc1.positionSheetByName('Planilha1');

  //Configure the chart settings
  SettingsChart.Height := 11000;
  SettingsChart.Width := 22000;
  SettingsChart.Position_X := 1500;
  SettingsChart.Position_Y := 5000;
  SettingsChart.StartRow := 0;
  SettingsChart.EndRow := 3;
  SettingsChart.PositionSheet := 0; //first tab
  SettingsChart.StartColumn := 'A';
  SettingsChart.EndColumn := 'B';
  SettingsChart.ChartName := 'TestChart';
  SettingsChart.typeChart := ctDefault;

  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctVertical;
  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctPie;
  OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctLine;
  OpenOffice_calc1.addChart(SettingsChart);

  sleep(3000);
  OpenOffice_calc1.saveFile( IWServerController.ContentPath  +  'sheet.xls');
  OpenOffice_calc1.CloseFile;

end;

procedure TIWForm1.iwbtn2Click(Sender: TObject);
begin
  if FileExists(IWServerController.ContentPath  +  'sheet.xls') then
  begin
    OpenOffice_calc1.DocVisible  := true;
    OpenOffice_calc1.URlFile     := IWServerController.ContentPath  +  'sheet.xls';
    OpenOffice_calc1.SheetName   := 'Planilha1';
    OpenOffice_calc1.startSheet;
  end
  else
   raise Exception.Create('File not exists');
end;

procedure TIWForm1.iwbtn3Click(Sender: TObject);
begin
  OpenOffice_calc1.CloseFile;
end;

procedure TIWForm1.iwbtn4Click(Sender: TObject);
begin
  OpenOffice_writer1.DocVisible := true;
  OpenOffice_writer1.startDoc;
  OpenOffice_writer1.setBold(true)
    .setFontHeight(16)
    .setValue('T�tulo: Apresentando o componente Libre Office writer via Delphi <3' + #13#13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .setValue('Neste exemplo estou mostrando a cria��o de documentos via c�digo, de um jeito simples, r�pido e f�cil.' + #13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .setValue('Espero que seja �til e que estejam gostando, lembrando o componente � open source e totalmente free!' + #13#13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .SetBold(true)
    .setValue('Obrigado a todos pela presen�a!!!' + #13);
  OpenOffice_writer1.gotoEndOfSentence;

  sleep(3000);
  OpenOffice_writer1.saveFile( IWServerController.ContentPath  + 'doc.docx');
  OpenOffice_writer1.CloseFile;
end;

procedure TIWForm1.iwbtn5Click(Sender: TObject);
begin
  OpenOffice_writer1.DocVisible  := true;
  OpenOffice_writer1.URlFile     := IWServerController.ContentPath  + 'doc.docx';
  OpenOffice_writer1.startDoc;
end;

procedure TIWForm1.iwbtn6Click(Sender: TObject);
begin
  OpenOffice_writer1.CloseFile;
end;

initialization
  TIWForm1.SetAsMainForm;

end.
