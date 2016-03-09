unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, mbTBXComboEdit, mbTBXHint, 
  ExtCtrls, SUIForm, mbTBXGroupBox, Menus, SUIMainMenu, ComCtrls,
  SUIStatusBar, JvComponent, JvMagnet,
  AdvPageControl, SUIPageControl,
  SUITabControl, IniFiles, uPseParser, FileCtrl, u_dzCustomCharts,
  u_dzXYChart, SUIListBox, SUIImagePanel, SUIButton, SUIComboBox, SUIEdit,
  SUIMemo, Mask, mbTBXEdit, TBXDkPanels, mbTBXFixedStuff, mbTBXComboBox,
  AppEvnts, mbTBXMemo, jpeg, SUIPopupMenu, mbTBXPanel, Grids,
  mbTBXDrawGrid, mbTBXStringGrid, JvLabel, JvDateTimePicker, uPSEDataProvider,
  SUIGroupBox;

type
  TfrmMain = class(TForm)
    suiForm1: TsuiForm;
    mbTBXHint1: TmbTBXHint;
    JvFormMagnet1: TJvFormMagnet;
    PageControl: TsuiPageControl;
    tabConverter: TsuiTabSheet;
    tabNetForeign: TsuiTabSheet;
    tabAbout: TsuiTabSheet;
    mbTBXGroupBox1: TmbTBXGroupBox;
    edFile: TmbTBXComboEdit;
    mbTBXGroupBox2: TmbTBXGroupBox;
    edOutput: TmbTBXComboEdit;
    mbTBXLabel1: TmbTBXLabel;
    mbTBXLabel2: TmbTBXLabel;
    mbTBXLabel3: TmbTBXLabel;
    mbTBXLabel4: TmbTBXLabel;
    OpenDialog: TOpenDialog;
    Image1: TImage;
    mbTBXLabel5: TmbTBXLabel;
    suiPanel1: TsuiPanel;
    suiListBox1: TsuiListBox;
    suiPanel2: TsuiPanel;
    suiListBox2: TsuiListBox;
    dzXYChart1: TdzXYChart;
    dzXYChart2: TdzXYChart;
    TBXButton1: TTBXButton;
    TBXButton2: TTBXButton;
    TBXButton3: TTBXButton;
    TBXLink1: TTBXLink;
    TBXLink2: TTBXLink;
    mbTBXLabel6: TmbTBXLabel;
    mbTBXLabel7: TmbTBXLabel;
    mbTBXComboBox1: TmbTBXComboBox;
    btOk: TsuiButton;
    suiButton1: TsuiButton;
    cb1: TsuiComboBox;
    cb2: TsuiComboBox;
    cb3: TsuiComboBox;
    cb4: TsuiComboBox;
    cb5: TsuiComboBox;
    cb6: TsuiComboBox;
    cb7: TsuiComboBox;
    cb8: TsuiComboBox;
    rbInternet: TsuiRadioButton;
    rbFile: TsuiRadioButton;
    rbPaste: TsuiRadioButton;
    cbInternet: TsuiComboBox;
    edDateFormat: TsuiEdit;
    mmReport: TsuiMemo;
    edDelimiter: TsuiEdit;
    StatusBar: TsuiStatusBar;
    mbTBXMemo1: TmbTBXMemo;
    tabMarket: TsuiTabSheet;
    cbDate: TJvDateTimePicker;
    JvLabel16: TJvLabel;
    suiButton3: TsuiButton;
    AdvPageControl1: TAdvPageControl;
    tabGainers: TAdvTabSheet;
    suiPanel5: TsuiPanel;
    JvLabel1: TJvLabel;
    JvLabel2: TJvLabel;
    JvLabel3: TJvLabel;
    JvLabel5: TJvLabel;
    JvLabel26: TJvLabel;
    grdGainers: TmbTBXStringGrid;
    tabLosers: TAdvTabSheet;
    suiPanel6: TsuiPanel;
    JvLabel4: TJvLabel;
    JvLabel6: TJvLabel;
    JvLabel7: TJvLabel;
    JvLabel8: TJvLabel;
    JvLabel9: TJvLabel;
    grdLosers: TmbTBXStringGrid;
    tabActive: TAdvTabSheet;
    mbTBXPanel3: TmbTBXPanel;
    JvLabel11: TJvLabel;
    JvLabel12: TJvLabel;
    JvLabel13: TJvLabel;
    JvLabel14: TJvLabel;
    JvLabel15: TJvLabel;
    mbTBXStringGrid4: TmbTBXStringGrid;
    tabBuy: TAdvTabSheet;
    suiPanel7: TsuiPanel;
    JvLabel17: TJvLabel;
    JvLabel18: TJvLabel;
    JvLabel19: TJvLabel;
    JvLabel21: TJvLabel;
    JvLabel27: TJvLabel;
    grdNetBuy: TmbTBXStringGrid;
    tabSell: TAdvTabSheet;
    suiPanel8: TsuiPanel;
    JvLabel10: TJvLabel;
    JvLabel20: TJvLabel;
    JvLabel22: TJvLabel;
    JvLabel23: TJvLabel;
    JvLabel24: TJvLabel;
    grdSell: TmbTBXStringGrid;
    tabActivity: TAdvTabSheet;
    grdMisc: TmbTBXStringGrid;
    mbTBXHint2: TmbTBXHint;
    menuRangeMode: TsuiPopupMenu;
    Daily1: TMenuItem;
    Weekly1: TMenuItem;
    Monthly1: TMenuItem;
    N1: TMenuItem;
    Ontop: TMenuItem;
    chkVolume: TsuiCheckBox;
    suiPanel4: TsuiPanel;
    JvLabel28: TJvLabel;
    JvLabel29: TJvLabel;
    JvLabel30: TJvLabel;
    JvLabel32: TJvLabel;
    grdMarket: TmbTBXStringGrid;
    suiTabSheet1: TsuiTabSheet;
    suiGroupBox1: TsuiGroupBox;
    dtDatePurge: TJvDateTimePicker;
    btnPurge: TsuiButton;
    procedure rbInternetChange(Sender: TObject);
    procedure rbFileChange(Sender: TObject);
    procedure rbPasteChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tabConverterShow(Sender: TObject);
    procedure tabNetForeignShow(Sender: TObject);
    procedure mmReportChange(Sender: TObject);
    procedure edFileButtonClick(Sender: TObject);
    procedure edDateFormatChange(Sender: TObject);
    procedure edOutputChange(Sender: TObject);
    procedure cb1Change(Sender: TObject);
    procedure edOutputButtonClick(Sender: TObject);
    procedure btOkClick(Sender: TObject);
    procedure suiButton1Click(Sender: TObject);
    procedure FormMinimize(Sender: TObject);
    procedure grdMarketDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure grdGainersDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure suiButton3Click(Sender: TObject);
    procedure OntopClick(Sender: TObject);
    procedure tabMarketShow(Sender: TObject);
    procedure suiTabSheet1Show(Sender: TObject);
    procedure btnPurgeClick(Sender: TObject);
  private
    { Private declarations }
    DataFormat: TStringList;
    procedure ParamComplete;
    procedure LoadIndeces;
    procedure LoadTopGainers;
    procedure LoadTopLosers;
    procedure LoadTopForeignBuy;
    procedure LoadTopForeignSell;
    procedure LoadMisc;
  public
    { Public declarations }
    PSE: TPSECollection;
    PSEDataProvider: TPSEDataProvider;
    procedure LoadData;
  end;

var
  frmMain: TfrmMain;

implementation

uses httpsend, SQLiteTable3, DateUtils, uExNotice;

{$R *.dfm}

const
  MAX_ROWS = 10;
  FONT_SIZE = 8;

procedure TfrmMain.rbInternetChange(Sender: TObject);
begin
  cbInternet.Enabled := true;
  cbInternet.Color := clWindow;
  edFile.Enabled := false;
  edFile.Color := clBtnFace;
  btOk.Enabled := true;
  mmReport.Enabled := false;
  mmReport.Color := clBtnFace;
  StatusBar.SimpleText := '';
end;

procedure TfrmMain.rbFileChange(Sender: TObject);
begin
  cbInternet.Enabled := false;
  cbInternet.Color := clBtnFace;
  edFile.Enabled := true;
  edFile.Color := clWindow;
  edOutput.Enabled := true;
  edOutput.Color := clWindow;
  mmReport.Enabled := false;
  mmReport.Color := clBtnFace;
  StatusBar.SimpleText := '';  
end;

procedure TfrmMain.rbPasteChange(Sender: TObject);
begin
  cbInternet.Enabled := false;
  cbInternet.Color := clBtnFace;
  edFile.Enabled := false;
  edFile.Color := clBtnFace;
  mmReport.Lines.Clear;
  mmReport.Enabled := true;
  mmReport.Color := clWindow;
  StatusBar.SimpleText := '';  
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  DataFormat := TStringList.Create;
  DataFormat.Add('S=SYMBOL');
  DataFormat.Add('D=SYMBOL');
  DataFormat.Add('O=OPEN');
  DataFormat.Add('H=HIGH');
  DataFormat.Add('L=LOW');
  DataFormat.Add('C=CLOSE');
  DataFormat.Add('V=VOLUME');
  DataFormat.Add('I=OPEN INT.');
  PageControl.ActivePageIndex := 0;
  btOk.Enabled := false;

  PSE := TPSECollection.Create;
  PSEDataProvider := TPSEDataProvider.Create;
  PSEDataProvider.SQLiteDatabase := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'\pseget.dat');

  Application.OnMinimize := FormMinimize;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
var
  IniFile: TIniFile;
begin
  DataFormat.Free;
  PSE.Free;
  PSEDataProvider.SQLiteDatabase.Free;
  PSEDataProvider.Free;
  IniFile := TIniFile.Create(ChangeFileExt(Application.ExeName, '.INI'));
  try
    IniFile.WriteInteger('form', 'left', Left);
    IniFile.WriteInteger('form', 'top', Top);
    if Ontop.Checked then
      IniFile.WriteString('form', 'ontop', 'y')
    else
      IniFile.WriteString('form', 'ontop', 'n');
  finally
    IniFile.Free;
  end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
  IniFile: TIniFile;
  StrArr: TStrArray;
  i: SmallInt;
  ArrCombo: array[0..7] of TCustomComboBox;
  ReportSource: ShortInt;
begin
  IniFile := TIniFile.Create(ChangeFileExt(Application.ExeName, '.INI'));
  try
    Left := IniFile.ReadInteger('form', 'left', 183);
    Top := IniFile.ReadInteger('form', 'top', 137);
    if LowerCase(IniFile.ReadString('form', 'ontop', 'n')) = 'y' then
      OntopClick(Self);
    edOutput.Text := IniFile.ReadString('options', 'output_dir', ExtractFilePath(Application.ExeName));
    edDateFormat.Text := IniFile.ReadString('options', 'date_format', 'MM/DD/YYYY');
    edDelimiter.Text := IniFile.ReadString('options', 'delimiter', ',');
    StrArr := Explode(edDelimiter.Text, IniFile.ReadString('options', 'data_format', 'S,D,O,H,L,C,V,I'));
    ReportSource := IniFile.ReadInteger('options', 'report_source', 3);
    case ReportSource of
      1: rbInternet.Checked := true;
      2: rbFile.Checked := true;
      else
        rbPaste.Checked := true;
    end;
    ArrCombo[0] := cb1;
    ArrCombo[1] := cb2;
    ArrCombo[2] := cb3;
    ArrCombo[3] := cb4;
    ArrCombo[4] := cb5;
    ArrCombo[5] := cb6;
    ArrCombo[6] := cb7;
    ArrCombo[7] := cb8;
    for i := 0 to High(ArrCombo) do
      ArrCombo[i].ItemIndex := 8;

    for i := 0 to High(StrArr) do
      if DataFormat.IndexOfName(StrArr[i]) >= 0 then
        ArrCombo[i].ItemIndex := DataFormat.IndexOfName(StrArr[i]);

    chkVolume.Checked := LowerCase(IniFile.ReadString('options', 'ivalue_as_volume', 'y')) = 'y';

  finally
    IniFile.Free;
  end;
end;

procedure TfrmMain.tabConverterShow(Sender: TObject);
begin
  Width := 442;
  //Height := 558;
end;

procedure TfrmMain.tabNetForeignShow(Sender: TObject);
begin
  Width := 571;
end;

procedure TfrmMain.mmReportChange(Sender: TObject);
begin
  if Trim(mmReport.Text) <> '' then
    btOk.Enabled := true
  else
    btOk.Enabled := false;
  StatusBar.SimpleText := '';    
end;

procedure TfrmMain.edFileButtonClick(Sender: TObject);
var
  i: Integer;
begin
  if OpenDialog.Execute then
  begin
    for i := 0 to OpenDialog.Files.Count - 1 do
      if i = 0 then
        edFile.Text := OpenDialog.Files[i]
      else
        edFile.Text := edFile.Text+','+OpenDialog.Files[i];
    ParamComplete;
  end;
end;

procedure TfrmMain.ParamComplete;
begin
  if (edFile.Text <> '') then
  begin
    btOk.Enabled := true;
  end else
  begin
    btOk.Enabled := false;
  end;
end;

procedure TfrmMain.edDateFormatChange(Sender: TObject);
begin
  StatusBar.SimpleText := '';
end;

procedure TfrmMain.edOutputChange(Sender: TObject);
begin
  StatusBar.SimpleText := '';
end;

procedure TfrmMain.cb1Change(Sender: TObject);
begin
  StatusBar.SimpleText := '';
end;

procedure TfrmMain.edOutputButtonClick(Sender: TObject);
var
 ChosenDirectory: String;
begin

  ChosenDirectory := ExtractFilePath(Application.ExeName);
  if SelectDirectory('Select output directory',
                    'C:', ChosenDirectory) then

  begin

    edOutput.Text := ChosenDirectory;
    if rbFile.Checked then
      ParamComplete;
  end;
end;

procedure TfrmMain.btOkClick(Sender: TObject);
var
  i: integer;
  OutFile: String;
  AHttpSend: THTTPSend;
  StrArray: TStrArray;
  IniFile: TIniFile;
  OutFormat: String;
  PSEDataProvider: TPSEDataProvider;
  DB: TSQLiteDatabase;
const
  DirSepChar = '\';
  function GetDataFormat: String;
  begin
    if cb1.Text <> '' then
      Result := DataFormat.Names[cb1.ItemIndex];

    if cb2.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb2.ItemIndex]
    end;

    if cb3.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb3.ItemIndex];
    end;

    if cb4.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb4.ItemIndex];
    end;

    if cb5.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb5.ItemIndex];
    end;

    if cb6.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb6.ItemIndex];
    end;
    if cb7.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb7.ItemIndex];
    end;
    if cb8.Text <> '' then
    begin
      if Result <> '' then
        Result := Result + edDelimiter.Text;
      Result := Result + DataFormat.Names[cb8.ItemIndex];
    end;
    if Result = '' then
      Result := 'S,D,O,H,L,C,V,I'

  end;
begin
  IniFile := TIniFile.Create(ChangeFileExt(Application.ExeName, '.INI'));
  DB := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'\pseget.dat');
  PSEDataProvider := TPSEDataProvider.Create;
  PSEDataProvider.SQLiteDatabase := DB;
  Screen.Cursor := crHourGlass;
  PSE.Clear;
  try
    OutFormat := GetDataFormat;
    if rbFile.Checked then
    begin
      if (edFile.Text = '') then
      begin
        Screen.Cursor := crDefault;
        ShowMessage('Input file please.');
        Exit;
      end;
      if edDelimiter.Text = '' then
        edDelimiter.Text := ' ';
      StrArray := Explode(',', edFile.Text);
      for i := 0 to High(StrArray) do
      begin
        StatusBar.SimpleText := 'Converting file ' + StrArray[i];
        PSE.Clear;
        PSE.QuoteDocument.LoadFromFile(StrArray[i]);
        PSE.ParseDocument;
        OutFile := ChangeFileExt(StrArray[i], '.csv');
        if Trim(edOutput.Text) = '' then
          PSE.SaveToFile(OutFile, OutFormat, Trim(edDateFormat.Text), chkVolume.Checked)
        else
          PSE.SaveToFile(edOutput.Text+DirSepChar+ExtractFileName(OutFile), OutFormat, Trim(edDateFormat.Text), chkVolume.Checked);
        PSEDataProvider.AppendPSETrade(PSE);
        Application.ProcessMessages;
      end;
      Screen.Cursor := crDefault;
      StatusBar.SimpleText := Format('Successfully converted %d PSE report file(s)', [High(StrArray)+1]);
    end else if rbInternet.Checked then
    begin
      AHttpSend := THTTPSend.Create;
      try
        StatusBar.SimpleText := 'Downloading...please wait...';
        AHttpSend.HTTPMethod('GET', cbInternet.Text);
        PSE.QuoteDocument.LoadFromStream(AHttpSend.Document);
        if PSE.QuoteDocument.Count = 0 then
          ShowMessage('Document is empty!')
        else if PSE.QuoteDocument.Count < 10 then
          ShowMessage(PSE.QuoteDocument.Text)
        else
        begin
          if PSE.QuoteDocument[0] = '<pre>' then
          begin
            PSE.QuoteDocument.Delete(0);
            PSE.QuoteDocument.Delete(PSE.QuoteDocument.Count - 1);
          end;
          PSE.ParseDocument;
          Outfile := 'pseget.out.csv';
          if Trim(edOutput.Text) = '' then
            PSE.SaveToFile(ExtractFilePath(Application.ExeName)+DirSepChar+ OutFile, OutFormat, Trim(edDateFormat.Text), chkVolume.checked)
          else
            PSE.SaveToFile(edOutput.Text+DirSepChar+OutFile, OutFormat, Trim(edDateFormat.Text), chkVolume.Checked);
          Screen.Cursor := crDefault;
          StatusBar.SimpleText := 'Download complete.';
          PSEDataProvider.AppendPSETrade(PSE);
        end;
      finally
        AHttpSend.Free;
        Screen.Cursor := crDefault;
      end;

    end else
    begin
      PSE.QuoteDocument.Assign(mmReport.Lines);
      PSE.ParseDocument;
      Outfile := 'pseget.out.csv';
      if Trim(edOutput.Text) = '' then
        PSE.SaveToFile(ExtractFilePath(Application.ExeName)+DirSepChar+ OutFile, OutFormat, Trim(edDateFormat.Text), chkVolume.Checked)
      else
        PSE.SaveToFile(edOutput.Text+DirSepChar+OutFile, OutFormat, Trim(edDateFormat.Text), chkVolume.Checked);
      Screen.Cursor := crDefault;
      StatusBar.SimpleText := 'Conversion complete.';
      PSEDataProvider.AppendPSETrade(PSE);
    end;

    if PSE.Count > 0 then
    begin
      cbDate.DateTime := PSE.TradeDate;
      LoadData;    
      PageControl.ActivePage := tabMarket;
      frmNotice.Show;
      frmNotice.mmNotice.Text := PSE.ExchangeNotice.Text;
    end;

    IniFile.WriteString('options', 'output_dir', edOutput.Text);
    IniFile.WriteString('options', 'data_format', OutFormat);
    IniFile.WriteString('options', 'date_format', Trim(edDateFormat.Text));
    IniFile.WriteString('options', 'delimiter', edDelimiter.Text);
    if rbInternet.Checked then
      IniFile.WriteInteger('options', 'report_source', 1)
    else if rbFile.Checked then
      IniFile.WriteInteger('options', 'report_source', 2)
    else
      IniFile.WriteInteger('options', 'report_source', 3);
    if chkVolume.Checked then
      IniFile.WriteString('options', 'ivalue_as_volume', 'y')
    else
      IniFile.WriteString('options', 'ivalue_as_volume', 'n');
  finally
    IniFile.Free;
    Screen.Cursor := crDefault;
    DB.Free;
    PSEDataProvider.Free;
  end;
end;

procedure TfrmMain.suiButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.FormMinimize(Sender: TObject);
begin
  frmNotice.WindowState := wsMinimized;
end;

procedure TfrmMain.LoadIndeces;
var
  Dataset: TSQLiteTable;
  SQL: String;
  PreviousClose, PointsGain, PercentChg: Double;
begin
  SQL := 'SELECT * FROM TRADETAB WHERE SYMBOL = ''^PHISIX'' AND' +
         ' DATE = ' + IntToStr(DateTimeToUnix(cbDate.Date));
  Dataset := PSEDataProvider.GetData(SQL);
  grdMarket.RowCount := 0;
  grdMarket.Rows[0].Clear;
  grdMarket.ColWidths[0] := 90;
  grdMarket.ColWidths[1] := 65;
  grdMarket.ColWidths[2] := 60;
  grdMarket.ColWidths[3] := 50;
  grdMarket.ColWidths[4] := 80;
  try
    if Dataset.RowCount > 0 then
    begin
      with grdMarket do
      begin
        PreviousClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
        PercentChg := GetPercentChg(StrToFloat(Dataset.FieldByName['CLOSE']), PreviousClose);
        PointsGain := StrToFloat(Dataset.FieldByName['CLOSE']) - PreviousClose;

        Rows[0].Append(Dataset.FieldByName['SYMBOL']);
        Rows[0].Append(Dataset.FieldByName['CLOSE']);
        Rows[0].Append(FloatToStr(PointsGain));
        Rows[0].Append(FloatToStr(PercentChg));
        Rows[0].Append(Dataset.FieldByName['OPENINTEREST']);
      end;

      Dataset := PSEDataProvider.GetIndeces(cbDate.Date);
      while not Dataset.EOF do
      begin
        with grdMarket do
        begin
          PreviousClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
          PercentChg := GetPercentChg(StrToFloat(Dataset.FieldByName['CLOSE']), PreviousClose);
          PointsGain := StrToFloat(Dataset.FieldByName['CLOSE']) - PreviousClose;

          RowCount := RowCount + 1;
          Rows[RowCount - 1].Clear;
          Rows[RowCount - 1].Append(Dataset.FieldByName['SYMBOL']);
          Rows[RowCount - 1].Append(Dataset.FieldByName['CLOSE']);
          Rows[RowCount - 1].Append(FloatToStr(PointsGain));
          Rows[RowCount - 1].Append(FloatToStr(PercentChg));
          Rows[RowCount - 1].Append(Dataset.FieldByName['OPENINTEREST']);
        end;
        Dataset.Next;
      end;
    end;
  finally
    Dataset.Free;
  end;
end;

procedure TfrmMain.LoadMisc;
var
  Dataset: TSQLiteTable;
begin
  Dataset := PSEDataProvider.GetMiscData(cbDate.Date, MAX_ROWS);
  try
    grdMisc.RowCount := 0;
    grdMisc.Rows[0].Clear;
    grdMisc.ColWidths[0] := 150;
    grdMisc.ColWidths[1] := 150;
    if Dataset.RowCount > 0 then
    begin
      with grdMisc do
      begin
        RowCount := 11;
        Rows[0].Clear;
        Rows[0].Append('Advances');
        Rows[0].Append(Dataset.FieldByName['ADVANCE']);

        Rows[1].Clear;
        Rows[1].Append('Declines');
        Rows[1].Append(Dataset.FieldByName['DECLINE']);

        Rows[2].Clear;
        Rows[2].Append('Unchanged');
        Rows[2].Append(Dataset.FieldByName['UNCHANGED']);

        Rows[3].Clear;
        Rows[3].Append('Traded Issues');
        Rows[3].Append(Dataset.FieldByName['TRADEDISSUES']);

        Rows[4].Clear;
        Rows[4].Append('Number Of Trades');
        Rows[4].Append(Dataset.FieldByName['NUMTRADES']);

        Rows[6].Clear;
        Rows[6].Append('Composite');
        Rows[6].Append(Dataset.FieldByName['VALUE']);

        Rows[7].Clear;
        Rows[7].Append('Total Foreign Buying');
        Rows[7].Append(Dataset.FieldByName['FOREIGNBUY']);

        Rows[8].Clear;
        Rows[8].Append('Total Foreign Selling');
        Rows[8].Append(Dataset.FieldByName['FOREIGNSELL']);

        Rows[9].Clear;
        Rows[9].Append('Main Board Cross Volume');
        Rows[9].Append(Dataset.FieldByName['CROSSVOLUME']);

        Rows[10].Clear;
        Rows[10].Append('Main Board Cross Value');
        Rows[10].Append(Dataset.FieldByName['CROSSVALUE']);
      end;
    end;
  finally
    Dataset.Free;
  end;
end;

procedure TfrmMain.LoadTopForeignBuy;
var
  Dataset: TSQLiteTable;
  CurrentRow: Byte;
  StrValue: String;
  PrevClose: Double;
begin
  Dataset := PSEDataProvider.GetTopForeignBuy(cbDate.Date, MAX_ROWS);
  try
    grdNetBuy.RowCount := 0;
    grdNetBuy.Rows[0].Clear;
    grdNetBuy.ColWidths[0] := 40;
    grdNetBuy.ColWidths[1] := 75;
    grdNetBuy.ColWidths[3] := 75;
    grdNetBuy.ColWidths[4] := 75;
    grdNetBuy.ColWidths[5] := 0;

    if Dataset.RowCount > 0 then
    begin
      while not Dataset.EOF do
      begin
        with grdNetBuy do
        begin
          CurrentRow := RowCount - 1;
          StrValue := Dataset.FieldAsString(Dataset.FieldIndex['VALUE']);
          PrevClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
          Rows[CurrentRow].Clear;
          Rows[CurrentRow].Add(Dataset.FieldByName['SYMBOL']);
          Rows[CurrentRow].Add(Dataset.FieldByName['VOLUME']);
          Rows[CurrentRow].Add(FloatToStr(Trunc(StrToFloat(StrValue) / 1000)));
          Rows[CurrentRow].Add(Dataset.FieldByName['CLOSE']);
          Rows[CurrentRow].Add(FloatToStr(StrToFloat(Dataset.FieldByName['OPENINTEREST']) / 1000000));
          Rows[CurrentRow].Add(FloatToStr(PrevClose));
          if RowCount < 10 then
            RowCount := RowCount + 1;
        end;
        Dataset.Next;
      end;
    end;
  finally
    Dataset.Free;
  end;

end;


procedure TfrmMain.LoadTopForeignSell;
var
  Dataset: TSQLiteTable;
  CurrentRow: Byte;
  StrValue: String;
  PrevClose: Double;
begin
  Dataset := PSEDataProvider.GetTopForeignSell(cbDate.Date, MAX_ROWS);
  try
    grdSell.RowCount := 0;
    grdSell.Rows[0].Clear;
    grdSell.ColWidths[0] := 40;
    grdSell.ColWidths[1] := 75;
    grdSell.ColWidths[3] := 75;
    grdSell.ColWidths[4] := 75;
    grdSell.ColWidths[5] := 0;
    if Dataset.RowCount > 0 then
    begin
      while not Dataset.EOF do
      begin
        with grdSell do
        begin
          CurrentRow := RowCount - 1;
          StrValue := Dataset.FieldAsString(Dataset.FieldIndex['VALUE']);
          PrevClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
          Rows[CurrentRow].Clear;
          Rows[CurrentRow].Append(Dataset.FieldByName['SYMBOL']);
          Rows[CurrentRow].Append(Dataset.FieldByName['VOLUME']);
          Rows[CurrentRow].Append(FloatToStr(Trunc(StrToFloat(StrValue) / 1000)));
          Rows[CurrentRow].Append(Dataset.FieldByName['CLOSE']);
          Rows[CurrentRow].Append(FloatToStr(StrToFloat(Dataset.FieldByName['OPENINTEREST']) / 1000000));
          Rows[CurrentRow].Append(FloatToStr(PrevClose));
          if RowCount < 10 then
          RowCount := RowCount + 1;
        end;
        Dataset.Next;
      end;
    end;
  finally
    Dataset.Free;
  end;

end;


procedure TfrmMain.LoadTopGainers;
var
  Dataset: TSQLiteTable;
  CurrentRow: Byte;
  StrValue: String;
  PreviousClose, PercentChg: Double;
begin
  Dataset := PSEDataProvider.GetTopGainers(cbDate.Date, MAX_ROWS);
  try
    grdGainers.RowCount := 0;
    grdGainers.Rows[0].Clear;
    grdGainers.ColWidths[0] := 45;
    grdGainers.ColWidths[1] := 80;
    grdGainers.ColWidths[2] := 80;
    grdGainers.ColWidths[3] := 80;
    grdGainers.ColWidths[4] := 60;
    if Dataset.RowCount > 0 then
    begin
      while not Dataset.EOF do
      begin
        with grdGainers do
        begin
          CurrentRow := RowCount - 1;
          StrValue := Dataset.FieldAsString(Dataset.FieldIndex['VALUE']);
          PreviousClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
          PercentChg := GetPercentChg(StrToFloat(Dataset.FieldByName['CLOSE']), PreviousClose);

          Rows[CurrentRow].Clear;
          Rows[CurrentRow].Append(Dataset.FieldByName['SYMBOL']);
          Rows[CurrentRow].Append(Dataset.FieldByName['VOLUME']);
          Rows[CurrentRow].Append(FloatToStr(StrToFloat(StrValue) / 1000));
          Rows[CurrentRow].Append(Dataset.FieldByName['CLOSE']);
          Rows[CurrentRow].Append(FloatToStr(PercentChg));
          if RowCount < 10 then
          RowCount := RowCount + 1;
        end;
        Dataset.Next;
      end;
    end;
  finally
    Dataset.Free;
  end;

end;


procedure TfrmMain.LoadTopLosers;
var
  Dataset: TSQLiteTable;
  CurrentRow: Byte;
  StrValue: String;
  PreviousClose, PercentChg: Double;
begin
  Dataset := PSEDataProvider.GetTopLosers(cbDate.Date, MAX_ROWS);
  try
    grdLosers.RowCount := 0;
    grdLosers.Rows[0].Clear;
    grdLosers.ColWidths[0] := 45;
    grdLosers.ColWidths[1] := 80;
    grdLosers.ColWidths[2] := 80;
    grdLosers.ColWidths[3] := 80;
    grdLosers.ColWidths[4] := 60;
    if Dataset.RowCount > 0 then
    begin
      while not Dataset.EOF do
      begin
        with grdLosers do
        begin
          CurrentRow := RowCount - 1;
          StrValue := Dataset.FieldAsString(Dataset.FieldIndex['VALUE']);
          PreviousClose := PSEDataProvider.GetPreviousClose(Dataset.FieldByName['SYMBOL'], cbDate.Date);
          PercentChg := GetPercentChg(StrToFloat(Dataset.FieldByName['CLOSE']), PreviousClose);

          Rows[CurrentRow].Clear;
          Rows[CurrentRow].Append(Dataset.FieldByName['SYMBOL']);
          Rows[CurrentRow].Append(Dataset.FieldByName['VOLUME']);
          Rows[CurrentRow].Append(FloatToStr(StrToFloat(StrValue) / 1000));
          Rows[CurrentRow].Append(Dataset.FieldByName['CLOSE']);
          Rows[CurrentRow].Append(FloatToStr(PercentChg));
          if RowCount < 10 then
          RowCount := RowCount + 1;
        end;
        Dataset.Next;
      end;
    end;
  finally
    Dataset.Free;
  end;

end;


procedure TfrmMain.LoadData;
begin
  LoadIndeces;
  LoadTopGainers;
  LoadTopLosers;
  LoadTopForeignBuy;
  LoadTopForeignSell;
  LoadMisc;
end;

procedure TfrmMain.grdMarketDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  Str: String;
  Price: Double;
  StrGrid: TmbTBXStringGrid;
begin
  StrGrid := Sender as TmbTBXStringGrid;
  with StrGrid.Canvas do
  begin
    Str := Trim(StrGrid.Cells[ACol, ARow]);
    if Str <> '' then
    begin
    Price := StrToFloat(StrGrid.Cells[2, ARow]);
    if Price  < 0 then
      Font.Color := clRed
    else if Price = 0 then
      Font.Color := clGray
    else
      Font.Color := clGreen;
    Font.Size := FONT_SIZE;
    Font.Name := 'Verdana';
    Brush.Style := bsClear;

    case ACol of
      0: DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_LEFT OR DT_VCENTER );
      1, 2, 3:
        begin
          Str := FormatFloat('###,##0.00', Abs(StrToFloat(Str)));
          DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_RIGHT OR DT_VCENTER	);
        end;
      4:
        begin
          Price := StrToFloat(Str);
          if Price > 0 then
          begin
            Font.Color := clGreen;
            Str := FormatFloat('###,##0', Price);
          end else if Price < 0 then
          begin
            Font.Color := clRed;
            Str := FormatFloat('###,##0', Abs(Price));
          end else
          begin
            Font.Color := clGray;
          end;
          DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_RIGHT OR DT_VCENTER	);
        end;
      end;
    end;
  end;
end;
procedure TfrmMain.grdGainersDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  Str: String;
  StrGrid: TmbTBXStringGrid;
  Amount: Extended;
begin

  StrGrid := Sender as TmbTBXStringGrid;
  with StrGrid.Canvas do
  begin
    Brush.Style := bsClear;
    Str := Trim(StrGrid.Cells[ACol, ARow]);
    Font.Size := FONT_SIZE;    
    Font.Name := 'Verdana';
    Font.Color := clBlack;
    if Str <> '' then
    begin
      case ACol of
        0: DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_LEFT OR DT_VCENTER );
        1, 2: begin
            if Trunc(StrToFloat(Str)) = 0 then
              Str := FormatFloat('###,##0.00', StrToFloat(Str))
            else
              Str := FormatFloat('###,###', StrToFloat(Str));
            DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_RIGHT OR DT_VCENTER	);
           end;
        3:
          begin
            Amount := StrToFloat(Str);
            if (Amount >= 505) or ((Amount >= 101) and (Amount <= 250)) then
              Str := FormatFloat('###,##0', Amount)
            else if ((Amount >= 252.5) and (Amount <= 500)) or
                    ((Amount >= 0.26) and (Amount <= 100)) then
              Str := FormatFloat('###,##0.00', Amount)
            else
              Str := FormatFloat('###,##0.0000', Amount);
            if (StrGrid.Name = 'grdSell') or (StrGrid.Name = 'grdNetBuy') then
            begin
              if (Amount - StrToFloat(StrGrid.Rows[ARow].Strings[5])) < 0 then
                Font.Color := clRed
              else if (Amount - StrToFloat(StrGrid.Rows[ARow].Strings[5])) > 0 then
                Font.Color := clGreen
              else
                Font.Color := clGray;
            end;
            DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_RIGHT OR DT_VCENTER	);
          end;

        4:
          begin
            Amount := StrToFloat(Str);
            if Amount < 0 then
              Font.Color := clRed
            else
              Font.Color := clGreen;
            if (StrGrid.Name = 'grdSell') or (StrGrid.Name = 'grdNetBuy') then
            begin
              if Trunc(Amount) = 0 then
                Str := FormatFloat('###,##0.00', Abs(Amount))
              else
                Str := FormatFloat('###,##', Abs(Amount));              
            end else
              Str := FormatFloat('###,##0.00', Abs(Amount));

            DrawText(Handle, PChar(Str), Length(Str), Rect, DT_SINGLELINE OR DT_RIGHT OR DT_VCENTER	);
          end;
      end;
    end;
  end;

end;

procedure TfrmMain.suiButton3Click(Sender: TObject);
begin
  LoadData;
end;

procedure TfrmMain.OntopClick(Sender: TObject);
begin
  if OnTop.Checked then
  begin
    SetWindowPos(frmMain.Handle,HWND_NOTOPMOST,
                 Left,Top,Width,Height,
                 SWP_SHOWWINDOW);
  end else
  begin
    SetWindowPos(frmMain.Handle,HWND_TOPMOST,
                 Left,Top,Width,Height,
                 SWP_SHOWWINDOW);
  end;
  frmMain.Update;
  OnTop.Checked := not OnTop.Checked;
end;

procedure TfrmMain.tabMarketShow(Sender: TObject);
begin
  if grdMarket.RowCount = 1 then
  begin
    cbDate.DateTime := PSEDataProvider.GetLastTradeDate;
    LoadData;
  end;
end;

procedure TfrmMain.suiTabSheet1Show(Sender: TObject);
begin
  dtDatePurge.Date := Date - 60;
end;

procedure TfrmMain.btnPurgeClick(Sender: TObject);
var
  SQLiteDatabase: TSQLiteDatabase;
  UDate: Int64;
begin
  SQLiteDatabase := TSQLiteDatabase.Create('pseget.dat');
  try
    UDate := DateTimeToUnix(dtDatePurge.DateTime);
    SQLiteDatabase.ExecSQL('DELETE FROM TRADETAB WHERE DATE < ' + IntToStr(UDate));
    SQLiteDatabase.ExecSQL('vacuum');
  finally
    SQLiteDatabase.Free;
    // reread database
    PSEDataProvider.SQLiteDatabase.Free;
    PSEDataProvider.SQLiteDatabase := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'\pseget.dat');
  end;
  StatusBar.SimpleText := 'Data purging complete.';
end;

end.
