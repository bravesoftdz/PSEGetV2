unit uExNotice;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, JvComponent, JvMagnet, ExtCtrls, SUIForm, uPseParser,
  ComCtrls, AdvPageControl, Grids, mbTBXDrawGrid,
  mbTBXStringGrid, SUIImagePanel, StdCtrls, mbTBXPanel,
  JvDateTimePicker, uPSEDataProvider, mbTBXHint, SUIButton,
  SUIPopupMenu, JvLabel, Menus, SUIMemo, SUIScrollBar;

type
  TfrmNotice = class(TForm)
    JvFormMagnet1: TJvFormMagnet;
    suiForm1: TsuiForm;
    mmNotice: TsuiMemo;
    suiScrollBar1: TsuiScrollBar;
    procedure FormShow(Sender: TObject);
    procedure grdMarketDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure grdGainersDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  frmNotice: TfrmNotice;

implementation

uses uMain;

{$R *.dfm}

const
  MAX_ROWS = 10;

procedure TfrmNotice.FormShow(Sender: TObject);
begin
  Height := frmMain.Height;
  Left := frmMain.Left + frmMain.Width;
  Top := frmMain.Top;
end;


procedure TfrmNotice.grdMarketDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
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

procedure TfrmNotice.grdGainersDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
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

procedure TfrmNotice.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  //Params.ExStyle := Params.ExStyle or WS_EX_APPWINDOW; //or WS_EX_TOPMOST;
  //Params.WndParent := GetDesktopwindow;
end;

end.
