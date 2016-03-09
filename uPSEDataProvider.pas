unit uPSEDataProvider;

interface
uses uPseParser, SQLiteTable3, DateUtils, SysUtils;

type
  TQueryMode = (qryDaily, qryWeekly, qryMonthly);

  TPSEDataProvider = class
    private
      FSQLiteDatabase: TSQLiteDatabase;
    public
      QueryMode: TQueryMode;
      constructor Create;
      function GetTopGainers(const ATradeDate: TDateTime;
                             const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetTopLosers(const ATradeDate: TDateTime;
                            const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetMostActive(const ATradeDate: TDateTime;
                             const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetTopForeignBuy(const ATradeDate: TDateTime;
                                const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetTopForeignSell(const ATradeDate: TDateTime;
                                 const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetIndeces(const ATradeDate: TDateTime): TSQLiteTable;
      function GetMiscData(const ATradeDate: TDateTime;
                           const NumRows: Integer; Offset: Integer = 0): TSQLiteTable;
      function GetData(SQLStr: String): TSQLiteTable;
      function GetPercentChange(ASymbol: String; ATradeDate: TDateTime): Double;
      function GetPreviousClose(ASymbol: String; ATradeDate: TDateTime): Double;
      function GetLastTradeDate: TDateTime;
      function TradeDateExist(TradeDate: TDateTime): Boolean;
      procedure AppendPSETrade(const APSECollection: TPSECollection);
      property SQLiteDatabase: TSQLiteDatabase read FSQLiteDatabase
                                               write FSQLiteDatabase;
  end;

  function GetPercentChg(LastTrade, PreviousTrade: Double): Double;

implementation

function GetPercentChg(LastTrade, PreviousTrade: Double): Double;
begin
  Result := 0;
  if (LastTrade = 0) or (PreviousTrade = 0) then
    Exit;
  Result := (LastTrade - PreviousTrade) / PreviousTrade;
  Result := Result * 100;
end;

{ TPSEDataProvider }

procedure TPSEDataProvider.AppendPSETrade(
  const APSECollection: TPSECollection);
var
  i, j: Integer;
  StockItem: TStockItem;
  PhisixValue: Double;
  Sector: TSectorCollection;
  SQLStr: String;
  Tradekey: integer;
  PreviousClose: Double;
  PercentChg: Double;
  OpenInterest: Integer;
const
  InsertStr = 'INSERT INTO TRADETAB VALUES(null, %d, ''%s'', %12.5f, %12.5f, %12.5f, %12.5f, %d, %12.5f, %d, %s)';
begin
  FSQLiteDatabase.BeginTransaction;
  try
    if TradeDateExist(APSECollection.TradeDate) then
    begin
      FSQLiteDatabase.ExecSQL('DELETE FROM TRADETAB WHERE DATE = ' +
                              IntToStr(DateTimeToUnix(APSECollection.TradeDate)));
      {FSQLiteDatabase.ExecSQL('DELETE FROM MISCTAB WHERE DATE = ' +
                             IntToStr(DateTimeToUnix(APSECollection.TradeDate)));}
    end;

    with APSECollection do
    begin
      if TradeDate < StrToDate('01/02/2006') then
        PhisixValue := Items[7].SectorCollection.Value
      else
        PhisixValue := Items[10].SectorCollection.Value;
      for i := 0 to Count - 1 do
      begin
        Sector := Items[i].SectorCollection;
        //PreviousClose := GetPreviousClose(Sector.Code, TradeDate);
        //PercentChg := GetPercentChg(Sector.Close, PreviousClose);
        if Sector.Code = '^PHISIX' then
          OpenInterest := Trunc(TotalForeignBuying - TotalForeignSelling) div 1000
        else
          OpenInterest := Trunc(Sector.ForeignBuy - Sector.ForeignSell) div 1000;
        SQLStr := Format(InsertStr, [
                                DateTimeToUnix(TradeDate),
                                Sector.Code,
                                Sector.Open,
                                Sector.High,
                                Sector.Low,
                                Sector.Close,
                                Sector.Volume,
                                Sector.Value,
                                OpenInterest,
                                FormatFloat('#0.00', (Sector.Value / PhisixValue) * 100){,
                                Abs(Sector.Close - PreviousClose),
                                FormatFloat('#0.00', PercentChg)}
                                ]);
        FSQLiteDatabase.ExecSQL(SQLStr);
        for j := 0 to Sector.Count - 1 do
        begin
          StockItem := Sector.Items[j];
          if StockItem.Close = 0 then
            Continue;
          //PreviousClose := GetPreviousClose(StockItem.Symbol, TradeDate);
          //PercentChg := GetPercentChg(StockItem.Close, PreviousClose);
          SQLStr := Format(InsertStr, [
                                DateTimeToUnix(TradeDate),
                                StockItem.Symbol,
                                StockItem.Open,
                                StockItem.High,
                                StockItem.Low,
                                StockItem.Close,
                                StockItem.Volume,
                                StockItem.Value,
                                StockItem.NetForeignBuy,
                                FormatFloat('#0.00', (StockItem.Value / PhisixValue) * 100){,
                                Abs(StockItem.Close - PreviousClose),
                                FormatFloat('#0.00', PercentChg)}
                                ]);
          FSQLiteDatabase.ExecSQL(SQLStr);
        end;
      end;

      SQLStr := 'INSERT INTO MISCTAB VALUES(null, %d, %d, %d, %d, %d, %d, %12.5f, %12.5f, %d, %d)';
      SQLStr := Format(SQLStr, [DateTimeToUnix(TradeDate),
                                NumAdvances,
                                NumDeclines,
                                NumUnchanged,
                                NumTraded,
                                NumTrades,
                                TotalForeignBuying,
                                TotalForeignSelling,
                                MainCrossVolume,
                                MainCrossValue]);
      FSQLiteDatabase.ExecSQL(SQLStr);
    end;

    FSQLiteDatabase.Commit;
  except
    FSQLiteDatabase.Rollback;
    raise;
  end;
end;

constructor TPSEDataProvider.Create;
begin
  QueryMode := qryDaily;
end;

function TPSEDataProvider.GetData(SQLStr: String): TSQLiteTable;
begin
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetIndeces(const ATradeDate: TDateTime): TSQLiteTable;
var
  SQLStr: String;
begin
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT * FROM TRADETAB WHERE SYMBOL LIKE ''^%'' AND DATE = ' +
            IntToStr(DateTimeToUnix(ATradeDate)) + 'AND SYMBOL ';
      if ATradeDate < StrToDateTime('01/02/2006') then
        SQLStr := SQLStr + ' <> ''^PHISIX'' '
      else
        SQLStr := SQLStr + ' NOT IN( ''^PHISIX'', ''^PREFERRED'', ''^WARRANT'', ''^SME'') ';

      SQLStr := SQLStr + ' ORDER BY SYMBOL';
  end;
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetLastTradeDate: TDateTime;
var
  SQL: String;
begin
  SQL := 'SELECT MAX(DATE) FROM TRADETAB';
  Result := UnixToDateTime(FSQLiteDatabase.GetTableValue(SQL));
end;

function TPSEDataProvider.GetMiscData(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
var
  SQLStr: String;
  UDate: Int64;
begin
  UDate := DateTimeToUnix(ATradeDate);
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT MISCTAB.*, (SELECT VALUE FROM TRADETAB WHERE SYMBOL = ''^PHISIX'' AND DATE = %d) AS VALUE FROM MISCTAB WHERE DATE = %d';
  end;
  SQLStr := Format(SQLStr, [UDate, UDate]);
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetMostActive(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
begin
end;

function TPSEDataProvider.GetPercentChange(ASymbol: String;
  ATradeDate: TDateTime): Double;
var
  SQL: String;
  PreviousClose: Double;
  LastTrade: Double;
  Dataset: TSQLiteTable;
begin
  Result := 0;
  LastTrade := 0;
  PreviousClose := 0;
  SQL := Format('SELECT CLOSE FROM TRADETAB WHERE SYMBOL = ''%s'' AND ' +
                ' DATE < %d ORDER BY DATE DESC LIMIT 1',
                [ASymbol, DateTimeToUnix(ATradeDate)]);
  Dataset := FSQLiteDatabase.GetTable(SQL);
  if Dataset.RowCount > 0 then
    PreviousClose := Dataset.FieldAsDouble(0);
  Dataset.Free;

  SQL := Format('SELECT CLOSE FROM TRADETAB WHERE SYMBOL = ''%s'' AND ' +
                ' DATE = %d', [ASymbol, DateTimeToUnix(ATradeDate)]);

  Dataset := FSQLiteDatabase.GetTable(SQL);
  if Dataset.RowCount > 0 then
    LastTrade := Dataset.FieldAsDouble(0);
  Dataset.Free;

  if (PreviousClose = 0) or (LastTrade = 0) then
    Exit;
  Result := (LastTrade - PreviousClose) / PreviousClose;
  Result := Result * 100;
end;

function TPSEDataProvider.GetPreviousClose(ASymbol: String;
  ATradeDate: TDateTime): Double;
var
  SQL: String;
  Dataset: TSQLiteTable;
begin
  Result := 0;
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQL := 'SELECT CLOSE FROM TRADETAB WHERE SYMBOL = ''%s'' AND DATE < %d ORDER BY DATE DESC LIMIT 1';
  end;
  SQL := Format(SQL, [ASymbol, DateTimeToUnix(ATradeDate)]);
  Dataset := FSQLiteDatabase.GetTable(SQL);
  try
    if Dataset.RowCount > 0 then
      Result := StrToFloat(Dataset.Fields[0]);
  finally
    Dataset.Free;
  end;
end;

function TPSEDataProvider.GetTopForeignBuy(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
var
  SQLStr: String;
begin
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT * FROM TRADETAB WHERE DATE = %d AND OPENINTEREST > 0' +
                ' AND SYMBOL NOT LIKE ''^%s'''  +
                ' ORDER BY OPENINTEREST DESC LIMIT %d OFFSET %d';
  end;
  SQLStr := Format(SQLStr, [DateTimeToUnix(ATradeDate), '%', NumRows, Offset]);
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetTopForeignSell(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
var
  SQLStr: String;
begin
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT * FROM TRADETAB WHERE DATE = %d AND OPENINTEREST < 0' +
                ' AND SYMBOL NOT LIKE ''^%s''' +
                ' ORDER BY OPENINTEREST LIMIT %d OFFSET %d';
  end;
  SQLStr := Format(SQLStr, [DateTimeToUnix(ATradeDate), '%', NumRows, Offset]);
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetTopGainers(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
var
  SQLStr: String;
  UDate: Int64;
begin
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT C.*, (' +
                ' ((C.CLOSE - (SELECT A.CLOSE FROM TRADETAB A WHERE A.SYMBOL = C.SYMBOL AND A.DATE < %d ORDER BY A.DATE DESC LIMIT 1) ) /' +
                '   (SELECT B.CLOSE FROM TRADETAB B WHERE B.SYMBOL = C.SYMBOL AND B.DATE < %d ORDER BY B.DATE DESC LIMIT 1)) * 100' +
                ' ) AS X' +
                ' FROM TRADETAB C WHERE C.DATE = %d AND C.SYMBOL NOT LIKE ''^%s''' +
                ' ORDER BY X DESC LIMIT %d OFFSET %d';

  end;
  UDate := DateTimeToUnix(ATradeDate);
  SQLStr := Format(SQLStr, [UDate, UDate, UDate, '%', NumRows, Offset]);
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.GetTopLosers(const ATradeDate: TDateTime;
  const NumRows: Integer; Offset: Integer): TSQLiteTable;
var
  SQLStr: String;
  UDate: Int64;
begin
  case QueryMode of
    qryWeekly: ;
    qryMonthly: ;
    else
      SQLStr := 'SELECT C.*, (' +
            ' ((C.CLOSE - (SELECT A.CLOSE FROM TRADETAB A WHERE A.SYMBOL = C.SYMBOL AND A.DATE < %d ORDER BY A.DATE DESC LIMIT 1) ) /' +
            '   (SELECT B.CLOSE FROM TRADETAB B WHERE B.SYMBOL = C.SYMBOL AND B.DATE < %d ORDER BY B.DATE DESC LIMIT 1)) * 100' +
            ' ) AS X' +
            ' FROM TRADETAB C WHERE C.DATE = %d AND C.SYMBOL NOT LIKE ''^%s''' +
            ' AND (SELECT A.CLOSE FROM TRADETAB A WHERE A.SYMBOL = C.SYMBOL AND A.DATE < %d ORDER BY A.DATE DESC LIMIT 1) <> 0'+
            ' ORDER BY X ASC LIMIT %d OFFSET %d';
  end;
  UDate := DateTimeToUnix(ATradeDate);
  SQLStr := Format(SQLStr, [UDate, UDate, UDate, '%', UDate, NumRows, Offset]);
  Result := FSQLiteDatabase.GetTable(SQLStr);
end;

function TPSEDataProvider.TradeDateExist(TradeDate: TDateTime): Boolean;
var
  SQLStr: String;
begin
  SQLStr := 'SELECT COUNT(*) AS CTR FROM MISCTAB WHERE DATE = ' + IntToStr(DateTimeToUnix(TradeDate));
  Result := SQLiteDatabase.GetTableValue(SQLStr) > 0;
end;

end.
