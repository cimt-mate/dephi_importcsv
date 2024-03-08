unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Uni, UniProvider, OracleUniProvider, MemDS, Vcl.Grids, Vcl.DBGrids,
  DBAccess, Vcl.ExtDlgs,System.IniFiles ,DateUtils, System.ImageList,
  Vcl.ImgList, Vcl.Buttons, Vcl.Menus, Winapi.Winsock   , System.Types  ,System.IOUtils,System.StrUtils

;

type
  TForm1 = class(TForm)
    FolderDialog: TFileOpenDialog;
    EditFolderPath: TEdit;
    OpenFolderPath: TButton;
    ButtonReadCSV: TButton;
    StringGridCSV: TStringGrid;
    SpeedButtonIMP: TSpeedButton;
    UniConnection: TUniConnection;
    OracleUniProvider: TOracleUniProvider;
    UniQuery: TUniQuery;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    procedure OpenFolderPathClick(Sender: TObject);
    procedure ButtonReadClick(Sender: TObject);
    procedure SpeedButtonIMPClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    procedure LoadCSVFilesIntoGrid(const FolderPath: string);
    procedure ImportDataToDatabase;
    procedure SetupDatabaseQuery;
    procedure LoadConnectionParameters;
    procedure WriteLog(const LogMessage: string);
    procedure InsertLogData(LogDateTime: TDateTime; LogKBN: Integer;
  Message, PGName, PGVersion, MessageID: string; Confirm: Integer; ConfirmDate: TDateTime);
    function GetProgramName: string;
    function GetAppVersion: string;
    procedure ReadSettings;
    procedure CreateStringGrid(var Grid: TStringGrid; AParent: TWinControl);
    procedure StringGridCSVDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    function GetColumnIndexByHeaderName(StringGrid: TStringGrid; HeaderName: string): Integer;
    procedure UpdateResultColumn(Row: Integer; const ResultText: string);
    function GetBunmFromBucd(const BucdValue: string): string;
    function FormatDateTimeStr(const DateStr, TimeStr: string): string;
    function GetMaxTime(const Time1, Time2: string): string;
    function GetTimeInMinutes(const TimeStr: string): Integer;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}



procedure TForm1.ReadSettings;
var
  IniFile: TIniFile;
  IniFileName: string;
  Choice: string;
begin
 // Read ini file
  IniFileName := ExtractFilePath(Application.ExeName) + '/GRD/DS11EPN100.ini';
  IniFile := TIniFile.Create(IniFileName);
  try
    EditFolderPath.Text := IniFile.ReadString('Settings', 'Path', '');
    // Initialize other settings as needed.
  finally
    IniFile.Free;
  end;
end;


procedure TForm1.FormCreate(Sender: TObject);
var
  CustomBlue: TColor;
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
  V1, V2, V3, V4: Word;
  VersionString, FileNameString, LocalHostString, TimeString: string;
  WSAData: TWSAData;
  IniFile: TIniFile;
  FileName: string;
  // Can't declare Username, Password *Conflict with UnitConnection Variable Name
  DirectDBName, User, Pass: string;
begin
  CustomBlue := rgb(194,209,254); // Standard blue color
  Self.Color := CustomBlue;
  CreateStringGrid(StringGridCSV, Self);

  StatusBar1.Panels.Clear;
  ReadSettings;

  // Initialize Winsock
  WSAStartup(MAKEWORD(2, 2), WSAData);
  try
    // Add panel for the file name
    FileNameString := '' + ExtractFileName(Application.ExeName);
    with StatusBar1.Panels.Add do
    begin
      Width := 95;
      Text := FileNameString;
    end;

    // Add panel for version info
    VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
    if VerInfoSize > 0 then
    begin
      GetMem(VerInfo, VerInfoSize);
      try
        if GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo) then
        begin
          if VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize) then
          begin
            V1 := HiWord(VerValue^.dwFileVersionMS);
            V2 := LoWord(VerValue^.dwFileVersionMS);
            V3 := HiWord(VerValue^.dwFileVersionLS);
            V4 := LoWord(VerValue^.dwFileVersionLS);
            VersionString := Format('%d.%d.%d.%d', [V1, V2, V3, V4]);
          end;
        end;
      finally
        FreeMem(VerInfo, VerInfoSize);
      end;
    end
    else
    begin
      VersionString := 'Version not found';
    end;

    with StatusBar1.Panels.Add do
    begin
      Width := 70;
      Text := VersionString;
    end;

    // Add panel for localhost
      FileName := ExtractFilePath(Application.ExeName) + '/Setup/SetUp.Ini';// Assumes the INI file is in the same directory as the application
      IniFile := TIniFile.Create(FileName);
      DirectDBName := IniFile.ReadString('Setting', 'DIRECTDBNAME', '');
      User := IniFile.ReadString('Setting', 'USERNAME', '');
    with StatusBar1.Panels.Add do
    begin
      Width := 300;
      Text := DirectDBName +':' + User;
    end;

    // Add panel for current time
        with StatusBar1.Panels.Add do
    begin
      Width := 270;
    end;
      Timer1.Interval := 1000; // Trigger every 1000 milliseconds (1 second)
      Timer1.Enabled := True;
      Timer1.OnTimer := Timer1Timer;
      finally
    WSACleanup;
  end;
end;




function TForm1.GetAppVersion: string;
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then
    RaiseLastOSError;

  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then
    RaiseLastOSError;

  if VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    Result := Format('%d.%d.%d.%d',
      [HiWord(FixedPtr^.dwFileVersionMS), LoWord(FixedPtr^.dwFileVersionMS),
       HiWord(FixedPtr^.dwFileVersionLS), LoWord(FixedPtr^.dwFileVersionLS)])
  else
    Result := '';
end;

function TForm1.GetProgramName: string;
begin
  Result := ExtractFileName(Application.ExeName);
end;

procedure TForm1.SetupDatabaseQuery;
begin
  UniConnection := TUniConnection.Create(nil);
  LoadConnectionParameters;
  UniQuery := TUniQuery.Create(nil);
  UniQuery.Connection := UniConnection;
  UniQuery.Open;
end;

procedure TForm1.LoadConnectionParameters;
var
  IniFile: TIniFile;
  FileName: string;
  // Can't declare Username, Password *Conflict with UnitConnection Variable Name
  DirectDBName, User, Pass: string;
begin
  FileName := ExtractFilePath(Application.ExeName) + 'Setup\SetUp.Ini'; // Use backslash for path in Windows
  // Check if the INI file exists
  if not FileExists(FileName) then
  begin
    WriteLog('Error: INI file not found at ' + FileName); // Replace WriteLog with your actual logging procedure
    InsertLogData(Now,1,'Error: INI file not found at ' + FileName,GetProgramName,GetAppVersion, '', 0, Now) ;
    Exit; // Exit the procedure if the file does not exist
  end;

  IniFile := TIniFile.Create(FileName);
  try
    DirectDBName := IniFile.ReadString('Setting', 'DIRECTDBNAME', '');
    User := IniFile.ReadString('Setting', 'USERNAME', '');
    Pass := IniFile.ReadString('Setting', 'PASSWORD', '');
    with UniConnection do
    begin
      if not Connected then
      begin
        ProviderName := 'Oracle';
        SpecificOptions.Values['Direct'] := 'True';
        Server := DirectDBName;
        Username := User;
        Password := Pass;
        Connect; // Establish the connection
      end;
    end;
  finally
    IniFile.Free; // Always free the TIniFile object when done
  end;
end;

procedure TForm1.WriteLog(const LogMessage: string);
var
  LogFileName: string;
  LogFile: TextFile;
  LineCount: Integer;
  TempList: TStringList;
begin
  LogFileName := ExtractFilePath(Application.ExeName) + '/DS11EPN100_log.log';

  // Counting the number of lines in the log file
  LineCount := 0;
  if FileExists(LogFileName) then
  begin
    TempList := TStringList.Create;
    try
      TempList.LoadFromFile(LogFileName);
      LineCount := TempList.Count;
    finally
      TempList.Free;
    end;
  end;

  // Clear the log file if it exceeds 200 lines
  if LineCount >= 200 then
  begin
    AssignFile(LogFile, LogFileName);
    Rewrite(LogFile); // This clears the file
    CloseFile(LogFile);
  end;
  // Append the new log message
  try
    AssignFile(LogFile, LogFileName);
    if FileExists(LogFileName) then
      Append(LogFile)
    else
      Rewrite(LogFile);

    Writeln(LogFile, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ': ' + LogMessage);
  finally
    CloseFile(LogFile);
  end;
end;


procedure TForm1.InsertLogData(LogDateTime: TDateTime; LogKBN: Integer;
  Message, PGName, PGVersion, MessageID: string; Confirm: Integer; ConfirmDate: TDateTime);
var
  UniQuery: TUniQuery;
begin
  UniQuery := TUniQuery.Create(nil);
  try
    UniQuery.Connection := UniConnection; // Assuming UniConnection is already set up

    UniQuery.SQL.Text := 'INSERT INTO LOGDATA (LOGYMD, LOGKBN, MESSAGE, ' +
                         'PGNAME, PGVERSION, MESSAGEID, CONFIRM, CONFIRMDATE) ' +
                         'VALUES (:LOGYMD, :LOGKBN, :MESSAGE, :PGNAME, ' +
                         ':PGVERSION, :MESSAGEID, :CONFIRM, :CONFIRMDATE)';

    UniQuery.Params.ParamByName('LOGYMD').AsDateTime := LogDateTime;
    UniQuery.Params.ParamByName('LOGKBN').AsInteger := LogKBN;
    UniQuery.Params.ParamByName('MESSAGE').AsString := Message;
    UniQuery.Params.ParamByName('PGNAME').AsString := PGName;
    UniQuery.Params.ParamByName('PGVERSION').AsString := PGVersion;
    UniQuery.Params.ParamByName('MESSAGEID').AsString := MessageID;
    UniQuery.Params.ParamByName('CONFIRM').AsInteger := Confirm;
    UniQuery.Params.ParamByName('CONFIRMDATE').AsDateTime := ConfirmDate;

    UniQuery.Prepare;
    UniQuery.Execute;
  finally
    UniQuery.Free;
  end;
end;


procedure TForm1.OpenFolderPathClick(Sender: TObject);
begin
  FolderDialog.Options := FolderDialog.Options + [fdoPickFolders];
  if FolderDialog.Execute then
  begin
    EditFolderPath.Text := FolderDialog.FileName;
  end;
end;

procedure TForm1.SpeedButtonIMPClick(Sender: TObject);
begin
   ImportDataToDatabase;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  // Update only the panel for current time
  StatusBar1.Panels[StatusBar1.Panels.Count - 1].Text := DateTimeToStr(Now);
end;


procedure TForm1.ButtonReadClick(Sender: TObject);
begin
  LoadCSVFilesIntoGrid(EditFolderPath.Text);
end;

procedure TForm1.ImportDataToDatabase;
var
  Row, Col: Integer;
  Value1, Value2, JhValue : Integer;
  InsertQuery: TUniQuery;
  SQL, SeizonoValue, BucdValue, BunmValue: string;
  Gkoteicd, Kikaicd, Jigucd, Tantocd, Ymds: string;
  MaxTime, FormattedDateTime: string;
  MaxJDSEQNO, NewJDSEQNO,GHIMOKUCDValue: Integer;
  YujintankaValue,KikaitankaValue,KoteitankaValue,YujinkinValue,MujinkinValue,KinsumValue: Double;
begin
  // Load the database connection parameters
  LoadConnectionParameters;

  // Check if the UniConnection is connected
  if UniConnection.Connected then
  begin
    // Initialize the query component
    InsertQuery := TUniQuery.Create(nil);
    try
      InsertQuery.Connection := UniConnection;

      // Start a transaction
      UniConnection.StartTransaction;

      for Row := 1 to StringGridCSV.RowCount - 1 do
      begin
        try
          // Prepare data for insertion
          SeizonoValue := StringGridCSV.Cells[9, Row];
          BucdValue := StringGridCSV.Cells[10, Row];
          BunmValue := GetBunmFromBucd(BucdValue);
          Gkoteicd := StringGridCSV.Cells[4, Row];
          Kikaicd := StringGridCSV.Cells[14, Row];
          Jigucd := StringGridCSV.Cells[21, Row];
          Tantocd := StringGridCSV.Cells[3, Row];
          Ymds := StringGridCSV.Cells[1, Row];
          MaxTime := GetMaxTime(StringGridCSV.Cells[11, Row], StringGridCSV.Cells[17, Row]);
          FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime);
          if not TryStrToInt(StringGridCSV.Cells[13, Row], Value1) then Value1 := 0;
          if not TryStrToInt(StringGridCSV.Cells[20, Row], Value2) then Value2 := 0;
          JhValue := Value1 + Value2;

          // Get yujintanka value from tantomst
          InsertQuery.SQL.Text := 'SELECT tanka1 FROM tantomst WHERE tantocd = :tantocd';
          InsertQuery.ParamByName('tantocd').AsString := Tantocd;
          InsertQuery.Open;
          YujintankaValue := InsertQuery.FieldByName('tanka1').AsFloat; // Assuming tanka1 is a float
          InsertQuery.Close;

          // Get Kikaitanka value from the kikaimst table using kikaicd
          if Kikaicd.IsEmpty then
            KikaitankaValue := 0.0
          else
          begin
          InsertQuery.SQL.Text := 'SELECT KIKAITANKA FROM kikaimst WHERE kikaicd = :kikaicd';
          InsertQuery.ParamByName('kikaicd').AsString := Kikaicd;
          InsertQuery.Open;
          if not InsertQuery.FieldByName('KIKAITANKA').IsNull then
            KikaitankaValue := InsertQuery.FieldByName('KIKAITANKA').AsFloat
          else
            KikaitankaValue := 0.0; // Set default value if null
          InsertQuery.Close;
          end;

          // Get koteitanka value from the kouteigmst table using Gkoteicd
          InsertQuery.SQL.Text := 'SELECT GTANKA FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          if not InsertQuery.FieldByName('GTANKA').IsNull then
            KoteitankaValue := InsertQuery.FieldByName('GTANKA').AsFloat
          else
            KoteitankaValue := 0.0; // Set default value if null
          InsertQuery.Close;

          // Retrieve GHIMOKUCD from KOUTEIGMST
          InsertQuery.SQL.Text := 'SELECT GHIMOKUCD FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          if not InsertQuery.FieldByName('GHIMOKUCD').IsNull then
            GHIMOKUCDValue := InsertQuery.FieldByName('GHIMOKUCD').AsInteger
          else
            GHIMOKUCDValue := 0; // Default value for integer
          InsertQuery.Close;

          // Get the maximum JDSEQNO from the JISEKIDATA table
          InsertQuery.SQL.Text := 'SELECT MAX(JDSEQNO) AS MaxJDSEQNO FROM JISEKIDATA';
          InsertQuery.Open;
          MaxJDSEQNO := InsertQuery.FieldByName('MaxJDSEQNO').AsInteger;
          InsertQuery.Close;
          // Increment the maximum JDSEQNO by 1 to get the new JDSEQNO
          NewJDSEQNO := MaxJDSEQNO + 1;

          // YujinkinValue from VALUE1 x YujintankaValue / 60;
          YujinkinValue := Value1 * YujintankaValue / 60;
          // MujinkinValue from  Value2 x KikaitankaValue / 60;
          MujinkinValue := Value2 * KikaitankaValue / 60;
          //kinsum
          KinsumValue :=  YujinkinValue +   MujinkinValue;


          // Construct and execute the SQL statement
          SQL := 'INSERT INTO JISEKIDATA (JDSEQNO, seizono, bunm, bucd, gkoteicd, kikaicd, jigucd, tantocd, ymds, KMSEQNO , '+
          'jh ,jmaedanh,jatodanh,jkbn,jyujinh,jmujinh,yujintanka,kikaitanka,koteitanka,GHIMOKUCD,yujinkin,mujinkin,kinsum) ' +
          'VALUES (:NewJDSEQNO, :SeizonoValue, :BunmValue, :BucdValue, :Gkoteicd, :Kikaicd, :Jigucd, :Tantocd, TO_DATE(:FormattedDateTime, ''YYYY-MM-DD HH24:MI:SS''),  '+
          '1,:JhValue,0,0,4,:Value1,:Value2,:YujintankaValue,:KikaitankaValue,:KoteitankaValue,:GHIMOKUCDValue,:YujinkinValue,:MujinkinValue,:KinsumValue)';

          InsertQuery.SQL.Text := SQL;
          InsertQuery.ParamByName('NewJDSEQNO').AsInteger := NewJDSEQNO;
          InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
          InsertQuery.ParamByName('BunmValue').AsString := BunmValue;
          InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.ParamByName('Kikaicd').AsString := Kikaicd;
          InsertQuery.ParamByName('Jigucd').AsString := Jigucd;
          InsertQuery.ParamByName('Tantocd').AsString := Tantocd;
          InsertQuery.ParamByName('FormattedDateTime').AsString := FormattedDateTime;
          InsertQuery.ParamByName('JhValue').AsInteger := JhValue;  // Set JhValue as an integer parameter
          InsertQuery.ParamByName('Value1').AsInteger := Value1; //jyujinh
          InsertQuery.ParamByName('Value2').AsInteger := Value2;   //jmujinh
          InsertQuery.ParamByName('YujintankaValue').AsFloat := YujintankaValue; // Set YujintankaValue as a float parameter
          InsertQuery.ParamByName('KikaitankaValue').AsFloat := KikaitankaValue;
          InsertQuery.ParamByName('KoteitankaValue').AsFloat := KoteitankaValue;
          InsertQuery.ParamByName('GHIMOKUCDValue').AsInteger := GHIMOKUCDValue;
          InsertQuery.ParamByName('YujinkinValue').AsFloat := YujinkinValue;
          InsertQuery.ParamByName('MujinkinValue').AsFloat := MujinkinValue;
          InsertQuery.ParamByName('KinsumValue').AsFloat := KinsumValue;

          InsertQuery.Execute;

          // Update the "Result" column with "Imported"
          UpdateResultColumn(Row, 'Imported');
        except
          on E: Exception do
          begin
            // Log the error and update the "Result" column with "NG"
            WriteLog('Error importing row ' + IntToStr(Row) + ': ' + E.Message);
            UpdateResultColumn(Row, 'NG');
          end;
        end;
      end;

      // Commit the transaction
      UniConnection.Commit;
      ShowMessage('Data import process completed');
    finally
      InsertQuery.Free;
    end;
  end
  else
  begin
    ShowMessage('Error: Database connection is not established.');
  end;
end;

function TForm1.GetTimeInMinutes(const TimeStr: string): Integer;
var
  Hour, Min: Integer;
begin
  Result := 0;
  if (TimeStr <> '') and (Pos('.', TimeStr) > 0) then
  begin
    Hour := StrToIntDef(Copy(TimeStr, 1, Pos('.', TimeStr) - 1), 0);
    Min := StrToIntDef(Copy(TimeStr, Pos('.', TimeStr) + 1, Length(TimeStr)), 0);
    Result := Hour * 60 + Min;
  end;
end;

function TForm1.GetMaxTime(const Time1, Time2: string): string;
var
  Time1Minutes, Time2Minutes: Integer;
  ValidTime1, ValidTime2: Boolean;
begin
  // Initialize the result
  Result := '';

  // Check if the time strings are not empty and contain a period
  ValidTime1 := (Time1 <> '') and (Pos('.', Time1) > 0);
  ValidTime2 := (Time2 <> '') and (Pos('.', Time2) > 0);

  // Convert the time strings from 'HH.MM' format to total minutes if valid
  if ValidTime1 then
    Time1Minutes := StrToInt(Copy(Time1, 1, Pos('.', Time1) - 1)) * 60 + StrToInt(Copy(Time1, Pos('.', Time1) + 1, Length(Time1)));
  if ValidTime2 then
    Time2Minutes := StrToInt(Copy(Time2, 1, Pos('.', Time2) - 1)) * 60 + StrToInt(Copy(Time2, Pos('.', Time2) + 1, Length(Time2)));

  // Compare the total minutes and return the maximum time in 'HH:NN:SS' format
  if ValidTime1 and ValidTime2 then
  begin
    if Time1Minutes > Time2Minutes then
      Result := Format('%2d:%2d:00', [Time1Minutes div 60, Time1Minutes mod 60])  // Convert minutes back to 'HH:NN:SS' format
    else
      Result := Format('%2d:%2d:00', [Time2Minutes div 60, Time2Minutes mod 60]); // Convert minutes back to 'HH:NN:SS' format
  end
  else if ValidTime1 then
    Result := Format('%2d:%2d:00', [Time1Minutes div 60, Time1Minutes mod 60])
  else if ValidTime2 then
    Result := Format('%2d:%2d:00', [Time2Minutes div 60, Time2Minutes mod 60]);
end;




function TForm1.FormatDateTimeStr(const DateStr, TimeStr: string): string;
var
  DateTime: TDateTime;
begin
  // First, convert the date string from 'dd/mm/yyyy' to 'yyyy-mm-dd'
  if TryStrToDate(DateStr, DateTime) then
    Result := FormatDateTime('yyyy-mm-dd', DateTime)
  else
    Result := '';

  // Then append the time in 'hh:nn:ss' format
  Result := Result + ' ' + TimeStr ;  // Adding seconds as '00'
end;


function TForm1.GetBunmFromBucd(const BucdValue: string): string;
var
  Query: TUniQuery;
begin
  Result := ''; // Default result is empty string, indicating not found or an error
  Query := TUniQuery.Create(nil);
  try
    Query.Connection := UniConnection; // Use your existing database connection
    Query.SQL.Text := 'SELECT bunm FROM buhinmst WHERE bucd = :Bucd';
    Query.ParamByName('Bucd').AsString := BucdValue;
    Query.Open;
    if not Query.IsEmpty then
      Result := Query.Fields[0].AsString; // Assuming 'bunm' is the first field
  finally
    Query.Free;
  end;
end;



procedure TForm1.UpdateResultColumn(Row: Integer; const ResultText: string);
begin
  StringGridCSV.Cells[0, Row] := ResultText;
end;


procedure TForm1.LoadCSVFilesIntoGrid(const FolderPath: string);
var
  Files: TStringDynArray;
  CSVLines: TStringList;
  FilePath, Line: string;
  Row, Col, MaxCol: Integer;
  CSVHeaderRead: Boolean;
begin
  Files := TDirectory.GetFiles(FolderPath, '*.csv');
  CSVHeaderRead := False;
  MaxCol := 0;

  for FilePath in Files do
  begin
    CSVLines := TStringList.Create;
    try
      CSVLines.LoadFromFile(FilePath);

      // Determine column count from header if it's not been set yet
      if not CSVHeaderRead then
      begin
        var HeaderCells := CSVLines[0].Split([',']);
        MaxCol := Length(HeaderCells) + 1; // Add 1 to the column count for the Result column
        StringGridCSV.ColCount := MaxCol;
        StringGridCSV.RowCount := 1; // Start from row 1 to skip the title row
        CSVHeaderRead := True;
      end;

      // Skip the header line in the CSV file
      for Row := 2 to CSVLines.Count do
      begin
        Line := CSVLines[Row - 1];
        var Cells := Line.Split([',']);
        // Add a new row to the grid for each line read
        StringGridCSV.RowCount := StringGridCSV.RowCount + 1;
        // Fill the row with data, starting from column 1
        for Col := 1 to High(Cells) + 1 do
        begin
          if Col < MaxCol then
            StringGridCSV.Cells[Col, StringGridCSV.RowCount - 1] := Cells[Col - 1]; // Start filling from column 1
        end;
      end;
    finally
      CSVLines.Free;
    end;
  end;
end;



function TForm1.GetColumnIndexByHeaderName(StringGrid: TStringGrid; HeaderName: string): Integer;
var
  Col: Integer;
begin
  Result := -1; // Default result if header not found
  for Col := 0 to StringGrid.ColCount - 1 do
  begin
    if StringGrid.Cells[Col, 0] = HeaderName then // Assuming row 0 has the headers
    begin
      Result := Col;
      Break;
    end;
  end;
end;

procedure TForm1.StringGridCSVDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    var
      S: string;
      InfoColumnIndex: Integer;
      Grid: TStringGrid;
    const
      CellPadding = 2; // Adjust the padding as needed
    begin
      Grid := Sender as TStringGrid;
      S := Grid.Cells[ACol, ARow]; // Get the cell text
      InfoColumnIndex := GetColumnIndexByHeaderName(StringGridCSV, 'Info');
      // Check if this is a header cell
      if ARow = 0 then
      begin
        if (ACol = 0) then
          Grid.Canvas.Brush.Color := clWebLightYellow
        else
          Grid.Canvas.Brush.Color := clWebLightBlue; // Use the specific color you want for the header
        // Header cell formatting

        Grid.Canvas.FillRect(Rect);
        Grid.Canvas.Font.Color := clWindowText;
        DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end
      else if ACol = 0 then // Change for your "Result" column index
      begin
        // "NG" cell formatting
        if AnsiStartsText('NG', S) then
        begin
          Grid.Canvas.Brush.Color := clYellow; // Entire cell background color for "NG" cells
          Grid.Canvas.Font.Color := clRed; // Text color for "NG" cells
          Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
        end
        else if S = 'Imported' then
        begin
          Grid.Canvas.Brush.Color := clWebLightGreen; // Entire cell background color for "NG" cells
          Grid.Canvas.Font.Color := clBlack; // Text color for "NG" cells
          Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
        end
        else
        begin
          Grid.Canvas.Brush.Color := clWindow; // Default background color for cells
          Grid.Canvas.Font.Color := clWindowText; // Default text color for cells
          Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
        end;
        DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end
      else if (ARow > 0) and (ACol = InfoColumnIndex) then
      begin
        // Check If Info no blank
        if S <> '' then
        begin
          Grid.Canvas.Brush.Color := clYellow; // Entire cell background color for "NG" cells
          Grid.Canvas.Font.Color := clRed; // Text color for "NG" cells
          Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
          // Align the text to the left with padding
          Inc(Rect.Left, CellPadding);
          DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE);
        end;
      end
      else
      begin
        // Default cell formatting
        Grid.Canvas.Brush.Color := clWindow; // Default background color for cells
        Grid.Canvas.Font.Color := clWindowText; // Default text color for cells
        Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
        // Adjust the Rect to add padding on the left
        Inc(Rect.Left);
        DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE);
      end;
    end;

procedure TForm1.CreateStringGrid(var Grid: TStringGrid; AParent: TWinControl);
begin
  // Assign the OnDrawCell event handler
  Grid.OnDrawCell := StringGridCSVDrawCell;

  // Set number of columns
  Grid.ColCount := 22;
  Grid.RowCount := 1;


  // Set the headers
  Grid.Cells[0, 0] := 'Result';
  Grid.Cells[1, 0] := 'Date';
  Grid.Cells[2, 0] := 'Name';
  Grid.Cells[3, 0] := 'Employee Code';
  Grid.Cells[4, 0] := 'Code D';
  Grid.Cells[5, 0] := 'Mold Code';
  Grid.Cells[6, 0] := 'Model';
  Grid.Cells[7, 0] := 'Lamp Name';
  Grid.Cells[8, 0] := 'Part Name';
  Grid.Cells[9, 0] := 'Modify Job No.';
  Grid.Cells[10, 0] := 'Part Code';
  Grid.Cells[11, 0] := 'Start';
  Grid.Cells[12, 0] := 'Finish';
  Grid.Cells[13, 0] := 'Min';
  Grid.Cells[14, 0] := 'M/C Code';
  Grid.Cells[15, 0] := 'Mach.master';
  Grid.Cells[16, 0] := 'Date';
  Grid.Cells[17, 0] := 'Strat';
  Grid.Cells[18, 0] := 'Date';
  Grid.Cells[19, 0] := 'Finish';
  Grid.Cells[20, 0] := 'Min';
  Grid.Cells[21, 0] := 'ATC';
  Grid.Cells[22, 0] := 'Remark';
  // Set column widths for result column
  Grid.ColWidths[0] := 150;
  Grid.ColAlignments[0] := taCenter;

  // Set grid options to show lines
  Grid.Options := Grid.Options + [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];
    // Responsive layout management
  Grid.Anchors := [akLeft,akTop,akRight,akBottom];

end;


end.

