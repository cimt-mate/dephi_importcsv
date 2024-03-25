unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Uni, UniProvider, OracleUniProvider, MemDS, Vcl.Grids, Vcl.DBGrids,
  DBAccess, Vcl.ExtDlgs,System.IniFiles ,DateUtils, System.ImageList,ImportSetting,
  Vcl.ImgList, Vcl.Buttons, Vcl.Menus, Winapi.Winsock   , System.Types  ,System.IOUtils,System.StrUtils
  ,IpHlpApi,IpTypes, Vcl.ButtonGroup, Vcl.ToolWin ,System.RegularExpressions

;

type
  TForm1 = class(TForm)
    FolderDialog: TFileOpenDialog;
    EditFolderPath: TEdit;
    StringGridCSV: TStringGrid;
    SpeedButtonIMP: TSpeedButton;
    UniConnection: TUniConnection;
    OracleUniProvider: TOracleUniProvider;
    UniQuery: TUniQuery;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    ImageList1: TImageList;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    ProgressBar1: TProgressBar;
    LabelPath: TLabel;
    MainMenu1: TMainMenu;
    file1: TMenuItem;
    file2: TMenuItem;
    Help1: TMenuItem;
    Help2: TMenuItem;
    Whatisthis1: TMenuItem;
    procedure OpenFolderPathClick(Sender: TObject);
    procedure ButtonReadClick(Sender: TObject);
    procedure SpeedButtonIMPClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure file2Click(Sender: TObject);
    procedure Help2Click(Sender: TObject);
    procedure Whatisthis1Click(Sender: TObject);
    procedure SetIndex;
    procedure FormResize(Sender: TObject);
    procedure AdjustLastColumnWidth(Grid: TStringGrid);
    procedure Managefile;
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
    procedure CheckGRDFolder;
    function GetStringGridRowData(Grid: TStringGrid; RowIndex: Integer): String;
    function GetCellValueByColumnName(StringGrid: TStringGrid; HeaderName: string; Row: Integer): string;
    function CalculateWorkingTime(StartTimeStr, EndTimeStr,Shift: string): String;
    function MaxDateTime(const A, B: TDateTime): TDateTime;
    function MinDateTime(const A, B: TDateTime): TDateTime;
    function IsMaxTime(CellValue1, CellValue2: string): string;
    function MaxFloat(const A, B: Double): Double;
    function MinFloat(const A, B: Double): Double;
    procedure UpdateErrorColumn(Row: Integer; ErrorMessage: string);
    procedure ClearStringGrid(Grid: TStringGrid);
    function IsValidTimeFormat(TimeStr: string): Boolean;
  end;

var
  Form1: TForm1;
  Result,Shift_n,Date,WorkerName,EmployeeCode,CodeD,CostProcessName,MoldCode,Model,LampName,PartName : Integer;
  ModifyJobNo,PartCode,PartMaster,Start,Finish,Min,MCCode,Machmaster,MachStart,MachDate,MachFinish,MachMin,ATC,Remark,Status : Integer;
  CodeA,CodeB,CodeC : Integer;
  MovePath : String;
  Operation,HasErrorFileChoice,HasLogFile,ErrorPath,FolderPath : String;
  CurrentDateTime: TDateTime;
implementation

{$R *.dfm}



function GetMACAddress: string;
var
  AdapterInfo: PIP_ADAPTER_INFO;
  BufLen: ULONG;
  pAdapter: PIP_ADAPTER_INFO;
begin
  Result := '';
  BufLen := 0;
  // First call to get the buffer length
  GetAdaptersInfo(nil, BufLen);
  if BufLen = 0 then
    Exit;

  // Allocate memory for the buffer
  GetMem(AdapterInfo, BufLen);
  try
    // Second call to get the adapter information
    if GetAdaptersInfo(AdapterInfo, BufLen) = ERROR_SUCCESS then
    begin
      pAdapter := AdapterInfo;
      // Iterate through all adapters and get the first non-zero MAC address
      while pAdapter <> nil do
      begin
        if pAdapter^.AddressLength = 6 then // Check for valid MAC address length
        begin
          Result := Format('%.2x-%.2x-%.2x-%.2x-%.2x-%.2x',
            [pAdapter^.Address[0], pAdapter^.Address[1], pAdapter^.Address[2],
            pAdapter^.Address[3], pAdapter^.Address[4], pAdapter^.Address[5]]);
          Break; // Exit loop on first valid MAC address
        end;
        pAdapter := pAdapter^.Next;
      end;
    end;
  finally
    FreeMem(AdapterInfo); // Free the allocated memory
  end;

  if Result = '' then
    Result := '00-00-00-00-00-00'; // Default MAC address if none found
end;


function GetWindowsUserName: string;
var
  UserName: array [0..MAX_PATH] of Char;
  Size: DWORD;
begin
  Size := MAX_PATH;
  if GetUserName(UserName, Size) then
    Result := UserName
  else
    Result := '';
end;

function GetFileVersion(const FileName: string): string;
var
  Size, Handle: DWORD;
  Buffer: array of Byte;
  FixedPtr: PVSFixedFileInfo;
begin
  Size := GetFileVersionInfoSize(PChar(FileName), Handle);
  if Size > 0 then
  begin
    SetLength(Buffer, Size);
    if GetFileVersionInfo(PChar(FileName), Handle, Size, Buffer) and
       VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    begin
      Result := Format('%d.%d.%d.%d',
        [HiWord(FixedPtr^.dwFileVersionMS), LoWord(FixedPtr^.dwFileVersionMS),
         HiWord(FixedPtr^.dwFileVersionLS), LoWord(FixedPtr^.dwFileVersionLS)]);
    end;
  end
  else
    Result := '';
end;

procedure TForm1.ReadSettings;
var
  IniFile: TIniFile;
  IniFileName: string;
  Choice: string;
begin
 // Read ini file
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'GRD\' + ChangeFileExt(ExtractFileName(Application.ExeName),'') + '.ini');
  try
    EditFolderPath.Text := IniFile.ReadString('Settings', 'FolderPath', '');
    FolderPath := IniFile.ReadString('Settings', 'FolderPath', '');
    ErrorPath := IniFile.ReadString('Settings', 'ErrorPath', '');
    MovePath :=  IniFile.ReadString('Settings', 'MovePath', '');
    Operation :=  IniFile.ReadString('Settings', 'Operation', '');
    HasLogFile :=  IniFile.ReadString('Settings', 'HasLogFile', '');
    HasErrorFileChoice :=  IniFile.ReadString('Settings', 'HasErrorFile', '');
  finally
    IniFile.Free;
  end;
end;

procedure TForm1.CheckGRDFolder;
var
  IniFileName: string;
  GRDFolder: string;
begin
  // Get the folder path
  GRDFolder := ExtractFilePath(Application.ExeName) + 'GRD';

  // Check if the folder exists, if not, create it
  if not DirectoryExists(GRDFolder) then
  begin
    if not CreateDir(GRDFolder) then
    begin
      // Handle the error if the folder cannot be created
      MessageBox(0, 'Unable to create the GRD directory.', 'Error', MB_OK or MB_ICONERROR);
      Exit;
    end;
  end;
end;


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //close
  CheckGRDFolder;
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
    StringGridCSV.Anchors := [akLeft, akTop, akRight, akBottom];
  end;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
   CheckGRDFolder;
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

procedure TForm1.Whatisthis1Click(Sender: TObject);
begin
  ShowMessage('TKOITO IMPORT ACTUAL')
end;

procedure TForm1.WriteLog(const LogMessage: string);
var
  LogFileName: string;
  LogFile: TextFile;
  LineCount: Integer;
  TempList: TStringList;
begin
  if HasErrorFileChoice = '1' then
  begin
        LogFileName := ErrorPath + '/KT10IMP100_log.log';

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

end;

procedure Tform1.Managefile;
var
  Files: TStringDynArray;
  FileName: string;
begin
  if HasLogFile = '1' then
    begin
        if Operation = 'Move' then
        begin
          // Get all files in the source directory
          Files := TDirectory.GetFiles(FolderPath);

          // Ensure the destination directory exists
          if not TDirectory.Exists(MovePath) then
            TDirectory.CreateDirectory(MovePath);

          // Move each file to the destination directory
          for FileName in Files do
            TFile.Move(FileName, TPath.Combine(MovePath, ExtractFileName(FileName)));
        end
        else
        begin
          // MovePath is empty, delete all files in the source directory
            for FileName in Files do
              TFile.Delete(FileName);
        end;
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

procedure TForm1.SpeedButton2Click(Sender: TObject);
    begin
      Form2 := TForm2.Create(Application);
      Form2.Show;
      SpeedButton1.Enabled := true;
    end;

procedure TForm1.SpeedButtonIMPClick(Sender: TObject);
var
  i,kmseqno, jdseqno: Integer;
  // Path Value
  CSVFolderPath, LogFolderPath, ErrorLogFolderPath : String;
  IsInserted, HasErrorLog, HasLog, IsLogDelete, updateYMDS: Boolean;
  ErrorLog: TStringList;
  CurFileName, ErrorFileName: string;
  IniFile: TIniFile;
  FmtSettings: TFormatSettings;

  procedure InitializeFromIniFile;
  begin
    with TIniFile.Create(ExtractFilePath(Application.ExeName) + 'GRD\' + ChangeFileExt(ExtractFileName(Application.ExeName),'') + '.ini') do
    try
      CSVFolderPath := ReadString('Settings', 'FolderPath', '');
      HasLog := ReadBool('Settings', 'HasLogFile', False);
      HasErrorLog := ReadBool('Settings', 'HasErrorFile', False);
      IsLogDelete := (ReadString('Settings', 'Operation', '') = 'Delete');
      LogFolderPath := ReadString('Settings', 'MovePath', '');
      ErrorLogFolderPath := ReadString('Settings', 'ErrorPath', '');
    finally
      Free;
    end;
  end;

  procedure HandleFileOperations;
  begin
    if HasLog then
    begin
      if CurFileName <> '' then
      begin
        if IsLogDelete then
          DeleteFile(CSVFolderPath + '\' + CurFileName)
        else
          RenameFile(CSVFolderPath + '\' + CurFileName, LogFolderPath + '\' + FormatDateTime('yyyymmddHHmm', Now) + '_Log_' + CurFileName);
      end;
    end;
  end;

  procedure LogError;
  begin
    if HasErrorLog then
    begin
      if ErrorLog.Count	> 1 then
       begin
        ErrorLog.SaveToFile(ErrorLogFolderPath + '\' + FormatDateTime('yyyymmddHHmm', Now) + '_Error_' + CurFileName);
        ErrorLog.Clear;
        ErrorLog.Add(GetStringGridRowData(StringGridCSV, 0));
       end;
    end;
  end;

  begin
    ErrorLog := TStringList.Create;
    try
      InitializeFromIniFile;
      ErrorLog.Add(GetStringGridRowData(StringGridCSV, 0));
      CurFileName := '';
      // Prepare Date Format Check
      FmtSettings := TFormatSettings.Create;
      FmtSettings.ShortDateFormat := 'dd/mm/yyyy'; // Specify the expected format
      FmtSettings.DateSeparator := '/';
      FormatSettings.LongTimeFormat := 'hh:nn';
      FormatSettings.TimeSeparator := ':';
      // Setup UniQuery using the established connection
      try
        for i := 1 to StringGridCSV.RowCount - 1 do // Assuming the first row contains headers
        begin
          if GetCellValueByColumnName(StringGridCSV, 'Result', i) = 'NG' then
          begin

            // Case first row
            if CurFileName = '' then
              CurFileName := GetCellValueByColumnName(StringGridCSV, 'Filename', i);

            // Case Change File Name then Create Error Log File And Move File
            if CurFileName <> GetCellValueByColumnName(StringGridCSV, 'Filename', i) then
            begin
              HandleFileOperations;
              LogError;
              CurFileName := GetCellValueByColumnName(StringGridCSV, 'Filename', i);
            end;

            // Add Error Row Data into ErrorLog StringList
            ErrorLog.Add(GetStringGridRowData(StringGridCSV,i));
          end
          else
          begin
          try
              StringGridCSV.Cells[GetColumnIndexByHeaderName(StringGridCSV, 'Result'), i] := 'Imported';
          except
            on E: Exception do
            begin
              StringGridCSV.Cells[GetColumnIndexByHeaderName(StringGridCSV, 'Result'), i] := 'NG Import';
              ErrorLog.Add(GetStringGridRowData(StringGridCSV,i) + ',' +  E.Message);
            end;
          end;

          end;
        end;
        HandleFileOperations;
        LogError;
        ImportDataToDatabase;
      finally

      end;
    finally
      ErrorLog.Free;
      SpeedButton1.Enabled := true;
    end;
end;

function TForm1.GetStringGridRowData(Grid: TStringGrid; RowIndex: Integer): String;
var
  ColIndex: Integer;
  RowData: String;
begin
  RowData := '';
  // Loop through all columns in the row
  for ColIndex := 0 to Grid.ColCount - 1 do
  begin
    // Concatenate the column data with a comma, but skip the last comma
    RowData := RowData + Grid.Cells[ColIndex, RowIndex];
    if ColIndex < Grid.ColCount - 1 then
      RowData := RowData + ',';
  end;
  Result := RowData;
end;

function TForm1.GetCellValueByColumnName(StringGrid: TStringGrid; HeaderName: string; Row: Integer): string;
var
  ColIndex: Integer;
begin
  Result := ''; // Default result if header not found or row is out of range
  if (Row < 0) or (Row >= StringGrid.RowCount) then Exit;

  ColIndex := GetColumnIndexByHeaderName(StringGrid, HeaderName);
  if ColIndex >= 0 then
  begin
    Result := StringGrid.Cells[ColIndex, Row];
  end;
end;



procedure TForm1.Timer1Timer(Sender: TObject);
begin
  // Update only the panel for current time
  StatusBar1.Panels[StatusBar1.Panels.Count - 1].Text := DateTimeToStr(Now);
end;


procedure TForm1.ButtonReadClick(Sender: TObject);
begin
  ClearStringGrid(StringGridCSV);
  LoadCSVFilesIntoGrid(EditFolderPath.Text);
  SpeedButton1.Enabled := false;

end;

procedure TForm1.ImportDataToDatabase;
var
  IniFile: TIniFile;
  IniFileName: string;
  CD2Value: string;
  Row, Col: Integer;
  Value1, Value2, JhValue, Sagyoh, Kikaikadoh , MinMan,MinMach: Integer;
  InsertQuery: TUniQuery;
  SQL, SeizonoValue, BucdValue,MachValue, BunmValue: string;
  Gkoteicd, Kikaicd, Jigucd, Tantocd, Ymds ,Ymde, Bikou, Jisekibikou: string;
  MaxTime, FormattedDateTime,FormattedDateEnd,time: string;
  MaxJDSEQNO, NewJDSEQNO, GHIMOKUCDValue: Integer;
  Tourokuymd: TDateTime;
  YujintankaValue, KikaitankaValue, KoteitankaValue, YujinkinValue, MujinkinValue, KinsumValue: Double;
  CompName, MACAddr, WinUserName, ExeName, ExeVersion: string;
  Buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: DWORD;
  StartTime, EndTime, TimeDifference: Double;
  ResultDate: TDateTime;
  timeS,timeE,shift : string;
  DateValue,DateMachineValue: TDateTime;
  DateStr: string;
  TimeStr,TimeFinish,TimeStrMach,TimeFinishMach,DateMach: string;
begin
  // Load the database connection parameters
  LoadConnectionParameters;
  SetIndex;

  // Check if the UniConnection is connected
  if not UniConnection.Connected then
  begin
    ShowMessage('Error: Database connection .');
    Exit;
  end;

    // Initialize the progress bar
  ProgressBar1.Max := StringGridCSV.RowCount - 1;
  ProgressBar1.Position := 0;

  // Initialize the query component
  InsertQuery := TUniQuery.Create(nil);
  try
    InsertQuery.Connection := UniConnection;

    // Start a transaction
    UniConnection.StartTransaction;

    for Row := 1 to StringGridCSV.RowCount - 1 do
    begin
      try
        // Read ini file
            IniFileName := ExtractFilePath(Application.ExeName) + '/Setup/DRLOGIN.ini';
            if not FileExists(IniFileName) then
            begin
              WriteLog('INI file not found: ' + IniFileName);
              UpdateErrorColumn(Row, 'INI file not found');
              Continue; // Skip to the next iteration of the loop
            end;
            IniFile := TIniFile.Create(IniFileName);
            try
              CD2Value := IniFile.ReadString('TLogOnForm', 'CD2', '');
              if CD2Value = '' then
              begin
                WriteLog('CD2 value not found or INI file not read correctly.');
                UpdateErrorColumn(Row, 'CD2 value not found');
                Continue; // Skip to the next iteration of the loop
              end;
            finally
              IniFile.Free;
            end;
       //Get ComputerName,MacAddress,WindowsUsername,ExecutableName,Executable Version
            // Get Computer Name
            Size := MAX_COMPUTERNAME_LENGTH + 1;
            if not GetComputerName(Buffer, Size) then
            begin
              WriteLog('Failed to get computer name.');
              UpdateErrorColumn(Row, 'Failed to get computer name');
              Continue; // Skip to the next iteration of the loop
            end;
            CompName := Buffer;
            // Get MAC Address
            MACAddr := GetMACAddress;
            if MACAddr = '' then
            begin
              WriteLog('Failed to get MAC address.');
              UpdateErrorColumn(Row, 'Failed to get MAC address');
              Continue; // Skip to the next iteration of the loop
            end;
            // Get Windows Username
            WinUserName := GetWindowsUserName;
            if WinUserName = '' then
            begin
              WriteLog('Failed to get Windows username.');
              UpdateErrorColumn(Row, 'Failed to get Windows username');
              Continue; // Skip to the next iteration of the loop
            end;
            // Get Executable Name
            ExeName := ExtractFileName(Application.ExeName);
            if ExeName = '' then
            begin
              WriteLog('Failed to get executable name.');
              UpdateErrorColumn(Row, 'Failed to get executable name');
              Continue; // Skip to the next iteration of the loop
            end;
            // Get Executable Version
            ExeVersion := GetFileVersion(Application.ExeName);
            if ExeVersion = '' then
            begin
              WriteLog('Failed to get executable version.');
              UpdateErrorColumn(Row, 'Failed to get executable version');
              Continue; // Skip to the next iteration of the loop
            end;

       //prepare and validation Data
          // WorkerCD,Job,CostProcess,ymds not null
          SeizonoValue := StringGridCSV.Cells[ModifyJobNo, Row];   //seizo,modify job no.
          Gkoteicd := StringGridCSV.Cells[CodeD, Row];      //CostProcess CD,CodeD,gkoteicd
          Tantocd := StringGridCSV.Cells[EmployeeCode, Row];       //Employee CD,tantocd
          Ymds := StringGridCSV.Cells[Date, Row];          //ymds , date start
          BucdValue := StringGridCSV.Cells[PartCode, Row]; //partcd

         //Check Date format DD/MM/YYYY
         DateStr := StringGridCSV.Cells[Date, Row]; // 'Date' should be replaced with the actual index of your date column
         if (DateStr = '') and not TryStrToDate(StringGridCSV.Cells[Date, Row], DateValue) then
            begin
              UpdateErrorColumn(Row, 'Invalid or missing date format');
              UpdateResultColumn(Row, 'NG');
              // You can also log the error, update a status column, etc.
              Continue; // Skip to the next iteration of the loop
            end;


         try
         //WorkerCD
         InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM tantomst WHERE tantocd = :tantocd';
         InsertQuery.ParamByName('tantocd').AsString := Tantocd;
         InsertQuery.Open;
         if (Tantocd <> '') and (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'Employee Code is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'WorkerCD SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;


         try
         //Cost process CD
         InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM kouteigmst WHERE Gkoteicd = :Gkoteicd' ;
         InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
         InsertQuery.Open;
         if (Gkoteicd <> '') and (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'Code D is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
         except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'Cost process CD SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;

         try
         //Mfg. No.
         InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM SEIZOMST WHERE Seizono = :SeizonoValue' ;
         InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
         InsertQuery.Open;
         if (SeizonoValue <> '') and (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'Modify Job No is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;

         except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'ModifyJobNo SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;

         try
         //Part Code
         if BucdValue <> '' then
         begin
           InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM BUHINMST WHERE bucd = :BucdValue' ;
           InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
           InsertQuery.Open;
           if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
              begin
                UpdateErrorColumn(Row, 'PartCode is Invalid');
                UpdateResultColumn(Row, 'NG');
                Continue; // Skip to the next iteration of the loop
              end;
         end;
         except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'PART SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;

         TimeStr := StringGridCSV.Cells[Start, Row];
         TimeFinish := StringGridCSV.Cells[Finish, Row];
         TimeStrMach := StringGridCSV.Cells[MachStart, Row];
         TimeFinishMach := StringGridCSV.Cells[MachFinish, Row];
         DateMach :=  StringGridCSV.Cells[MachDate, Row];
        if not TryStrToInt(StringGridCSV.Cells[Min, Row], MinMan) then MinMan := 0;
        if not TryStrToInt(StringGridCSV.Cells[MachMin, Row], MinMach) then MinMach := 0;
        JhValue := MinMan + MinMach;
        // Calculate additional fields
        Sagyoh := MinMan + 0 + 0;
        Kikaikadoh := MinMan + MinMach + 0 + 0;

         // Machine Unman
         if (TimeStrMach <> '') and (TimeFinishMach <> '') then
         begin
             if not (IsValidTimeFormat(TimeStrMach)) then
             begin
              UpdateErrorColumn(Row, 'TimeMachStart is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;

             if not IsValidTimeFormat(TimeFinishMach) then
             begin
              UpdateErrorColumn(Row, 'TimeMachFinish is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;

             if (DateMach = '') and not TryStrToDate(StringGridCSV.Cells[MachDate, Row], DateMachineValue ) then
             begin
             UpdateErrorColumn(Row, 'MachDate is null');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;

             if MinMach = 0  then
             begin
              UpdateErrorColumn(Row, 'MinMach is null');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;
             bikou := StringGridCSV.Cells[MachMin, Row] ;
             FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime);  //lasted ymds
         end
         //Worker Manned
          else if (TimeStr <> '') and (TimeFinish <> '') then
         begin
             if not IsValidTimeFormat(TimeStr) then
             begin
              UpdateErrorColumn(Row, 'TimeStart is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;

             if not IsValidTimeFormat(TimeFinish) then
             begin
              UpdateErrorColumn(Row, 'TimeFinish is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;
             if MinMan = 0  then
             begin
              UpdateErrorColumn(Row, 'MinMan is null');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
             end;
            timeS    :=  StringGridCSV.Cells[Start, Row];
            timeE    :=  StringGridCSV.Cells[Finish, Row];
            shift := StringGridCSV.Cells[Shift_n, Row];
            bikou := CalculateWorkingTime(timeS,timeE,shift);  // min
            FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime);  //lasted ymds
         end
          else
         begin
              UpdateErrorColumn(Row, 'Time is Invalid');
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
         end;
         //calculate for check MinMan
         if JhValue = StrToInt(bikou) then
          begin
            bikou := '0';  //cal min for sure
          end;


          try
         //Machine CD
         MachValue := StringGridCSV.Cells[MCCode, Row]; //partcd
         if MachValue <> '' then
         begin
           InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM kikaimst WHERE kikaicd = :MachValue' ;
           InsertQuery.ParamByName('MachValue').AsString := MachValue;
           InsertQuery.Open;
           if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
              begin
                UpdateErrorColumn(Row, 'MachineCD is Invalid');
                UpdateResultColumn(Row, 'NG');
                Continue; // Skip to the next iteration of the loop
              end;
         end;
              except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'Machine SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;

         try
         //Jigucd ATC
         Jigucd := StringGridCSV.Cells[ATC, Row];  //ATC , Jigucd
         if Jigucd <> '' then
         begin
           InsertQuery.SQL.Text := 'SELECT COUNT(*) AS Count FROM JIGUMST WHERE Jigucd = :Jigucd' ;
           InsertQuery.ParamByName('Jigucd').AsString := Jigucd;
           InsertQuery.Open;
           if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
              begin
                UpdateErrorColumn(Row, 'ATC is Invalid');
                UpdateResultColumn(Row, 'NG');
                Continue; // Skip to the next iteration of the loop
              end;
         end;
         except
            on E: Exception do
            begin
              // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
              UpdateErrorColumn(Row, 'ATC SQL is Invalid');
              UpdateErrorColumn(Row, E.Message);
              UpdateResultColumn(Row, 'NG');
              Continue; // Skip to the next iteration of the loop
            end;
          end;
         //Remark
         Jisekibikou := StringGridCSV.Cells[Remark, Row];

        // Prepare data for insertion
        Tourokuymd := Now;

        //MANAGE Ymde DateEnd Just Date no time
        try
          // Ymde date end
          if StringGridCSV.Cells[MachDate, Row] = '' then
          begin
            if TryStrToFloat(StringGridCSV.Cells[Start, Row], StartTime) and TryStrToFloat(StringGridCSV.Cells[Finish, Row], EndTime) then
            begin
              if EndTime < StartTime then
              begin
                if TryStrToDate(StringGridCSV.Cells[Date, Row], ResultDate) then
                begin
                  ResultDate := ResultDate + 1; // Add 1 day
                  Ymde := FormatDateTime('dd/mm/yyyy', ResultDate);
                end
                else
                begin
                  Ymde := 'Invalid Date'; // Handle invalid date in cell [2, Row]
                end;
              end
              else
              begin
                Ymde := StringGridCSV.Cells[Date, Row];
              end;
            end;
          end
          else
          begin
            Ymde := StringGridCSV.Cells[MachDate, Row];
          end;
        //MANAGE Ymde TIMEEND
        time := GetMaxTime(StringGridCSV.Cells[Finish, Row], StringGridCSV.Cells[MachFinish, Row]);
        //MANAGE YMDE DATETIME
        FormattedDateEnd := FormatDateTimeStr(Ymde, time); // lasted ymde
      except
        on E: Exception do
        begin
          // Handle any exceptions that occur during date processing
            WriteLog('Error: Ymde is missing in row ' + IntToStr(Row));
            UpdateResultColumn(Row, 'NG');
            UpdateErrorColumn(Row, 'Ymde : '+E.Message); // Update the error column with the error message
            Continue; // Skip to the next iteration of the loop
        end;
      end;
        //Cost Unit Price
          // Get yujintanka value from tantomst
          InsertQuery.SQL.Text := 'SELECT tanka1 FROM tantomst WHERE tantocd = :tantocd';
          InsertQuery.ParamByName('tantocd').AsString := Tantocd;
          InsertQuery.Open;
          YujintankaValue := InsertQuery.FieldByName('tanka1').AsFloat;
          InsertQuery.Close;
          // Get Kikaitanka value from the kikaimst table using kikaicd
          InsertQuery.SQL.Text := 'SELECT KIKAITANKA FROM kikaimst WHERE kikaicd = :MachValue';
          InsertQuery.ParamByName('MachValue').AsString := MachValue;
          InsertQuery.Open;
          KikaitankaValue := InsertQuery.FieldByName('KIKAITANKA').AsFloat;
          InsertQuery.Close;
          // Get koteitanka value from the kouteigmst table using Gkoteicd
          InsertQuery.SQL.Text := 'SELECT GTANKA FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          KoteitankaValue := InsertQuery.FieldByName('GTANKA').AsFloat;
          InsertQuery.Close;
          // Retrieve GHIMOKUCD from KOUTEIGMST
          InsertQuery.SQL.Text := 'SELECT GHIMOKUCD FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          GHIMOKUCDValue := InsertQuery.FieldByName('GHIMOKUCD').AsInteger;
          InsertQuery.Close;
          // Calculate YujinkinValue, MujinkinValue, and KinsumValue
          YujinkinValue := MinMan * YujintankaValue / 60;
          MujinkinValue := MinMach * KikaitankaValue / 60;
          KinsumValue := YujinkinValue + MujinkinValue;

        //GET PRIMARY KEY
          // Get the maximum JDSEQNO from the JISEKIDATA table
          InsertQuery.SQL.Text := 'SELECT MAX(JDSEQNO) AS MaxJDSEQNO FROM JISEKIDATA';
          InsertQuery.Open;
          MaxJDSEQNO := InsertQuery.FieldByName('MaxJDSEQNO').AsInteger;
          InsertQuery.Close;
          // Increment the maximum JDSEQNO by 1 to get the new JDSEQNO
          NewJDSEQNO := MaxJDSEQNO + 1;
        //Check Error Insertion
        try
        // Construct and execute the SQL statement
        SQL := 'INSERT INTO JISEKIDATA (JDSEQNO, seizono, bunm, bucd, gkoteicd, kikaicd, jigucd, tantocd, ymds, KMSEQNO, jh, ' +
               'jmaedanh, jatodanh, jkbn, jyujinh, jmujinh, yujintanka, kikaitanka, koteitanka, GHIMOKUCD, yujinkin, ' +
               'mujinkin, kinsum, bikou, tourokuymd, sagyoh, kikaikadoh, inptantocd, inpymd, jisekibikou, inppcname, ' +
               'inpmacaddress, inpusername, inpexename, inpversion,ymde) ' +

               'VALUES (:NewJDSEQNO, :SeizonoValue, :BunmValue, :BucdValue, :Gkoteicd, :Kikaicd, :Jigucd, :Tantocd, ' +
               'TO_DATE(:FormattedDateTime, ''YYYY-MM-DD HH24:MI:SS''), 1, :JhValue, 0, 0, 4, :MinMan, :MinMach, ' +
               ':YujintankaValue, :KikaitankaValue, :KoteitankaValue, :GHIMOKUCDValue, :YujinkinValue, :MujinkinValue, ' +
               ':KinsumValue, :Bikou, :Tourokuymd, :Sagyoh, :Kikaikadoh, :InptantocdValue, :Inpymd, :Jisekibikou, ' +
               ':Inppcname, :Inpmacaddress, :Inpusername, :Inpexename, :Inpversion,TO_DATE(:FormattedDateEnd, ''YYYY-MM-DD HH24:MI:SS''))';

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
        InsertQuery.ParamByName('JhValue').AsInteger := JhValue;
        InsertQuery.ParamByName('MinMan').AsInteger := MinMan;
        InsertQuery.ParamByName('MinMach').AsInteger := MinMach;
        InsertQuery.ParamByName('YujintankaValue').AsFloat := YujintankaValue;
        InsertQuery.ParamByName('KikaitankaValue').AsFloat := KikaitankaValue;
        InsertQuery.ParamByName('KoteitankaValue').AsFloat := KoteitankaValue;
        InsertQuery.ParamByName('GHIMOKUCDValue').AsInteger := GHIMOKUCDValue;
        InsertQuery.ParamByName('YujinkinValue').AsFloat := YujinkinValue;
        InsertQuery.ParamByName('MujinkinValue').AsFloat := MujinkinValue;
        InsertQuery.ParamByName('KinsumValue').AsFloat := KinsumValue;
        InsertQuery.ParamByName('Bikou').AsString := Bikou;
        InsertQuery.ParamByName('Tourokuymd').AsDate := Tourokuymd;
        InsertQuery.ParamByName('Sagyoh').AsInteger := Sagyoh;
        InsertQuery.ParamByName('Kikaikadoh').AsInteger := Kikaikadoh;
        InsertQuery.ParamByName('InptantocdValue').AsString := CD2Value;
        InsertQuery.ParamByName('Inpymd').AsDateTime := Tourokuymd;
        InsertQuery.ParamByName('Jisekibikou').AsString := Jisekibikou;
        InsertQuery.ParamByName('Inppcname').AsString := CompName;
        InsertQuery.ParamByName('Inpmacaddress').AsString := MACAddr;
        InsertQuery.ParamByName('Inpusername').AsString := WinUserName;
        InsertQuery.ParamByName('Inpexename').AsString := ExeName;
        InsertQuery.ParamByName('Inpversion').AsString := ExeVersion;
        InsertQuery.ParamByName('FormattedDateEnd').AsString := FormattedDateEnd;
        InsertQuery.Execute;

        ProgressBar1.Position := Row;
        UpdateErrorColumn(Row, 'ID'+Inttostr(NewJDSEQNO) );
        except
            on E: Exception do
            begin
              // Log the error and update the "Result" column with "NG"
              WriteLog('Error SQL Insertion ' + IntToStr(Row) + ': ' + E.Message);
              UpdateResultColumn(Row, 'NG');
              UpdateErrorColumn(Row, 'Insertion : '+E.Message); // Update the error column with the error message
              continue;
            end;
          end;

      except
        on E: Exception do
        begin
          // Log the error and update the "Result" column with "NG"
          WriteLog('Error importing row ' + IntToStr(Row) + ': ' + E.Message);
          UpdateResultColumn(Row, 'NG');
          UpdateErrorColumn(Row, 'General : '+E.Message); // Update the error column with the error message
          continue;
        end;
      end;
    end;
    // Commit the transaction
    UniConnection.Commit;
  finally
    InsertQuery.Free;
    Managefile;
    ProgressBar1.Position := 0; // Reset the progress bar
  end;
end;

function TForm1.IsMaxTime(CellValue1, CellValue2: string): string;
var
  Time1, Time2: TDateTime;
begin
  Time1 := EncodeTime(StrToInt(Copy(CellValue1, 1, 2)), StrToInt(Copy(CellValue1, 4, 2)), 0, 0);
  Time2 := EncodeTime(StrToInt(Copy(CellValue2, 1, 2)), StrToInt(Copy(CellValue2, 4, 2)), 0, 0);

  if Time1 > Time2 then
    Result := FormatDateTime('hh.nn', Time1)
  else
    Result := FormatDateTime('hh.nn', Time2);
end;

function TForm1.MaxDateTime(const A, B: TDateTime): TDateTime;
begin
  if A > B then
    Result := A
  else
    Result := B;
end;

function TForm1.MinDateTime(const A, B: TDateTime): TDateTime;
begin
  if A < B then
    Result := A
  else
    Result := B;
end;

function TForm1.CalculateWorkingTime(StartTimeStr, EndTimeStr, Shift: string): String;
var
  StartTime, EndTime, BreakStart, BreakEnd, WorkingTime, OverlapTime: Double;
  BreakTimes: array[0..9] of Double;
  i: Integer;
  TotalMinutes: Double;
  CrossesMidnight: Boolean;
begin
  try
    // Convert string times to float
    StartTime := StrToInt(Copy(StartTimeStr, 1, 2)) * 60 + StrToInt(Copy(StartTimeStr, 4, 2));
    EndTime := StrToInt(Copy(EndTimeStr, 1, 2)) * 60 + StrToInt(Copy(EndTimeStr, 4, 2));

    // Check if the time range crosses midnight
    CrossesMidnight := EndTime < StartTime;
    if CrossesMidnight then
    begin
      // Add 24 hours (in minutes) to the end time if it crosses midnight
      EndTime := EndTime + 24 * 60;
    end;

    // Define break times based on the shift (in minutes)
    if Shift = 'D' then
    begin
      BreakTimes[0] := 12 * 60 + 10;
      BreakTimes[1] := 13 * 60 + 10;
      BreakTimes[2] := 17 * 60;
      BreakTimes[3] := 17 * 60 + 30;
      BreakTimes[4] := 21 * 60 + 10;
      BreakTimes[5] := 21 * 60 + 30;
      BreakTimes[6] := 25 * 60 + 30; // Next day
      BreakTimes[7] := 25 * 60 + 50; // Next day
      BreakTimes[8] := 29 * 60 + 50; // Next day
      BreakTimes[9] := 30 * 60 + 10; // Next day
    end
    else    // Night shift
    begin
      BreakTimes[0] := 0;
      BreakTimes[1] := 1 * 60;
      BreakTimes[2] := 5 * 60;
      BreakTimes[3] := 5 * 60 + 20;
      BreakTimes[4] := 9 * 60;
      BreakTimes[5] := 9 * 60 + 20;
      BreakTimes[6] := 13 * 60 + 20;
      BreakTimes[7] := 13 * 60 + 40;
      BreakTimes[8] := 17 * 60 + 40;
      BreakTimes[9] := 18 * 60;
    end;

    // Calculate the working time excluding breaks
    WorkingTime := EndTime - StartTime;

    // Calculate and subtract overlap with break times
    for i := Low(BreakTimes) to High(BreakTimes) div 2 do
    begin
      BreakStart := BreakTimes[i * 2];
      BreakEnd := BreakTimes[i * 2 + 1];

      // Adjust break times for cases where they cross midnight
      if BreakEnd < BreakStart then
        BreakEnd := BreakEnd + 24 * 60; // Add 24 hours (in minutes)

      // Calculate overlap with working time
      if (BreakStart < EndTime) and (BreakEnd > StartTime) then
      begin
        OverlapTime := MinFloat(EndTime, BreakEnd) - MaxFloat(StartTime, BreakStart);
        WorkingTime := WorkingTime - OverlapTime;
      end;
    end;

    // Convert working time to minutes
    TotalMinutes := Round(WorkingTime);
    Result := FloatToStr(TotalMinutes);
  except
    on E: Exception do
    begin
      // Log the error and return an empty string
      WriteLog('Error calculating working time: ' + E.Message);
      Result := '';
    end;
  end;

end;

function TForm1.MaxFloat(const A, B: Double): Double;
begin
  if A > B then
    Result := A
  else
    Result := B;
end;

function TForm1.MinFloat(const A, B: Double): Double;
begin
  if A < B then
    Result := A
  else
    Result := B;
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

procedure TForm1.Help2Click(Sender: TObject);
begin
  ShowMessage('CIM : cim_th_mail@cim.co.jp')
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

procedure TForm1.UpdateErrorColumn(Row: Integer; ErrorMessage: string);
var
  i: Integer;
begin
  if StringGridCSV.ColCount <= 28 then // Check if the error column (25) exists
  begin
    StringGridCSV.ColCount := 29; // Ensure there are at least 26 columns
  end;

  // Set the width of all columns except column 25 to 50
  for i := 0 to StringGridCSV.ColCount - 1 do
  begin
    if i <> 29 then
      StringGridCSV.ColWidths[i] := 50;
  end;

  // Update the error message in column 25
  StringGridCSV.Cells[Status, Row] := StringGridCSV.Cells[Status, Row]+',' + ErrorMessage;
  StringGridCSV.Anchors := [akLeft, akTop, akRight, akBottom];
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

procedure TForm1.ClearStringGrid(Grid: TStringGrid);
var
  Col, Row: Integer;
begin
  for Col := 0 to Grid.ColCount - 1 do
    for Row := 1 to Grid.RowCount - 1 do // Start from 1 to keep the headers
      Grid.Cells[Col, Row] := '';
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
var
  i: Integer;
begin
  // Assign the OnDrawCell event handler
  Grid.OnDrawCell := StringGridCSVDrawCell;

  // Set number of columns
  Grid.ColCount := 29;
  Grid.RowCount := 1;
  SetIndex;


  for i := 1 to Grid.ColCount - 1 do
  begin
    Grid.ColWidths[i] := 50;
    Grid.ColAlignments[i] := taCenter;
    Grid.ColWidths[Status] := 250;
  end;

  // Set the headers
  Grid.Cells[Result, 0] := 'Result';
  Grid.Cells[Shift_n, 0] := 'Shift';
  Grid.Cells[Date, 0] := 'Date';
  Grid.Cells[WorkerName, 0] := 'Worker Name';
  Grid.Cells[EmployeeCode, 0] := 'Employee Code';
  Grid.Cells[CodeD, 0] := 'Code D';
  Grid.Cells[CostProcessName, 0] := 'Cost Process Name';
  Grid.Cells[MoldCode, 0] := 'Mold Code';
  Grid.Cells[Model, 0] := 'Model';
  Grid.Cells[LampName, 0] := 'Lamp Name';
  Grid.Cells[PartName, 0] := 'Part Name';
  Grid.Cells[ModifyJobNo, 0] := 'Modify Job No.';
  Grid.Cells[PartCode, 0] := 'Part Code';
  Grid.Cells[PartMaster, 0] := 'Part Master';
  Grid.Cells[Start, 0] := 'Start';
  Grid.Cells[Finish, 0] := 'Finish';
  Grid.Cells[Min, 0] := 'Min';
  Grid.Cells[MCCode, 0] := 'M/C Code';
  Grid.Cells[Machmaster, 0] := 'Mach.master';
  Grid.Cells[MachStart, 0] := 'Start';
  Grid.Cells[MachDate, 0] := 'Date';
  Grid.Cells[MachFinish, 0] := 'Finish';
  Grid.Cells[MachMin, 0] := 'Min';
  Grid.Cells[ATC, 0] := 'ATC';
  Grid.Cells[Remark, 0] := 'Remark';
  Grid.Cells[Status, 0] := 'Status';
  Grid.Cells[CodeA, 0] := 'CodeA';
  Grid.Cells[CodeB, 0] := 'CodeB';
  Grid.Cells[CodeC, 0] := 'CodeC';
  // Set column widths for result column
  Grid.ColWidths[0] := 150;
  Grid.ColAlignments[0] := taCenter;

  // Set grid options to show lines
  Grid.Options := Grid.Options + [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];

  // Responsive layout management
  Grid.Anchors := [akLeft, akTop, akRight, akBottom];

  // Optionally set the grid's parent
  Grid.Parent := AParent;

end;

procedure TForm1.SetIndex;
begin
  Result := 0;
  Shift_n := 1;
  Date := 2;
  WorkerName := 3;
  EmployeeCode := 4;
  CodeD := 5;
  CostProcessName := 6;
  MoldCode := 7;
  Model := 8;
  LampName := 9;
  PartName := 10;
  ModifyJobNo := 11;
  PartCode := 12;
  PartMaster := 13;
  Start := 14;
  Finish := 15;
  Min := 16;
  MCCode := 17;
  Machmaster := 18;
  MachStart := 19;
  MachDate := 20;
  MachFinish := 21;
  MachMin := 22;
  ATC := 23;
  Remark := 24;
  Status :=28;
  CodeA := 26;
  CodeB := 27;
  CodeC :=25;

end;

procedure TForm1.AdjustLastColumnWidth(Grid: TStringGrid);
var
  TotalWidth, OtherColsWidth, i: Integer;
begin
  OtherColsWidth := 0;
  for i := 0 to Grid.ColCount - 2 do // Exclude the last column
    Inc(OtherColsWidth, Grid.ColWidths[i]);

  TotalWidth := Grid.ClientWidth - OtherColsWidth - Grid.GridLineWidth * (Grid.ColCount - 1);
  if TotalWidth > 0 then
    Grid.ColWidths[Grid.ColCount - 1] := TotalWidth;
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  AdjustLastColumnWidth(StringGridCSV);
end;

function TForm1.IsValidTimeFormat(TimeStr: string): Boolean;
var
  RegEx: TRegEx;
  Match: TMatch;
  Hours, Minutes: Integer;
begin
  Result := False;
  RegEx := TRegEx.Create('^(?<hours>\d{2})\.(?<minutes>\d{2})$');
  Match := RegEx.Match(TimeStr);

  if Match.Success then
  begin
    if TryStrToInt(Match.Groups['hours'].Value, Hours) and
       TryStrToInt(Match.Groups['minutes'].Value, Minutes) then
    begin
      Result := (Hours >= 0) and (Hours <= 23) and (Minutes >= 0) and (Minutes <= 59);
    end;
  end;
end;

procedure MoveAllFiles;
var
  Files: TStringDynArray;
  FileName, NewFileName, DateTimeStr: string;
begin
  // Get all files in the source directory
  Files := TDirectory.GetFiles(FolderPath);

  // Format the current date and time as a string
  DateTimeStr := FormatDateTime('yyyy-mm-dd_hhnnss', Now);
  if HasLogFile = '1' then
  begin
         if Operation = 'Move' then
          begin
            // Ensure the destination directory exists
            if not TDirectory.Exists(MovePath) then
              TDirectory.CreateDirectory(MovePath);

            // Move each file to the destination directory with the date and time appended to the filename
            for FileName in Files do
            begin
              NewFileName := TPath.Combine(MovePath, TPath.GetFileNameWithoutExtension(FileName) + '_' + DateTimeStr + TPath.GetExtension(FileName));
              TFile.Move(FileName, NewFileName);
            end;
          end
          else
          begin
            // MovePath is empty, delete all files in the source directory
            for FileName in Files do
              TFile.Delete(FileName);
          end;
  end;


end;


procedure TForm1.file2Click(Sender: TObject);
begin
  halt;
end;

end.

