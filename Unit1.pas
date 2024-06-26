unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Uni, UniProvider, OracleUniProvider, MemDS,
  Vcl.Grids, Vcl.DBGrids,
  DBAccess, Vcl.ExtDlgs, System.IniFiles, DateUtils, System.ImageList,
  ImportSetting,
  Vcl.ImgList, Vcl.Buttons, Vcl.Menus, Winapi.Winsock, System.Types,
  System.IOUtils, System.StrUtils, IpHlpApi, IpTypes, Vcl.ButtonGroup,
  Vcl.ToolWin, System.RegularExpressions, Math ,Clipbrd

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
    Refresh1: TMenuItem;
    Copytocsv1: TMenuItem;
    SpeedButton3: TSpeedButton;
    LabelOK: TLabel;
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
    procedure CheckValues;
    procedure LogRowToCSV(Row: Integer; ErrorMessage: string);
    procedure CalculateMinutesDifference(const StartTime, EndTime: String; out TotalMinutes: Integer);
    procedure Refresh1Click(Sender: TObject);
    procedure Copytocsv1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    procedure LoadCSVFilesIntoGrid(const FolderPath: string);
    procedure ImportDataToDatabase;
    procedure SetupDatabaseQuery;
    procedure LoadConnectionParameters;
    procedure WriteLog(const LogMessage: string);
    function GetProgramName: string;
    function GetAppVersion: string;
    procedure ReadSettings;
    procedure CreateStringGrid(var Grid: TStringGrid; AParent: TWinControl);
    procedure StringGridCSVDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    function GetColumnIndexByHeaderName(StringGrid: TStringGrid;
      HeaderName: string): Integer;
    procedure UpdateResultColumn(Row: Integer; const ResultText: string);
    function GetBunmFromBucd(const BucdValue: string): string;
    function FormatDateTimeStr(const DateStr, TimeStr: string): string;
    function GetMaxTime(const Time1, Time2: string): string;
    function GetTimeInMinutes(const TimeStr: string): Integer;
    procedure CheckGRDFolder;
    function GetStringGridRowData(Grid: TStringGrid; RowIndex: Integer): String;
    function GetCellValueByColumnName(StringGrid: TStringGrid;
      HeaderName: string; Row: Integer): string;
    function CalculateWorkingTime(StartTimeStr, EndTimeStr,
      Shift: string): String;
    function MaxDateTime(const A, B: TDateTime): TDateTime;
    function MinDateTime(const A, B: TDateTime): TDateTime;
    function IsMaxTime(CellValue1, CellValue2: string): string;
    function MaxFloat(const A, B: Double): Double;
    function MinFloat(const A, B: Double): Double;
    procedure UpdateErrorColumn(Row: Integer; ErrorMessage: string);
    procedure ClearStringGrid(Grid: TStringGrid);
    function RoundDownTo(Value: Double; Decimals: Integer): Double;
    function IsValidTimeFormat(TimeStr: string): Boolean;
    procedure LogErrorRowToCSV(Row: Integer; ErrorMessage: string);
  end;

var
  Form1: TForm1;
  Result, Shift_n, Date, WorkerName, EmployeeCode, CodeD, CostProcessName,
    MoldCode, Model, LampName, PartName: Integer;
  ModifyJobNo, PartCode, PartMaster, Start, Finish, Min, MCCode, Machmaster,
    MachStart, MachDate, MachFinish, MachMin, ATC, Remark, Status,
    filename: Integer;
  CodeA, CodeB, CodeC: Integer;
  MovePath, ErrorMessageText: String;
  Operation, HasErrorFileChoice, HasLogFile, ErrorPath, FolderPath, Error,
    PathErrorCSV, ResultPathCSV, HasResult , FileNameStr: String;
  CurrentDateTime: TDateTime;
  Hours, Minutes,TotalMinutes,RowCheck: Integer;

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
        if pAdapter^.AddressLength = 6 then
        // Check for valid MAC address length
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
  UserName: array [0 .. MAX_PATH] of Char;
  Size: DWORD;
begin
  Size := MAX_PATH;
  if GetUserName(UserName, Size) then
    Result := UserName
  else
    Result := '';
end;

function GetFileVersion(const filename: string): string;
var
  Size, Handle: DWORD;
  Buffer: array of Byte;
  FixedPtr: PVSFixedFileInfo;
begin
  Size := GetFileVersionInfoSize(PChar(filename), Handle);
  if Size > 0 then
  begin
    SetLength(Buffer, Size);
    if GetFileVersionInfo(PChar(filename), Handle, Size, Buffer) and
      VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    begin
      Result := Format('%d.%d.%d.%d', [HiWord(FixedPtr^.dwFileVersionMS),
        LoWord(FixedPtr^.dwFileVersionMS), HiWord(FixedPtr^.dwFileVersionLS),
        LoWord(FixedPtr^.dwFileVersionLS)]);
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
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'GRD\' +
    ChangeFileExt(ExtractFileName(Application.ExeName), '') + '.ini');
  try
    EditFolderPath.Text := IniFile.ReadString('Settings', 'FolderPath', '');
    FolderPath := IniFile.ReadString('Settings', 'FolderPath', '');
    ErrorPath := IniFile.ReadString('Settings', 'ErrorPath', '');
    MovePath := IniFile.ReadString('Settings', 'MovePath', '');
    Operation := IniFile.ReadString('Settings', 'Operation', '');
    HasLogFile := IniFile.ReadString('Settings', 'HasLogFile', '');
    Error := IniFile.ReadString('Settings', 'Error', '');
    PathErrorCSV := IniFile.ReadString('Settings', 'PathErrorCSV', '');
    HasErrorFileChoice := IniFile.ReadString('Settings', 'HasErrorFile', '');
    ResultPathCSV := IniFile.ReadString('Settings', 'ResultPath', '');
    HasResult := IniFile.ReadString('Settings', 'HasResult', '');
  finally
    IniFile.Free;
  end;
end;

procedure TForm1.Refresh1Click(Sender: TObject);
begin
     //clear
    ClearStringGrid(StringGridCSV);
    SpeedButton1.Enabled := True;
    SpeedButtonIMP.Enabled := false;
     LabelOK.Caption := '';
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
      MessageBox(0, 'Unable to create the GRD directory.', 'Error',
        MB_OK or MB_ICONERROR);
      Exit;
    end;
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // close
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
  filename: string;
  // Can't declare Username, Password *Conflict with UnitConnection Variable Name
  DirectDBName, User, Pass: string;
begin
  CustomBlue := rgb(194, 209, 254); // Standard blue color
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
    filename := ExtractFilePath(Application.ExeName) + '/Setup/SetUp.Ini';
    // Assumes the INI file is in the same directory as the application
    IniFile := TIniFile.Create(filename);
    DirectDBName := IniFile.ReadString('Setting', 'DIRECTDBNAME', '');
    User := IniFile.ReadString('Setting', 'USERNAME', '');
    with StatusBar1.Panels.Add do
    begin
      Width := 300;
      Text := DirectDBName + ':' + User;
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
var
  i: Integer;
  AutoExportCheck: Boolean;
begin
  CheckGRDFolder;
  AutoExportCheck := False;
  for i := 1 to ParamCount do
  begin
    if UpperCase(ParamStr(i)) = '/AUTO' then
      begin
        try
           //Read
             LabelOK.Caption :=  '';
              ReadSettings;
              ClearStringGrid(StringGridCSV);
              LoadCSVFilesIntoGrid(EditFolderPath.Text);
              SpeedButton1.Enabled := false;
              CheckValues;
              SpeedButtonIMP.Enabled := True;
           //Import
              ImportDataToDatabase;
              SpeedButton1.Enabled := True;
              SpeedButtonIMP.Enabled := false;
           WriteLog('AutoExport Success');
          except
              on E: Exception do
              begin
                WriteLog('Error in AutoExport : ' + E.Message);
              end;
          end;
          Application.Terminate;
          Break;
      end;
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
    Result := Format('%d.%d.%d.%d', [HiWord(FixedPtr^.dwFileVersionMS),
      LoWord(FixedPtr^.dwFileVersionMS), HiWord(FixedPtr^.dwFileVersionLS),
      LoWord(FixedPtr^.dwFileVersionLS)])
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
  filename: string;
  // Can't declare Username, Password *Conflict with UnitConnection Variable Name
  DirectDBName, User, Pass: string;
begin
  filename := ExtractFilePath(Application.ExeName) + 'Setup\SetUp.Ini';
  // Use backslash for path in Windows
  // Check if the INI file exists
  if not FileExists(filename) then
  begin
    WriteLog('Error: INI file not found at ' + filename);
    // Replace WriteLog with your actual logging procedure
    Exit; // Exit the procedure if the file does not exist
  end;

  IniFile := TIniFile.Create(filename);
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
        UserName := User;
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
   CurrentDate: string;
begin
  if HasErrorFileChoice = '1' then
  begin

    CurrentDate := FormatDateTime('YYYYMMDD', Now);
    LogFileName := ErrorPath +'/'+ CurrentDate + '_KT10IMP100_log.log';

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

      Writeln(LogFile, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ': ' +
        LogMessage);
    finally
      CloseFile(LogFile);
    end;
  end;

end;

procedure TForm1.Managefile;
var
  Files: TStringDynArray;
  filename, newFilename, timestamp: string;
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

      // Move each file to the destination directory and rename it
      for filename in Files do
      begin
        timestamp := FormatDateTime('YYYYMMDDHHNN', Now);
        newFilename := 'Backup_' + timestamp + '_' + ExtractFileName(filename);
        TFile.Move(filename, TPath.Combine(MovePath, newFilename));
      end;
    end
    else
    begin
      // MovePath is empty, delete all files in the source directory
      Files := TDirectory.GetFiles(FolderPath);
      for filename in Files do
        TFile.Delete(filename);
    end;
  end;
end;

procedure TForm1.OpenFolderPathClick(Sender: TObject);
begin
  FolderDialog.Options := FolderDialog.Options + [fdoPickFolders];
  if FolderDialog.Execute then
  begin
    EditFolderPath.Text := FolderDialog.filename;
  end;
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  Form2 := TForm2.Create(Application);
  Form2.Show;
  SpeedButton1.Enabled := True;
end;

procedure TForm1.SpeedButton3Click(Sender: TObject);
begin
       //clear
     ClearStringGrid(StringGridCSV);
       SpeedButton1.Enabled := True;
      SpeedButtonIMP.Enabled := false;
      LabelOK.Caption := '';
end;

procedure TForm1.SpeedButtonIMPClick(Sender: TObject);
begin
  ImportDataToDatabase;
  SpeedButton1.Enabled := True;
  SpeedButtonIMP.Enabled := false;
end;

function TForm1.GetStringGridRowData(Grid: TStringGrid;
  RowIndex: Integer): String;
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

function TForm1.GetCellValueByColumnName(StringGrid: TStringGrid;
  HeaderName: string; Row: Integer): string;
var
  ColIndex: Integer;
begin
  Result := ''; // Default result if header not found or row is out of range
  if (Row < 0) or (Row >= StringGrid.RowCount) then
    Exit;

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
  LabelOK.Caption :=  '';
  ReadSettings;
  ClearStringGrid(StringGridCSV);
  LoadCSVFilesIntoGrid(EditFolderPath.Text);
  SpeedButton1.Enabled := false;
  CheckValues;
  SpeedButtonIMP.Enabled := True;
end;

procedure TForm1.CheckValues;
var
  OKCount, NGCount, i: Integer;
  IniFile: TIniFile;
  IniFileName: string;
  CD2Value: string;
  Row, Col: Integer;
  Value1, Value2, JhValue, Sagyoh, Kikaikadoh, MinMan, MinMach: Integer;
  InsertQuery: TUniQuery;
  SQL, SeizonoValue, BucdValue, MachValue, BunmValue: string;
  Gkoteicd, Kikaicd, Jigucd, Tantocd, Ymds, Ymde, Bikou, Jisekibikou: string;
  MaxTime, FormattedDateTime, FormattedDateEnd, time: string;
  MaxJDSEQNO, NewJDSEQNO, GHIMOKUCDValue: Integer;
  Tourokuymd: TDateTime;
  YujintankaValue, KikaitankaValue, KoteitankaValue, YujinkinValue,
    MujinkinValue, KinsumValue: Double;
  CompName, MACAddr, WinUserName, ExeName, ExeVersion: string;
  Buffer: array [0 .. MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: DWORD;
  StartTime, EndTime, TimeDifference: Double;
  ResultDate: TDateTime;
  timeS, timeE, Shift: string;
  DateValue, DateMachineValue: TDateTime;
  DateStr,Part_Name,Part_Master,Mold_Code,Model_info,Lamp_Name,MinManStr : string;
  TimeStr, TimeFinish, TimeStrMach, TimeFinishMach, DateMach , MinMachStr: string;
  num: Integer;
  FormattedKinsumValue: String;
   FormatSettings: TFormatSettings;
begin
  // Load the database connection parameters
  LoadConnectionParameters;
  SetIndex;
  OKCount := 0;
  NGCount := 0;
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
      num := 0;
      try
        if (StringGridCSV.Cells[Shift_n, Row] = '') and
           (StringGridCSV.Cells[Date, Row] = '') and
           (StringGridCSV.Cells[WorkerName, Row] = '') and
           (StringGridCSV.Cells[EmployeeCode, Row] = '') and
           (StringGridCSV.Cells[CodeD, Row] = '') and
           (StringGridCSV.Cells[CostProcessName, Row] = '') and
           (StringGridCSV.Cells[MoldCode, Row] = '') and
           (StringGridCSV.Cells[Model, Row] = '') and
           (StringGridCSV.Cells[LampName, Row] = '') and
           (StringGridCSV.Cells[PartName, Row] = '') and
           (StringGridCSV.Cells[ModifyJobNo, Row] = '') and
           (StringGridCSV.Cells[PartCode, Row] = '') and
           (StringGridCSV.Cells[PartMaster, Row] = '') and
           (StringGridCSV.Cells[Start, Row] = '') and
           (StringGridCSV.Cells[Finish, Row] = '') and
           (StringGridCSV.Cells[Min, Row] = '') and
           (StringGridCSV.Cells[MCCode, Row] = '') and
           (StringGridCSV.Cells[Machmaster, Row] = '') and
           (StringGridCSV.Cells[MachStart, Row] = '') and
           (StringGridCSV.Cells[MachDate, Row] = '') and
           (StringGridCSV.Cells[MachFinish, Row] = '') and
           (StringGridCSV.Cells[MachMin, Row] = '') and
           (StringGridCSV.Cells[ATC, Row] = '') and
           (StringGridCSV.Cells[Remark, Row] = '') and
           (StringGridCSV.Cells[CodeA, Row] = '') and
           (StringGridCSV.Cells[CodeC, Row] = '') and
           (StringGridCSV.Cells[CodeB, Row] = '') then
        begin
          continue;
          // Perform actions here if all cells are empty
        end
        else
         begin
           //

        end;

        IniFileName := ExtractFilePath(Application.ExeName) +
          '/Setup/DRLOGIN.ini';
        if not FileExists(IniFileName) then
        begin
          UpdateErrorColumn(Row, 'INI file not found');
          num := num + 1;
        end;
        IniFile := TIniFile.Create(IniFileName);
        try
          CD2Value := IniFile.ReadString('TLogOnForm', 'CD2', '');
          if CD2Value = '' then
          begin
            UpdateErrorColumn(Row, 'CD2 value not found');
            num := num + 1;
          end;
        finally
          IniFile.Free;
        end;
        // Get ComputerName,MacAddress,WindowsUsername,ExecutableName,Executable Version
        // Get Computer Name
        Size := MAX_COMPUTERNAME_LENGTH + 1;
        if not GetComputerName(Buffer, Size) then
        begin
          UpdateErrorColumn(Row, 'Failed to get computer name');
          num := num + 1;

        end;
        CompName := Buffer;
        // Get MAC Address
        MACAddr := GetMACAddress;
        if MACAddr = '' then
        begin
          UpdateErrorColumn(Row, 'Failed to get MAC address');
          num := num + 1;
        end;
        // Get Windows Username
        WinUserName := GetWindowsUserName;
        if WinUserName = '' then
        begin
          UpdateErrorColumn(Row, 'Failed to get Windows username');
          num := num + 1;
        end;
        // Get Executable Name
        ExeName := ExtractFileName(Application.ExeName);
        if ExeName = '' then
        begin
          UpdateErrorColumn(Row, 'Failed to get executable name');

          num := num + 1;
        end;
        // Get Executable Version
        ExeVersion := GetFileVersion(Application.ExeName);
        if ExeVersion = '' then
        begin
          UpdateErrorColumn(Row, 'Failed to get executable version');
          num := num + 1;
        end;

        // prepare and validation Data
        // WorkerCD,Job,CostProcess,ymds not null
        SeizonoValue := StringGridCSV.Cells[ModifyJobNo, Row];
        // seizo,modify job no.
        Gkoteicd := StringGridCSV.Cells[CodeD, Row];
        // CostProcess CD,CodeD,gkoteicd
        Tantocd := StringGridCSV.Cells[EmployeeCode, Row];
        // Employee CD,tantocd
        Ymds := StringGridCSV.Cells[Date, Row]; // ymds , date start
        BucdValue := StringGridCSV.Cells[PartCode, Row]; // partcd
        Part_Name := StringGridCSV.Cells[PartName, Row]; // partname
        Part_Master := StringGridCSV.Cells[PartMaster, Row];        //partmaster
        Mold_Code := StringGridCSV.Cells[MoldCode, Row]; //MoldCode
        Model_info := StringGridCSV.Cells[Model, Row]; //Model_info
        Lamp_Name := StringGridCSV.Cells[LampName, Row]; //LampName
        MinManStr := StringGridCSV.Cells[Min, Row];       //min man
        MinMachStr := StringGridCSV.Cells[MachMin, Row];  //min mach

        // Check Date format DD/MM/YYYY
        DateStr := StringGridCSV.Cells[Date, Row]; // 'Date' should be replaced with the actual index of your date column

        // Set up the format settings to match the desired date format
        FormatSettings := TFormatSettings.Create;
        FormatSettings.DateSeparator := '/';
        FormatSettings.ShortDateFormat := 'dd/mm/yyyy';

        // Try to convert the string to a TDateTime using the specified format
        if (DateStr = '') or not TryStrToDate(DateStr, DateValue, FormatSettings) then
        begin
          UpdateErrorColumn(Row, 'Invalid or Missing Date Format');
          UpdateResultColumn(Row, 'NG');
          num := num + 1; // Skip to the next iteration of the loop
        end
        else if not IsValidDate(YearOf(DateValue), MonthOf(DateValue), DayOf(DateValue)) then
        begin
          UpdateErrorColumn(Row, 'Invalid Date Values');
          UpdateResultColumn(Row, 'NG');
          num := num + 1;
        end;

        try
          // WorkerCD
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM tantomst WHERE tantocd = :tantocd';
          InsertQuery.ParamByName('tantocd').AsString := Tantocd;
          InsertQuery.Open;
          if (Tantocd = '') or (InsertQuery.FieldByName('Count').AsInteger
            <= 0) then
          begin
            UpdateErrorColumn(Row, 'Employee Code is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'WorkerCD SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;

        try
          // Cost process CD
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM kouteigmst WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          if (Gkoteicd = '') or (InsertQuery.FieldByName('Count').AsInteger
            <= 0) then
          begin
            UpdateErrorColumn(Row, 'Code D is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'Cost process CD SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;

        try
          // Mfg. No.
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM SEIZOMST WHERE Seizono = :SeizonoValue';
          InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
          InsertQuery.Open;
          if (SeizonoValue = '') or
            (InsertQuery.FieldByName('Count').AsInteger <= 0) then
          begin
            UpdateErrorColumn(Row, 'Modify Job No is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'ModifyJobNo SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;

        try
          // Part Code
          if BucdValue <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM BUHINMST WHERE bucd = :BucdValue';
            InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'PartCode is Invalid');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
            end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'PART SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;

        TimeStr := StringGridCSV.Cells[Start, Row];
        TimeFinish := StringGridCSV.Cells[Finish, Row];
        TimeStrMach := StringGridCSV.Cells[MachStart, Row];
        TimeFinishMach := StringGridCSV.Cells[MachFinish, Row];
        DateMach := StringGridCSV.Cells[MachDate, Row];

        if not TryStrToInt(StringGridCSV.Cells[Min, Row], MinMan) then
          MinMan := 0;
        if not TryStrToInt(StringGridCSV.Cells[MachMin, Row], MinMach) then
          MinMach := 0;
        JhValue := MinMan + MinMach;
        // Calculate additional fields
        Sagyoh := MinMan + 0 + 0;
        Kikaikadoh := MinMan + MinMach + 0 + 0;


        if MinManStr <> '' then
        begin
          if (TimeStr = '') or (TimeFinish = '') or (DateStr = '') then
          begin
              UpdateErrorColumn(Row, 'Time is Valid');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
          end;
          if (MinMachStr <> '') or (TimeStrMach <> '') or (TimeFinishMach <> '') or (DateMach <> '') then
          begin
              UpdateErrorColumn(Row, 'Contain both Man and Unman data');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
          end;
        end
        else if MinMachStr <> '' then
        begin
          if (TimeStrMach = '') or (TimeFinishMach = '') or (DateMach = '') then
            begin
                UpdateErrorColumn(Row, 'Time is Valid');
                UpdateResultColumn(Row, 'NG');
                num := num + 1;
            end;
          if (MinManStr <> '') or (TimeStr <> '') or (TimeFinish <> '') then
          begin
              UpdateErrorColumn(Row, 'Contain both Man and Unman data');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
          end;
        end;


        // Machine Unman
        if (TimeStrMach <> '') and (TimeFinishMach <> '') and (DateMach <> '') AND (TimeStr = '') and (TimeFinish = '')  then
        begin
          if not(IsValidTimeFormat(TimeStrMach)) then
          begin
            UpdateErrorColumn(Row, 'TimeMachStart is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if not IsValidTimeFormat(TimeFinishMach) then
          begin
            UpdateErrorColumn(Row, 'TimeMachFinish is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if (DateMach = '') and not TryStrToDate(StringGridCSV.Cells[MachDate,Row], DateMachineValue) then
          begin
            UpdateErrorColumn(Row, 'MachDate is null');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if MinMach = 0 then
          begin
            UpdateErrorColumn(Row, 'MinMach is null');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if (DateMach = DateStr) and ( strtofloat(TimeStrMach) > strtofloat(TimeFinishMach) ) then
          begin
            UpdateErrorColumn(Row, 'TimeStart&TimeEnd is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if (DateMach < DateStr) then
          begin
            UpdateErrorColumn(Row, 'TimeStart&TimeEnd is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          Bikou := StringGridCSV.Cells[MachMin, Row];
          MaxTime := GetMaxTime(StringGridCSV.Cells[Start, Row],
          StringGridCSV.Cells[MachStart, Row]);
          FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime); // lasted ymds
        end
        // Worker Manned
        else if (TimeStr <> '') and (TimeFinish <> '') AND (TimeStrMach = '') and (TimeFinishMach = '') and (DateMach = '')   then
        begin
          if not IsValidTimeFormat(TimeStr) then
          begin
            UpdateErrorColumn(Row, 'TimeStart is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          if not IsValidTimeFormat(TimeFinish) then
          begin
            UpdateErrorColumn(Row, 'TimeFinish is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
          if MinMan = 0 then
          begin
            UpdateErrorColumn(Row, 'MinMan is null');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
          timeS := StringGridCSV.Cells[Start, Row];
          timeE := StringGridCSV.Cells[Finish, Row];
          Shift := StringGridCSV.Cells[Shift_n, Row];
          if num = 0 then
          begin
            Bikou := CalculateWorkingTime(timeS, timeE, Shift); // minman
            TotalMinutes :=    StrtoInt(Bikou);
            MaxTime := GetMaxTime(StringGridCSV.Cells[Start, Row],
            StringGridCSV.Cells[MachStart, Row]);
            FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime); // lasted ymds
          end;
        end
        else
        begin
          UpdateErrorColumn(Row, 'Time is Invalid');
          UpdateResultColumn(Row, 'NG');
          num := num + 1;
        end;
        // calculate for check MinMan
        try
          // calculate for check MinMan
          if num = 0 then
          begin
            if JhValue = StrToInt(Bikou) then
            begin
              Bikou := '0'; // cal min for sure
            end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'Check JhMin is invalid');
            UpdateResultColumn(Row, 'NG');
            Continue; // Skip to the next iteration of the loop
          end;
        end;
        try
          // Machine CD
          MachValue := StringGridCSV.Cells[MCCode, Row]; // partcd
          if MachValue <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM kikaimst WHERE kikaicd = :MachValue';
            InsertQuery.ParamByName('MachValue').AsString := MachValue;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'MachineCD is Invalid');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
            end;
          end
          else  //MachValue = ''
          begin
               if (MinMachStr <> '') or (TimeStrMach <> '') or (TimeFinishMach <> '') or (DateMach <> '') then
                   begin
                        UpdateErrorColumn(Row, 'MachineCD is Null');
                        UpdateResultColumn(Row, 'NG');
                        num := num + 1; // Skip to the next iteration of the loop
                   end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'Machine SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;

        try
          // Jigucd ATC
          Jigucd := StringGridCSV.Cells[ATC, Row]; // ATC , Jigucd
          if Jigucd <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM JIGUMST WHERE Jigucd = :Jigucd';
            InsertQuery.ParamByName('Jigucd').AsString := Jigucd;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              UpdateErrorColumn(Row, 'ATC is Invalid');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
            end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'ATC SQL is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;
        end;
        // Remark
        Jisekibikou := StringGridCSV.Cells[Remark, Row];
        // Ensure that Jisekibikou does not exceed 40 characters
        if Length(Jisekibikou) > 40 then
          Jisekibikou := Copy(Jisekibikou, 1, 40);

        if Pos('"', Jisekibikou) > 0 then
        begin
              UpdateErrorColumn(Row, 'Remark error " character ');
              UpdateResultColumn(Row, 'NG');
              num := num + 1;
        end;

        // Prepare data for insertion
        Tourokuymd := Now;

        // MANAGE Ymde DateEnd Just Date no time
        try
          // Ymde date end
          if StringGridCSV.Cells[MachDate, Row] = '' then
          begin
            if TryStrToFloat(StringGridCSV.Cells[Start, Row], StartTime) and
              TryStrToFloat(StringGridCSV.Cells[Finish, Row], EndTime) then
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
                  Ymde := 'Invalid Date';
                  // Handle invalid date in cell [2, Row]
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
          // MANAGE Ymde TIMEEND
          time := GetMaxTime(StringGridCSV.Cells[Finish, Row],
            StringGridCSV.Cells[MachFinish, Row]);
          // MANAGE YMDE DATETIME
          FormattedDateEnd := FormatDateTimeStr(Ymde, time); // lasted ymde
          if StringGridCSV.Cells[MachMin, Row] <> '' then
            begin
                CalculateMinutesDifference(FormattedDateTime, FormattedDateEnd, TotalMinutes);
                Bikou := Inttostr(TotalMinutes);
                if JhValue = StrtoInt(Bikou) then
                begin
                    Bikou := '0';
                end;
            end

        except
          on E: Exception do
          begin
            // Handle any exceptions that occur during date processing
            UpdateResultColumn(Row, 'NG');
            UpdateErrorColumn(Row, 'Ymde is missing ');
            // Update the error column with the error message
            num := num + 1;
          end;
        end;
        // Cost Unit Price
        // Get yujintanka value from tantomst
        InsertQuery.SQL.Text :=
          'SELECT tanka1 FROM tantomst WHERE tantocd = :tantocd';
        InsertQuery.ParamByName('tantocd').AsString := Tantocd;
        InsertQuery.Open;
        YujintankaValue := InsertQuery.FieldByName('tanka1').AsFloat;
        InsertQuery.Close;
        // Get Kikaitanka value from the kikaimst table using kikaicd
        InsertQuery.SQL.Text :=
          'SELECT KIKAITANKA FROM kikaimst WHERE kikaicd = :MachValue';
        InsertQuery.ParamByName('MachValue').AsString := MachValue;
        InsertQuery.Open;
        KikaitankaValue := InsertQuery.FieldByName('KIKAITANKA').AsFloat;
        InsertQuery.Close;
        // Get koteitanka value from the kouteigmst table using Gkoteicd
        InsertQuery.SQL.Text :=
          'SELECT GTANKA FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
        InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
        InsertQuery.Open;
        KoteitankaValue := InsertQuery.FieldByName('GTANKA').AsFloat;
        InsertQuery.Close;
        // Retrieve GHIMOKUCD from KOUTEIGMST
        InsertQuery.SQL.Text :=
          'SELECT GHIMOKUCD FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
        InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
        InsertQuery.Open;
        GHIMOKUCDValue := InsertQuery.FieldByName('GHIMOKUCD').AsInteger;
        InsertQuery.Close;
        // Calculate YujinkinValue, MujinkinValue, and KinsumValue

        YujinkinValue := MinMan * YujintankaValue / 60;
        YujinkinValue := RoundDownTo(YujinkinValue, 2);
        MujinkinValue := MinMach * KikaitankaValue / 60;
        MujinkinValue := RoundDownTo(MujinkinValue, 2);
        KinsumValue := YujinkinValue + MujinkinValue;
        KinsumValue := RoundDownTo(KinsumValue, 2);

        // Format KinsumValue to 2 decimal places
        FormattedKinsumValue := FormatFloat('0.00', KinsumValue);
        KinsumValue := StrToFloat(FormattedKinsumValue);

        // GET PRIMARY KEY
        // Get the maximum JDSEQNO from the JISEKIDATA table
        InsertQuery.SQL.Text :=
          'SELECT MAX(JDSEQNO) AS MaxJDSEQNO FROM JISEKIDATA';
        InsertQuery.Open;
        MaxJDSEQNO := InsertQuery.FieldByName('MaxJDSEQNO').AsInteger;
        InsertQuery.Close;
        // Increment the maximum JDSEQNO by 1 to get the new JDSEQNO
        NewJDSEQNO := MaxJDSEQNO + 1;


      except
        on E: Exception do
        begin
          // Log the error and update the "Result" column with "NG"
          UpdateResultColumn(Row, 'NG');
          UpdateErrorColumn(Row, 'General : ' + E.Message);
          // Update the error column with the error message
          num := num + 1;
        end;
      end;
      ErrorMessageText := '';
      if num = 0 then
      begin
        UpdateResultColumn(Row, 'OK');
      end;
    end;
  finally
    for i := 1 to StringGridCSV.RowCount - 1 do
          begin
            if StringGridCSV.Cells[0, i] = 'OK' then
              Inc(OKCount)
            else if StringGridCSV.Cells[0, i] = 'NG' then
              Inc(NGCount);
          end;

          // Display the counts in the labels
    LabelOK.Caption := 'OK: ' + IntToStr(OKCount) +' NG: ' + IntToStr(NGCount);
    UniConnection.Rollback;
    InsertQuery.Free;
    ProgressBar1.Position := 0; // Reset the progress bar
  end;
end;

procedure TForm1.ImportDataToDatabase;
var
  OKCount, NGCount, i: Integer;
  IniFile: TIniFile;
  IniFileName: string;
  CD2Value: string;
  Row, Col: Integer;
  Value1, Value2, JhValue, Sagyoh, Kikaikadoh, MinMan, MinMach: Integer;
  InsertQuery: TUniQuery;
  SQL, SeizonoValue, BucdValue, MachValue,MachMasterValue, BunmValue,Mold_Code,Model_info,Lamp_Name,Part_Name,Part_Master: string;
  Gkoteicd, Kikaicd, Jigucd, Tantocd, Ymds, Ymde, Bikou, Jisekibikou: string;
  MaxTime, FormattedDateTime, FormattedDateEnd, time: string;
  MaxJDSEQNO, NewJDSEQNO, GHIMOKUCDValue, num: Integer;
  Tourokuymd: TDateTime;
  YujintankaValue, KikaitankaValue, KoteitankaValue, YujinkinValue,
    MujinkinValue, KinsumValue: Double;
  CompName, MACAddr, WinUserName, ExeName, ExeVersion: string;
  Buffer: array [0 .. MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: DWORD;
  StartTime, EndTime, TimeDifference: Double;
  ResultDate: TDateTime;
  timeS, timeE, Shift,tantonm,tantoname: string;
  DateValue, DateMachineValue: TDateTime;
  DateStr,Import_Result,MinManStr,MinMachStr: string;
  TimeStr, TimeFinish, TimeStrMach, TimeFinishMach, DateMach: string;
  Code_A,Code_B,CostProcess_Name,Code_C,File_Name,Result_Detail,textError :string;
  FormattedKinsumValue: String;
  FormatSettings: TFormatSettings;
begin
  // Load the database connection parameters
  LoadConnectionParameters;
  SetIndex;
  OKCount := 0;
  NGCount := 0;

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
        textError := '';
        if (StringGridCSV.Cells[Shift_n, Row] = '') and
           (StringGridCSV.Cells[Date, Row] = '') and
           (StringGridCSV.Cells[WorkerName, Row] = '') and
           (StringGridCSV.Cells[EmployeeCode, Row] = '') and
           (StringGridCSV.Cells[CodeD, Row] = '') and
           (StringGridCSV.Cells[CostProcessName, Row] = '') and
           (StringGridCSV.Cells[MoldCode, Row] = '') and
           (StringGridCSV.Cells[Model, Row] = '') and
           (StringGridCSV.Cells[LampName, Row] = '') and
           (StringGridCSV.Cells[PartName, Row] = '') and
           (StringGridCSV.Cells[ModifyJobNo, Row] = '') and
           (StringGridCSV.Cells[PartCode, Row] = '') and
           (StringGridCSV.Cells[PartMaster, Row] = '') and
           (StringGridCSV.Cells[Start, Row] = '') and
           (StringGridCSV.Cells[Finish, Row] = '') and
           (StringGridCSV.Cells[Min, Row] = '') and
           (StringGridCSV.Cells[MCCode, Row] = '') and
           (StringGridCSV.Cells[Machmaster, Row] = '') and
           (StringGridCSV.Cells[MachStart, Row] = '') and
           (StringGridCSV.Cells[MachDate, Row] = '') and
           (StringGridCSV.Cells[MachFinish, Row] = '') and
           (StringGridCSV.Cells[MachMin, Row] = '') and
           (StringGridCSV.Cells[ATC, Row] = '') and
           (StringGridCSV.Cells[Remark, Row] = '') and
           (StringGridCSV.Cells[CodeA, Row] = '') and
           (StringGridCSV.Cells[CodeC, Row] = '') and
           (StringGridCSV.Cells[CodeB, Row] = '') then
        begin
          continue;
          // Perform actions here if all cells are empty
        end
        else
         begin

        end;

        TotalMinutes := 0;
        Import_Result :='NG' ;
        num := 0;
        // Read ini file
        IniFileName := ExtractFilePath(Application.ExeName) +
          '/Setup/DRLOGIN.ini';
        if not FileExists(IniFileName) then
        begin
          WriteLog('Error Row ' + IntToStr(Row) + ' : INI file not found : ' +
            IniFileName);

          textError := textError +','+'INI file not found: ' + IniFileName;
          num := num + 1; // Skip to the next iteration of the loop
        end;
        IniFile := TIniFile.Create(IniFileName);
        try
          CD2Value := IniFile.ReadString('TLogOnForm', 'CD2', '');
          if CD2Value = '' then
          begin
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : CD2 value not found or INI file not read correctly.');
            
            textError :='CD2 value not found or INI file not read correctly.';
            num := num + 1; // Skip to the next iteration of the loop
          end;
        finally
          IniFile.Free;
        end;
        // Get ComputerName,MacAddress,WindowsUsername,ExecutableName,Executable Version
        // Get Computer Name
        Size := MAX_COMPUTERNAME_LENGTH + 1;
        if not GetComputerName(Buffer, Size) then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' : Failed to get computer name.');

          textError := textError +','+'Failed to get computer name.';
          num := num + 1; // Skip to the next iteration of the loop
        end;
        CompName := Buffer;
        // Get MAC Address
        MACAddr := GetMACAddress;
        if MACAddr = '' then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' : Failed to get MAC address.');
          textError := textError +','+'Failed to get MAC address.';

          num := num + 1; // Skip to the next iteration of the loop
        end;
        // Get Windows Username
        WinUserName := GetWindowsUserName;
        if WinUserName = '' then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' : Failed to get Windows username.');
          textError := textError +','+' Failed to get Windows username.';

          num := num + 1; // Skip to the next iteration of the loop
        end;
        // Get Executable Name
        ExeName := ExtractFileName(Application.ExeName);
        if ExeName = '' then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' : Failed to get executable name.');
          textError := textError +','+'Failed to get executable name.';

          num := num + 1; // Skip to the next iteration of the loop
        end;
        // Get Executable Version
        ExeVersion := GetFileVersion(Application.ExeName);
        if ExeVersion = '' then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' : Failed to get executable version.');

          textError := textError +','+'Failed to get executable version.';
          num := num + 1; // Skip to the next iteration of the loop
        end;

        // prepare and validation Data
        // WorkerCD,Job,CostProcess,ymds not null
        SeizonoValue := StringGridCSV.Cells[ModifyJobNo, Row];
        // seizo,modify job no.
        Gkoteicd := StringGridCSV.Cells[CodeD, Row];
        Code_A :=  StringGridCSV.Cells[CodeA, Row];
        Code_C :=  StringGridCSV.Cells[CodeC, Row];
        CostProcess_Name :=  StringGridCSV.Cells[CostProcessName, Row];
        // CostProcess CD,CodeD,gkoteicd
        Tantocd := StringGridCSV.Cells[EmployeeCode, Row];
        tantoname :=  StringGridCSV.Cells[WorkerName, Row];
        // Employee CD,tantocd
        Ymds := StringGridCSV.Cells[Date, Row]; // ymds , date start
        BucdValue := StringGridCSV.Cells[PartCode, Row]; // partcd
        Part_Name := StringGridCSV.Cells[PartName, Row]; // partname
        Part_Master := StringGridCSV.Cells[PartMaster, Row];        //partmaster
        Mold_Code := StringGridCSV.Cells[MoldCode, Row]; //MoldCode
        Model_info := StringGridCSV.Cells[Model, Row]; //Model_info
        Lamp_Name := StringGridCSV.Cells[LampName, Row]; //LampName
        MinManStr := StringGridCSV.Cells[Min, Row];       //min man
        MinMachStr := StringGridCSV.Cells[MachMin, Row];  //min mach


        // Check Date format DD/MM/YYYY
        DateStr := StringGridCSV.Cells[Date, Row]; // 'Date' should be replaced with the actual index of your date column

        // Set up the format settings to match the desired date format
        FormatSettings := TFormatSettings.Create;
        FormatSettings.DateSeparator := '/';
        FormatSettings.ShortDateFormat := 'dd/mm/yyyy';

        // Try to convert the string to a TDateTime using the specified format
        if (DateStr = '') or not TryStrToDate(DateStr, DateValue, FormatSettings) then
        begin
          WriteLog('Error Row ' + IntToStr(Row) +
            ' :  Invalid or Missing Date Format');

           textError := textError +','+'Invalid or Missing Date Format.';
          // You can also log the error, update a status column, etc.
          num := num + 1; // Skip to the next iteration of the loop
        end
        else if not IsValidDate(YearOf(DateValue), MonthOf(DateValue), DayOf(DateValue)) then
        begin
          WriteLog('Error Row ' + IntToStr(Row) + ' :  Invalid Date Value');

          textError := textError +','+'Invalid Date Value.';
          num := num + 1;
        end;

        try
          // WorkerCD
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM tantomst WHERE tantocd = :tantocd';
          InsertQuery.ParamByName('tantocd').AsString := Tantocd;
          InsertQuery.Open;
          if (Tantocd = '') or (InsertQuery.FieldByName('Count').AsInteger
            <= 0) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) +
              ' :  Employee Code is Invalid');

            textError := textError +','+('Employee Code is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : WorkerCD SQL is Invalid');

            textError := textError +','+('WorkerCD SQL is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;

        try
          // Cost process CD
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM kouteigmst WHERE Gkoteicd = :Gkoteicd';
          InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
          InsertQuery.Open;
          if (Gkoteicd = '') or (InsertQuery.FieldByName('Count').AsInteger
            <= 0) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : Code D is Invalid');

            textError := textError +','+('Code D is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : Cost process CD SQL is Invalid');

            textError := textError +','+('Cost process CD SQL is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;

        try
          // Mfg. No.
          InsertQuery.SQL.Text :=
            'SELECT COUNT(*) AS Count FROM SEIZOMST WHERE Seizono = :SeizonoValue';
          InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
          InsertQuery.Open;
          if (SeizonoValue = '') or
            (InsertQuery.FieldByName('Count').AsInteger <= 0) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : ModifyJobNo is Invalid');

            textError := textError +','+('ModifyJobNo is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;

        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : ModifyJobNo SQL is Invalid');

            textError := textError +','+( 'ModifyJobNo SQL is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;

        try
          // Part Code
          if BucdValue <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM BUHINMST WHERE bucd = :BucdValue';
            InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              WriteLog('Error Row ' + IntToStr(Row) + ' : PartCode is Invalid');
              textError := textError + ',' + ('PartCode is Invalid.');
              num := num + 1; // Skip to the next iteration of the loop
            end
            else
            begin
              InsertQuery.SQL.Text := 'SELECT bunm FROM BUHINMST WHERE bucd = :BucdValue';
              InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
              InsertQuery.Open;
              if not InsertQuery.EOF then
              begin
                BunmValue := InsertQuery.FieldByName('bunm').AsString;
                // Use BunmValue as needed in your code
              end;
            end;
            InsertQuery.Close;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : PARTCode SQL is Invalid');

            textError := textError +','+('PartCode SQL is Invalid.');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;

        TimeStr := StringGridCSV.Cells[Start, Row];
        TimeFinish := StringGridCSV.Cells[Finish, Row];
        TimeStrMach := StringGridCSV.Cells[MachStart, Row];
        TimeFinishMach := StringGridCSV.Cells[MachFinish, Row];
        DateMach := StringGridCSV.Cells[MachDate, Row];
        if not TryStrToInt(StringGridCSV.Cells[Min, Row], MinMan) then
          MinMan := 0;
        if not TryStrToInt(StringGridCSV.Cells[MachMin, Row], MinMach) then
          MinMach := 0;
        JhValue := MinMan + MinMach;
        // Calculate additional fields
        Sagyoh := MinMan + 0 + 0;
        Kikaikadoh := MinMan + MinMach + 0 + 0;

        if MinManStr <> '' then
        begin
            if (TimeStr = '') or (TimeFinish = '') or (DateStr = '') then
            begin
              WriteLog('Error Row ' + IntToStr(Row) +' : Time is Valid');

              textError := textError +','+('Time is Valid');
              num := num + 1; // Skip to the next iteration of the loop
            end;
            if (MinMachStr <> '') or (TimeStrMach <> '') or (TimeFinishMach <> '') or (DateMach <> '') then
            begin
              WriteLog('Error Row ' + IntToStr(Row) +' : Contain both Man and Unman data');

              textError := textError +','+('Contain both Man and Unman data');
              num := num + 1; // Skip to the next iteration of the loop
            end;
        end
        else if MinMachStr <> '' then
        begin
           if (TimeStrMach = '') or (TimeFinishMach = '') or (DateMach = '') then
            begin
              WriteLog('Error Row ' + IntToStr(Row) +' : Time is Valid');

              textError := textError +','+('Time is Valid');
              num := num + 1; // Skip to the next iteration of the loop
            end;

            if (MinManStr <> '') or (TimeStr <> '') or (TimeFinish <> '') then
            begin
              WriteLog('Error Row ' + IntToStr(Row) +' : Contain both Man and Unman data');

              textError := textError +','+('Contain both Man and Unman data');
              num := num + 1; // Skip to the next iteration of the loop
            end;
        end;

        // Machine Unman
       if (TimeStrMach <> '') and (TimeFinishMach <> '') and (DateMach <> '') AND (TimeStr = '') and (TimeFinish = '')  then
        begin
          if not(IsValidTimeFormat(TimeStrMach)) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : TimeMachStart is Invalid');

            textError := textError +','+('TimeMachStart is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          if not IsValidTimeFormat(TimeFinishMach) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : TimeMachFinish is Invalid');

            textError := textError +','+('TimeMachFinish is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          if (DateMach = '') and not TryStrToDate(StringGridCSV.Cells[MachDate,
            Row], DateMachineValue) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : MachDate is null');

            textError := textError +','+('MachDate is null');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          if MinMach = 0 then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : MinMach is null');

            textError := textError +','+( 'MinMach is null');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          if (DateMach = DateStr) and ( strtofloat(TimeStrMach) > strtofloat(TimeFinishMach) ) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : TimeStart&TimeEnd is Invalid');

            textError := textError +','+('TimeStart&TimeEnd is Invalid');
            num := num + 1;
          end;

          if (DateMach < DateStr) then
          begin
            UpdateErrorColumn(Row, 'TimeStart&TimeEnd is Invalid');
            textError := textError +','+( 'TimeStart&TimeEnd is Invalid');
            UpdateResultColumn(Row, 'NG');
            num := num + 1;
          end;

          Bikou := StringGridCSV.Cells[MachMin, Row];
          MaxTime := GetMaxTime(StringGridCSV.Cells[Start, Row],
          StringGridCSV.Cells[MachStart, Row]);
          FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime); // lasted ymds
        end
        // Worker Manned
       else if (TimeStr <> '') and (TimeFinish <> '') AND (TimeStrMach = '') and (TimeFinishMach = '') and (DateMach = '')   then
        begin
          if not IsValidTimeFormat(TimeStr) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : TimeStart is Invalid');

            textError := textError +','+('TimeStart is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          if not IsValidTimeFormat(TimeFinish) then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : TimeFinish is Invalid');

            textError := textError +','+( 'TimeFinish is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;
          if MinMan = 0 then
          begin
            WriteLog('Error Row ' + IntToStr(Row) + ' : MinMan is null');

            textError := textError +','+( 'MinMan is null');
            num := num + 1; // Skip to the next iteration of the loop
          end;

          timeS := StringGridCSV.Cells[Start, Row];
          timeE := StringGridCSV.Cells[Finish, Row];
          Shift := StringGridCSV.Cells[Shift_n, Row];
          if num = 0 then
          begin
            Bikou := CalculateWorkingTime(timeS, timeE, Shift); // minman
            TotalMinutes :=    StrtoInt(Bikou);
            MaxTime := GetMaxTime(StringGridCSV.Cells[Start, Row],
            StringGridCSV.Cells[MachStart, Row]);
            FormattedDateTime := FormatDateTimeStr(Ymds, MaxTime); // lasted ymds
          end;

        end
        else
        begin
          WriteLog('Error Row ' + IntToStr(Row) + ' : Time is Invalid');

          textError := textError +','+( 'Time is Invalid');
          num := num + 1; // Skip to the next iteration of the loop
        end;

        try
          // calculate for check MinMan
          if num = 0 then
          begin
            if JhValue = StrToInt(Bikou) then
            begin
              Bikou := '0'; // cal min for sure
            end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            UpdateErrorColumn(Row, 'Check JhMin is invalid');
            textError := textError +','+( 'Check JhMin is invalid');
            UpdateResultColumn(Row, 'NG');
            Continue; // Skip to the next iteration of the loop
          end;
        end;

        try
          // Machine CD
          MachValue := StringGridCSV.Cells[MCCode, Row]; // MachValue
          MachMasterValue := StringGridCSV.Cells[MachMaster, Row];
          if MachValue <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM kikaimst WHERE kikaicd = :MachValue';
            InsertQuery.ParamByName('MachValue').AsString := MachValue;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              WriteLog('Error Row ' + IntToStr(Row) + ' : Machine is Invalid');

              textError := textError +','+( 'Machine is Invalid');
              num := num + 1; // Skip to the next iteration of the loop
            end;
          end
          else  //MachValue = ''
          begin
               if (MinMachStr <> '') or (TimeStrMach <> '') or (TimeFinishMach <> '') or (DateMach <> '') then
                   begin
                        WriteLog('Error Row ' + IntToStr(Row) + ' : Machine is Null');

                        textError := textError +','+( 'Machine is Null');
                        num := num + 1; // Skip to the next iteration of the loop
                   end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) +
              ' : Machine SQL is Invalid');

            textError := textError +','+( 'Machine SQL is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;

        try
          // Jigucd ATC
          Jigucd := StringGridCSV.Cells[ATC, Row]; // ATC , Jigucd
          if Jigucd <> '' then
          begin
            InsertQuery.SQL.Text :=
              'SELECT COUNT(*) AS Count FROM JIGUMST WHERE Jigucd = :Jigucd';
            InsertQuery.ParamByName('Jigucd').AsString := Jigucd;
            InsertQuery.Open;
            if (InsertQuery.FieldByName('Count').AsInteger <= 0) then
            begin
              WriteLog('Error Row ' + IntToStr(Row) + ' : ATC is Invalid');

              textError := textError +','+( 'ATC is Invalidd');
              num := num + 1; // Skip to the next iteration of the loop
            end;
          end;
        except
          on E: Exception do
          begin
            // Handle exceptions such as connectivity issues, SQL syntax errors, etc.
            WriteLog('Error Row ' + IntToStr(Row) + ' : ATC SQL is Invalid');
            textError := textError +','+( 'ATC SQL is Invalid');
            num := num + 1; // Skip to the next iteration of the loop
          end;
        end;
        // Remark
        Jisekibikou := StringGridCSV.Cells[Remark, Row];
        // Ensure that Jisekibikou does not exceed 40 characters
        if Length(Jisekibikou) > 40 then
          Jisekibikou := Copy(Jisekibikou, 1, 40);

        if Pos('"', Jisekibikou) > 0 then
        begin
            UpdateErrorColumn(Row, 'Remark error " character ');
            textError := textError +','+( 'Remark error " character ');
            UpdateResultColumn(Row, 'NG');
            Continue; // Skip to the next iteration of the loop
        end;

        // Prepare data for insertion
        Tourokuymd := Now;

        // MANAGE Ymde DateEnd Just Date no time
        try
          // Ymde date end
          if StringGridCSV.Cells[MachDate, Row] = '' then
          begin
            if TryStrToFloat(StringGridCSV.Cells[Start, Row], StartTime) and
              TryStrToFloat(StringGridCSV.Cells[Finish, Row], EndTime) then
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
                  Ymde := 'Invalid Date';
                  // Handle invalid date in cell [2, Row]
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
          // MANAGE Ymde TIMEEND
          time := GetMaxTime(StringGridCSV.Cells[Finish, Row],
            StringGridCSV.Cells[MachFinish, Row]);
          // MANAGE YMDE DATETIME
          FormattedDateEnd := FormatDateTimeStr(Ymde, time); // lasted ymde
          if StringGridCSV.Cells[MachMin, Row] <> '' then
            begin
                CalculateMinutesDifference(FormattedDateTime, FormattedDateEnd, TotalMinutes);
                Bikou := Inttostr(TotalMinutes);
                if JhValue = StrtoInt(Bikou) then
                begin
                    Bikou := '0';
                end;
            end

        except
          on E: Exception do
          begin
            // Handle any exceptions that occur during date processing
            WriteLog('Error Row ' + IntToStr(Row) + ' : Ymde is missing');
            textError := textError +','+( 'Ymde is missing ');
            num := num + 1; // Skip to the next iteration of the loop
          end;

        end;


        //sum error and send message
        if num > 0 then
        begin
          WriteLog('Result : NG = ' + IntToStr(num) + ' in Row ' + IntToStr(Row));
          LogRowToCSV(Row, textError);
          LogErrorRowToCSV(Row, textError);
        end
        else
        begin
              // Cost Unit Price
              // Get yujintanka value from tantomst
              InsertQuery.SQL.Text :=
                'SELECT tanka1,tantonm FROM tantomst WHERE tantocd = :tantocd';
              InsertQuery.ParamByName('tantocd').AsString := Tantocd;
              InsertQuery.Open;
              YujintankaValue := InsertQuery.FieldByName('tanka1').AsFloat;
              tantonm := InsertQuery.FieldByName('tantonm').AsString;
              InsertQuery.Close;
              // Get Kikaitanka value from the kikaimst table using kikaicd
              InsertQuery.SQL.Text :=
                'SELECT KIKAITANKA FROM kikaimst WHERE kikaicd = :MachValue';
              InsertQuery.ParamByName('MachValue').AsString := MachValue;
              InsertQuery.Open;
              KikaitankaValue := InsertQuery.FieldByName('KIKAITANKA').AsFloat;
              InsertQuery.Close;
              // Get koteitanka value from the kouteigmst table using Gkoteicd
              InsertQuery.SQL.Text :=
                'SELECT GTANKA FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
              InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
              InsertQuery.Open;
              KoteitankaValue := InsertQuery.FieldByName('GTANKA').AsFloat;
              InsertQuery.Close;
              // Retrieve GHIMOKUCD from KOUTEIGMST
              InsertQuery.SQL.Text :=
                'SELECT GHIMOKUCD FROM KOUTEIGMST WHERE Gkoteicd = :Gkoteicd';
              InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
              InsertQuery.Open;
              GHIMOKUCDValue := InsertQuery.FieldByName('GHIMOKUCD').AsInteger;
              InsertQuery.Close;
              // Calculate YujinkinValue, MujinkinValue, and KinsumValue
              YujinkinValue := MinMan * YujintankaValue / 60;
              YujinkinValue := RoundDownTo(YujinkinValue, 2);
              MujinkinValue := MinMach * KikaitankaValue / 60;
              MujinkinValue := RoundDownTo(MujinkinValue, 2);
              KinsumValue := YujinkinValue + MujinkinValue;
              KinsumValue := RoundDownTo(KinsumValue, 2);

              // Format KinsumValue to 2 decimal places
              FormattedKinsumValue := FormatFloat('0.00', KinsumValue);
              KinsumValue := StrToFloat(FormattedKinsumValue);

              // GET PRIMARY KEY
              // Get the maximum JDSEQNO from the JISEKIDATA table
              InsertQuery.SQL.Text :=
                'SELECT MAX(JDSEQNO) AS MaxJDSEQNO FROM JISEKIDATA';
              InsertQuery.Open;
              MaxJDSEQNO := InsertQuery.FieldByName('MaxJDSEQNO').AsInteger;
              InsertQuery.Close;
              // Increment the maximum JDSEQNO by 1 to get the new JDSEQNO
              NewJDSEQNO := MaxJDSEQNO + 1;

              // Update SEQNO in HATUBAN based on JISEKIDATA
              InsertQuery.SQL.Text :=
                'UPDATE HATUBAN ' +
                'SET SEQNO = SEQNO + 1 ' +
                'WHERE ID = ''JISEKIDATA''';
              InsertQuery.ExecSQL;

              UniConnection.Commit;

              // Check Error Insertion
              try
                // Construct and execute the SQL statement
                SQL := 'INSERT INTO JISEKIDATA (JDSEQNO, seizono, bunm, bucd, gkoteicd, kikaicd, jigucd, tantocd, ymds, KMSEQNO, jh, '
                  + 'jmaedanh, jatodanh, jkbn, jyujinh, jmujinh, yujintanka, kikaitanka, koteitanka, GHIMOKUCD, yujinkin, '
                  + 'mujinkin, kinsum, bikou, tourokuymd, sagyoh, kikaikadoh, inptantocd, inpymd, jisekibikou, inppcname, '
                  + 'inpmacaddress, inpusername, inpexename, inpversion,ymde) ' +

                  'VALUES (:NewJDSEQNO, :SeizonoValue, :BunmValue, :BucdValue, :Gkoteicd, :Kikaicd, :Jigucd, :Tantocd, '
                  + 'TO_DATE(:FormattedDateTime, ''YYYY-MM-DD HH24:MI:SS''), 1, :JhValue, 0, 0, 4, :MinMan, :MinMach, '
                  + ':YujintankaValue, :KikaitankaValue, :KoteitankaValue, :GHIMOKUCDValue, :YujinkinValue, :MujinkinValue, '
                  + ':KinsumValue, :Bikou, :Tourokuymd, :Sagyoh, :Kikaikadoh, :InptantocdValue, :Inpymd, :Jisekibikou, '
                  + ':Inppcname, :Inpmacaddress, :Inpusername, :Inpexename, :Inpversion,TO_DATE(:FormattedDateEnd, ''YYYY-MM-DD HH24:MI:SS''))';

                InsertQuery.SQL.Text := SQL;
                InsertQuery.ParamByName('NewJDSEQNO').AsInteger := NewJDSEQNO;
                InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
                InsertQuery.ParamByName('BunmValue').AsString := BunmValue;
                InsertQuery.ParamByName('BucdValue').AsString := BucdValue;
                InsertQuery.ParamByName('Gkoteicd').AsString := Gkoteicd;
                InsertQuery.ParamByName('Kikaicd').AsString := MachValue;
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

                Import_Result :='OK' ;
                WriteLog('Result : Suscess in row ' + IntToStr(Row) + ' ID : ' +
                IntToStr(NewJDSEQNO));
                UpdateErrorColumn(Row, 'ID' + IntToStr(NewJDSEQNO));
                UpdateResultColumn(Row, 'Imported');
                LogRowToCSV(Row, 'Suscess ID : ' + IntToStr(NewJDSEQNO));

              except
                on E: Exception do
                begin

                  // Log the error and update the "Result" column with "NG"
                  WriteLog('Error Insertion In Row ' + IntToStr(Row) + ' : ' +
                    E.Message);
                  LogRowToCSV(Row, 'Error Insertion : ' + E.Message);
                  UpdateErrorColumn(Row, 'Error: Insertion SQL '+ E.Message);
                  UpdateResultColumn(Row, 'NG');
                  Continue;
                end;
              end;
        end;

        try  //insert into history actual database_tkoito
                File_Name := StringGridCSV.Cells[28, Row];
                Result_Detail := StringGridCSV.Cells[29, Row];
        //insert into history actual database_tkoito
        SQL := 'INSERT INTO HISTORY_IMPORT_ACTUAL (Transaction_Date ,Import_Result ,Start_Date ,Worker_Name ,Employee_Code , '
            + 'Code_A ,Code_B ,Code_D ,Cost_Process_Name ,Modify_Job_No ,Mold_Code ,Model_info ,Lamp_Name ,Part_Name ,Part_Code ,Part_master , '
            + 'Code_C ,Man_Start ,Man_Finish ,Man_Min ,MC_Code ,Mach_master ,Unman_Start ,Unman_Finish_Date ,Unman_Finish ,Unman_Min ,ATC , '
            + 'Remark ,Memo ,Man_Start_Datetime ,Man_End_Datetime ,Man_Min_Cal ,Mach_Start_Datetime ,Mach_End_Datetime ,Result_Detail ,File_Name) ' +

            'VALUES (:Transaction_Date ,:Import_Result ,:Start_Date ,:tantonm,:tantocd ,:Code_A ,:Code_B ,:Code_D ,:CostProcess_Name , '
            + ':SeizonoValue ,:Mold_Code ,:Model_info ,:Lamp_Name ,:Part_Name ,:Part_Code ,:Part_Master ,:Code_C ,:TimeStr ,:TimeFinish ,:MinMan , '
            + ':MachValue ,:MachMasterValue ,:Unman_Start ,:Unman_Finish_Date ,:Unman_Finish ,:Unman_Min ,:ATC ,:Remark ,:Memo ,:Man_Start_Datetime , '
            + ':Man_End_Datetime,:Man_Min_Cal,:Mach_Start_Datetime,'
            + ':Mach_End_Datetime , :Result_Detail ,:File_Name)';

          InsertQuery.SQL.Text := SQL;
          InsertQuery.ParamByName('Transaction_Date').AsDateTime := Tourokuymd;
          InsertQuery.ParamByName('Import_Result').AsString := Import_Result;
          InsertQuery.ParamByName('Start_Date').AsString := DateStr;
          InsertQuery.ParamByName('tantonm').AsString := tantoname;
          InsertQuery.ParamByName('tantocd').AsString := Tantocd;
          InsertQuery.ParamByName('Code_A').AsString := Code_A;
          InsertQuery.ParamByName('Code_B').AsString := Shift;
          InsertQuery.ParamByName('Code_D').AsString := Gkoteicd;
          InsertQuery.ParamByName('CostProcess_Name').AsString := CostProcess_Name;
          InsertQuery.ParamByName('SeizonoValue').AsString := SeizonoValue;
          InsertQuery.ParamByName('Mold_Code').AsString := Mold_Code;
          InsertQuery.ParamByName('Model_info').AsString := Model_info;
          InsertQuery.ParamByName('Lamp_Name').AsString := Lamp_Name;
          InsertQuery.ParamByName('Part_Name').AsString := BunmValue;
          InsertQuery.ParamByName('Part_Code').AsString := BucdValue;
          InsertQuery.ParamByName('Part_Master').AsString := Part_Master;
          InsertQuery.ParamByName('Code_C').AsString := Code_C;
          InsertQuery.ParamByName('TimeStr').AsString := TimeStr;      //STR but in database is 00.00
          InsertQuery.ParamByName('TimeFinish').AsString := TimeFinish;  //STR but in database is 00.00
          InsertQuery.ParamByName('MinMan').AsString := MinManStr;
          InsertQuery.ParamByName('MachValue').AsString := MachValue;
          InsertQuery.ParamByName('MachMasterValue').AsString := MachMasterValue;
          InsertQuery.ParamByName('Unman_Start').AsString := TimeStrMach ;
          InsertQuery.ParamByName('Unman_Finish_Date').AsString := DateMach ;
          InsertQuery.ParamByName('Unman_Finish').AsString := TimeFinishMach ;
          InsertQuery.ParamByName('Unman_Min').AsString := MinMachStr;
          InsertQuery.ParamByName('ATC').AsString := Jigucd;
          InsertQuery.ParamByName('Remark').AsString := Jisekibikou;
          InsertQuery.ParamByName('Memo').AsString := bikou;
          if  MinMan <> 0 then
          begin
            InsertQuery.ParamByName('Man_Start_Datetime').AsString := FormattedDateTime;
            InsertQuery.ParamByName('Man_End_Datetime').AsString := FormattedDateEnd;

            InsertQuery.ParamByName('Mach_Start_Datetime').AsString := '';
            InsertQuery.ParamByName('Mach_End_Datetime').AsString := '';
          end
          else
          begin
            InsertQuery.ParamByName('Man_Start_Datetime').AsString := '';
            InsertQuery.ParamByName('Man_End_Datetime').AsString := '';

            InsertQuery.ParamByName('Mach_Start_Datetime').AsString := FormattedDateTime;
            InsertQuery.ParamByName('Mach_End_Datetime').AsString := FormattedDateEnd;
          end;
          InsertQuery.ParamByName('Man_Min_Cal').AsInteger := TotalMinutes;
          InsertQuery.ParamByName('Result_Detail').AsString :=  Result_Detail;
          InsertQuery.ParamByName('File_Name').AsString :=  File_Name;

          InsertQuery.Execute;

          ProgressBar1.Position := Row;
          WriteLog('Result : LogImport Suscess : ' + IntToStr(NewJDSEQNO));

        except
          on E: Exception do
          begin
            // Log the error and update the "Result" column with "NG"
            WriteLog('Error LogImport In Row ' + IntToStr(Row) + ' : ' + E.Message);
            Continue;
          end;
        end;


      except
        on E: Exception do
        begin
          // Log the error and update the "Result" column with "NG"
          WriteLog('Error General In Row ' + IntToStr(Row) + ' : ' + E.Message);
          LogRowToCSV(Row, 'Error General : ' + E.Message);
          Continue;
        end;
      end;
    end;
    // Commit the transaction
    UniConnection.Commit;
  finally
      for i := 1 to StringGridCSV.RowCount - 1 do
          begin
            if StringGridCSV.Cells[0, i] = 'Imported' then
              Inc(OKCount)
            else if StringGridCSV.Cells[0, i] = 'NG' then
              Inc(NGCount);
          end;

          // Display the counts in the labels
    LabelOK.Caption := 'Imported: ' + IntToStr(OKCount) +' NG: ' + IntToStr(NGCount);
    InsertQuery.Free;
    Managefile;
    ProgressBar1.Position := 0; // Reset the progress bar
  end;
end;

procedure Tform1.CalculateMinutesDifference(const StartTime, EndTime: String; out TotalMinutes: Integer);
var
  StartDT, EndDT: TDateTime;
  FormatSettings: TFormatSettings;
begin
  // Define your date and time format
  FormatSettings := TFormatSettings.Create;
  FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  FormatSettings.LongTimeFormat := 'hh:nn:ss';
  FormatSettings.DateSeparator := '-';
  FormatSettings.TimeSeparator := ':';

  // Convert the start and end time strings to TDateTime using the custom format settings
  StartDT := StrToDateTime(StartTime, FormatSettings);
  EndDT := StrToDateTime(EndTime, FormatSettings);

  // Calculate the total difference in minutes
  TotalMinutes := MinutesBetween(StartDT, EndDT);
end;

procedure TForm1.LogRowToCSV(Row: Integer; ErrorMessage: string);
var
  ErrorLog: TextFile;
  ErrorFileName: string;
  Col: Integer;
  RowValues: string;
  HeaderRow: string;
  ColumnTitles: array of string; // Array to store column titles
  CurrentDate: string;
begin
  if HasResult = '1' then
  Begin
    // Initialize the column titles array
    SetLength(ColumnTitles, StringGridCSV.ColCount+2);
    ColumnTitles[0] := 'Row';
    ColumnTitles[1] := 'Result';
    ColumnTitles[2] := 'Shift';
    ColumnTitles[3] := 'Date';
    ColumnTitles[4] := 'Name';
    ColumnTitles[5] := 'Employee Code';
    ColumnTitles[6] := 'Code A';
    ColumnTitles[7] := 'Code B';
    ColumnTitles[8] := 'Code D';
    ColumnTitles[9] := 'Cost Process Name';
    ColumnTitles[10] := 'Modify Job No.';
    ColumnTitles[11] := 'Mold Code';
    ColumnTitles[12] := 'Model';
    ColumnTitles[13] := 'Lamp Name';
    ColumnTitles[14] := 'Part Name';
    ColumnTitles[15] := 'Part Code';
    ColumnTitles[16] := 'Part master';
    ColumnTitles[17] := 'Code C';
    ColumnTitles[18] := 'Start';
    ColumnTitles[19] := 'Finish';
    ColumnTitles[20] := 'Min';
    ColumnTitles[21] := 'M/C Code';
    ColumnTitles[22] := 'Mach.master';
    ColumnTitles[23] := 'Start';
    ColumnTitles[24] := 'Date';
    ColumnTitles[25] := 'Finish';
    ColumnTitles[26] := 'Min';
    ColumnTitles[27] := 'ATC';
    ColumnTitles[28] := 'Remark';
    ColumnTitles[29] := 'FileName';
    ColumnTitles[30] := 'Result_detail';

    // Add more titles as needed...
    // Initialize the error log file
    CurrentDate := FormatDateTime('YYYYMMDD', Now);
    // Append the date to the file name
    ErrorFileName := ResultPathCSV + '/' + CurrentDate + '_LogImport' + '.csv';

    AssignFile(ErrorLog, ErrorFileName);
    if FileExists(ErrorFileName) then
      Append(ErrorLog)
    else
    begin
      Rewrite(ErrorLog);
      // Construct the header row with column titles from the array
      HeaderRow := '';
      for Col := 0 to High(ColumnTitles) do
      begin
        if Col > 0 then
          HeaderRow := HeaderRow + ',';
        HeaderRow := HeaderRow + ColumnTitles[Col];
      end;
      // Write the header row to the CSV file
      Writeln(ErrorLog, HeaderRow);
    end;

    // Construct the comma-separated string of all values in the row
    RowValues := '';
    for Col := 0 to StringGridCSV.ColCount - 2 do
    begin
      if Col > 0 then
        RowValues := RowValues + ',';
      RowValues := RowValues + StringGridCSV.Cells[Col, Row];
    end;

    RowValues :=  InttoStr(Row) + ',' + RowValues + ErrorMessage;
    Writeln(ErrorLog, RowValues);
    CloseFile(ErrorLog);
  End
  else
  begin
    // Nothing to do!!
  end;

end;

procedure TForm1.LogErrorRowToCSV(Row: Integer; ErrorMessage: string);
var
  ErrorLog: TextFile;
  ErrorFileName: string;
  Col: Integer;
  RowValues: string;
  HeaderRow: string;
  ColumnTitles: array of string; // Array to store column titles
  CurrentDate: string;
begin
  if Error = '1' then
  Begin
    // Initialize the column titles array
    SetLength(ColumnTitles, StringGridCSV.ColCount);
    ColumnTitles[0] := 'Row';
    ColumnTitles[1] := 'Shift';
    ColumnTitles[2] := 'Date';
    ColumnTitles[3] := 'Name';
    ColumnTitles[4] := 'Employee Code';
    ColumnTitles[5] := 'Code A';
    ColumnTitles[6] := 'Code B';
    ColumnTitles[7] := 'Code D';
    ColumnTitles[8] := 'Cost Process Name';
    ColumnTitles[9] := 'Modify Job No.';
    ColumnTitles[10] := 'Mold Code';
    ColumnTitles[11] := 'Model';
    ColumnTitles[12] := 'Lamp Name';
    ColumnTitles[13] := 'Part Name';
    ColumnTitles[14] := 'Part Code';
    ColumnTitles[15] := 'Part master';
    ColumnTitles[16] := 'Code C';
    ColumnTitles[17] := 'Start';
    ColumnTitles[18] := 'Finish';
    ColumnTitles[19] := 'Min';
    ColumnTitles[20] := 'M/C Code';
    ColumnTitles[21] := 'Mach.master';
    ColumnTitles[22] := 'Start';
    ColumnTitles[23] := 'Date';
    ColumnTitles[24] := 'Finish';
    ColumnTitles[25] := 'Min';
    ColumnTitles[26] := 'ATC';
    ColumnTitles[27] := 'Remark';
    ColumnTitles[28] := 'FileName';
    ColumnTitles[29] := 'Result_detail';

    // Add more titles as needed...
    // Initialize the error log file
    CurrentDate := FormatDateTime('YYYYMMDD', Now);
    // Append the date to the file name
    ErrorFileName := PathErrorCSV + '/' + CurrentDate + '_ErrorLog' + '.csv';

    AssignFile(ErrorLog, ErrorFileName);
    if FileExists(ErrorFileName) then
      Append(ErrorLog)
    else
    begin
      Rewrite(ErrorLog);
      // Construct the header row with column titles from the array
      HeaderRow := '';
      for Col := 0 to High(ColumnTitles) do
      begin
        if Col > 0 then
          HeaderRow := HeaderRow + ',';
        HeaderRow := HeaderRow + ColumnTitles[Col];
      end;
      // Write the header row to the CSV file
      Writeln(ErrorLog, HeaderRow);
    end;
    // Construct the comma-separated string of all values in the row
    RowValues := '';
    for Col := 1 to StringGridCSV.ColCount - 2 do
    begin
      if Col > 1 then
        RowValues := RowValues + ',';
        RowValues := RowValues + StringGridCSV.Cells[Col, Row];
    end;
    // Append the error message as the last column
    RowValues := InttoStr(Row) + ',' + RowValues + ErrorMessage;
    // Write the row's values with the error message to the CSV file
    Writeln(ErrorLog, RowValues);
    // Close the error log file
    CloseFile(ErrorLog);
  End
  else
  begin
    // Nothing to do!!
    WriteLog('LogErrorRowToCSV');
  end;

end;

function TForm1.IsMaxTime(CellValue1, CellValue2: string): string;
var
  Time1, Time2: TDateTime;
begin
  Time1 := EncodeTime(StrToInt(Copy(CellValue1, 1, 2)),
    StrToInt(Copy(CellValue1, 4, 2)), 0, 0);
  Time2 := EncodeTime(StrToInt(Copy(CellValue2, 1, 2)),
    StrToInt(Copy(CellValue2, 4, 2)), 0, 0);

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

function TForm1.CalculateWorkingTime(StartTimeStr, EndTimeStr,
  Shift: string): String;
var
  StartTime, EndTime, BreakStart, BreakEnd, WorkingTime, OverlapTime: Double;
  BreakTimes: array [0 .. 9] of Double;
  i: Integer;
  TotalMinutes: Double;
  CrossesMidnight: Boolean;
begin
  try
    // Convert string times to float
    StartTime := StrToInt(Copy(StartTimeStr, 1, 2)) * 60 +
      StrToInt(Copy(StartTimeStr, 4, 2));
    EndTime := StrToInt(Copy(EndTimeStr, 1, 2)) * 60 +
      StrToInt(Copy(EndTimeStr, 4, 2));

    // Check if the time range crosses midnight
    CrossesMidnight := EndTime < StartTime;
    if CrossesMidnight then
    begin
      // Add 24 hours (in minutes) to the end time if it crosses midnight
      EndTime := EndTime + 24 * 60;
    end;


       if Shift = 'N' then // Night shift
    begin
      BreakTimes[0] := 24 * 60;        // 00:00 next day
      BreakTimes[1] := 25 * 60;        // 01:00 next day
      BreakTimes[2] := 29 * 60;        // 05:00 next day
      BreakTimes[3] := 29 * 60 + 20;   // 05:20 next day
      BreakTimes[4] := 33 * 60;        // 09:00 next day
      BreakTimes[5] := 33 * 60 + 20;   // 09:20 next day
      BreakTimes[6] := 37 * 60 + 20;   // 13:20 next day
      BreakTimes[7] := 37 * 60 + 40;   // 13:40 next day
      BreakTimes[8] := 41 * 60 + 40;   // 17:40 next day
      BreakTimes[9] := 42 * 60;        // 18:00 next day
    end
    else if Shift = 'D' then // Day shift
    begin
      BreakTimes[0] := 12 * 60 + 10;   // 12:10
      BreakTimes[1] := 13 * 60 + 10;   // 13:10
      BreakTimes[2] := 17 * 60;        // 17:00
      BreakTimes[3] := 17 * 60 + 30;   // 17:30
      BreakTimes[4] := 21 * 60 + 10;   // 21:10
      BreakTimes[5] := 21 * 60 + 30;   // 21:30
      BreakTimes[6] := 25 * 60 + 30;   // 01:30 next day
      BreakTimes[7] := 25 * 60 + 50;   // 01:50 next day
      BreakTimes[8] := 29 * 60 + 50;   // 05:50 next day
      BreakTimes[9] := 30 * 60 + 10;   // 06:10 next day
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
        OverlapTime := MinFloat(EndTime, BreakEnd) -
          MaxFloat(StartTime, BreakStart);
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
      WriteLog('Error calculating working time : ' + E.Message);
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
    Min := StrToIntDef(Copy(TimeStr, Pos('.', TimeStr) + 1,
      Length(TimeStr)), 0);
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
    Time1Minutes := StrToInt(Copy(Time1, 1, Pos('.', Time1) - 1)) * 60 +
      StrToInt(Copy(Time1, Pos('.', Time1) + 1, Length(Time1)));
  if ValidTime2 then
    Time2Minutes := StrToInt(Copy(Time2, 1, Pos('.', Time2) - 1)) * 60 +
      StrToInt(Copy(Time2, Pos('.', Time2) + 1, Length(Time2)));

  // Compare the total minutes and return the maximum time in 'HH:NN:SS' format
  if ValidTime1 and ValidTime2 then
  begin
    if Time1Minutes > Time2Minutes then
      Result := Format('%2d:%2d:00', [Time1Minutes div 60, Time1Minutes mod 60])
      // Convert minutes back to 'HH:NN:SS' format
    else
      Result := Format('%2d:%2d:00', [Time2Minutes div 60, Time2Minutes mod 60]
        ); // Convert minutes back to 'HH:NN:SS' format
  end
  else if ValidTime1 then
    Result := Format('%2d:%2d:00', [Time1Minutes div 60, Time1Minutes mod 60])
  else if ValidTime2 then
    Result := Format('%2d:%2d:00', [Time2Minutes div 60, Time2Minutes mod 60]);
end;

function TForm1.FormatDateTimeStr(const DateStr, TimeStr: string): string;
var
  DateTime: TDateTime;
  FormattedTime: string;
begin
  // First, convert the date string from 'dd/mm/yyyy' to 'yyyy-mm-dd'
  if TryStrToDate(DateStr, DateTime) then
    Result := FormatDateTime('yyyy-mm-dd', DateTime)
  else
    Result := 'NULL';

  // Format the TimeStr to ensure two digits for hours, minutes, and seconds
  if TryStrToTime(TimeStr, DateTime) then
    FormattedTime := FormatDateTime('hh:nn:ss', DateTime)
  else
    FormattedTime := 'NULL'; // Default to '00:00:00' if the time string is invalid

  // Append the formatted time
  Result := Result + ' ' + FormattedTime;
end;

function TForm1.GetBunmFromBucd(const BucdValue: string): string;
var
  Query: TUniQuery;
begin
  Result := '';
  // Default result is empty string, indicating not found or an error
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
  if StringGridCSV.ColCount <= 29 then // Check if the error column (25) exists
  begin
    StringGridCSV.ColCount := 30; // Ensure there are at least 26 columns
  end;

  // Set the width of all columns except column 25 to 50
  for i := 0 to StringGridCSV.ColCount - 1 do
  begin
    if i <> 30 then
      StringGridCSV.ColWidths[i] := 50;
  end;

  // Update the error message in column 25.
  StringGridCSV.Cells[Status, Row] := StringGridCSV.Cells[Status, Row] + ',' +
    ErrorMessage;

  StringGridCSV.Anchors := [akLeft, akTop, akRight, akBottom];
end;

procedure TForm1.LoadCSVFilesIntoGrid(const FolderPath: string);
var
  Files: TStringDynArray;
  CSVLines: TStringList;
  FilePath, Line, FileName: string;
  Row, Col, MaxCol, FilenameColIndex: Integer;
  CSVHeaderRead: Boolean;
begin
  Files := TDirectory.GetFiles(FolderPath, '*.csv');
  CSVHeaderRead := false;
  MaxCol := 0;

  for FilePath in Files do
  begin
    CSVLines := TStringList.Create;
    filename := ExtractFileName(FilePath);
    FileNameStr := filename;
    try
      CSVLines.LoadFromFile(FilePath);

      if not CSVHeaderRead then
      begin
        MaxCol := Length(CSVLines[0].Split([',']));
        StringGridCSV.ColCount := MaxCol + 2; // Adding two for extra columns, one likely for 'Results' and one for 'Filename'
        FilenameColIndex := StringGridCSV.ColCount - 1; // The last column index
        StringGridCSV.RowCount := 1; // To skip the header row in the display
        CSVHeaderRead := True;
      end;

      for Row := 1 to CSVLines.Count - 1 do
      begin
        Line := CSVLines[Row];
        // Skip lines that consist only of commas
        if Line.Replace(',', '').Trim = '' then
          Continue;

        StringGridCSV.RowCount := StringGridCSV.RowCount + 1;
        var Cells := Line.Split([',']);

        // Populate the grid cells with CSV data
        for Col := 1 to High(Cells) + 1 do
          if Col <= MaxCol then
          begin
             StringGridCSV.Cells[Col, StringGridCSV.RowCount - 1] := Cells[Col - 1];
             StringGridCSV.Cells[28, StringGridCSV.RowCount - 1] :=  FileNameStr ;
        // Add filename to the last column, ensure you use the correct column inde
          end;

      end;

    finally
      CSVLines.Free;
    end;
  end;
end;

function TForm1.GetColumnIndexByHeaderName(StringGrid: TStringGrid;
  HeaderName: string): Integer;
var
  Col: Integer;
begin
  Result := -1; // Default result if header not found
  for Col := 0 to StringGrid.ColCount - 1 do
  begin
    if StringGrid.Cells[Col, 0] = HeaderName then
    // Assuming row 0 has the headers
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

procedure TForm1.Copytocsv1Click(Sender: TObject);
var
  s: string;
  Row, Col: Integer;
begin
  s := '';
  for Row := 0 to StringGridCSV.RowCount - 1 do
  begin
    for Col := 0 to StringGridCSV.ColCount - 1 do
    begin
      s := s + StringGridCSV.Cells[Col, Row];
      if Col < StringGridCSV.ColCount - 1 then
        s := s + #9; // Tab character
    end;
    if Row < StringGridCSV.RowCount - 1 then
      s := s + #13#10; // Newline characters
  end;
  Clipboard.AsText := s;
end;

procedure TForm1.StringGridCSVDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  S: string;
  InfoColumnIndex, TextWidth, NewColWidth: Integer;
  Grid: TStringGrid;
const
  CellPadding = 2; // Adjust the padding as needed
begin
  Grid := Sender as TStringGrid;
  S := Grid.Cells[ACol, ARow]; // Get the cell text
  InfoColumnIndex := GetColumnIndexByHeaderName(StringGridCSV, 'Info');

  // Calculate text width and adjust column width if necessary
  Grid.Canvas.Font := Grid.Font; // Use the grid's font for measurement
  TextWidth := Grid.Canvas.TextWidth(S) + 2 * CellPadding;
  NewColWidth := Max(Grid.ColWidths[ACol], TextWidth);
  if NewColWidth > Grid.ColWidths[ACol] then
    Grid.ColWidths[ACol] := NewColWidth;

  // Check if this is a header cell
  if ARow = 0 then
  begin
    if (ACol = 0) then
      Grid.Canvas.Brush.Color := clWebLightYellow
    else
      Grid.Canvas.Brush.Color := clWebLightBlue;
    // Use the specific color you want for the header
    // Header cell formatting

    Grid.Canvas.FillRect(Rect);
    Grid.Canvas.Font.Color := clWindowText;
    DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect,
      DT_CENTER or DT_VCENTER or DT_SINGLELINE);
  end
  else if ACol = 0 then // Change for your "Result" column index
  begin
    // "NG" cell formatting
    if AnsiStartsText('NG', S) then
    begin
      Grid.Canvas.Brush.Color := clYellow;
      // Entire cell background color for "NG" cells
      Grid.Canvas.Font.Color := clRed; // Text color for "NG" cells
      Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
    end
    else if S = 'Imported' then
    begin
      Grid.Canvas.Brush.Color := clWebLightGreen;
      // Entire cell background color for "NG" cells
      Grid.Canvas.Font.Color := clBlack; // Text color for "NG" cells
      Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
    end
    else
    begin
      Grid.Canvas.Brush.Color := clWindow; // Default background color for cells
      Grid.Canvas.Font.Color := clWindowText; // Default text color for cells
      Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
    end;
    DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect,
      DT_CENTER or DT_VCENTER or DT_SINGLELINE);
  end
  else if (ARow > 0) and (ACol = InfoColumnIndex) then
  begin
    // Check If Info no blank
    if S <> '' then
    begin
      Grid.Canvas.Brush.Color := clYellow;
      // Entire cell background color for "NG" cells
      Grid.Canvas.Font.Color := clRed; // Text color for "NG" cells
      Grid.Canvas.FillRect(Rect); // Fill the cell with the brush color
      // Align the text to the left with padding
      Inc(Rect.Left, CellPadding);
      DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect,
        DT_LEFT or DT_VCENTER or DT_SINGLELINE);
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
    DrawText(Grid.Canvas.Handle, PChar(S), Length(S), Rect,
      DT_LEFT or DT_VCENTER or DT_SINGLELINE);
  end;
end;

function TForm1.RoundDownTo(Value: Double; Decimals: Integer): Double;
var
  Factor: Double;
begin
  Factor := Power(10, Decimals);
  Result := Int(Value * Factor) / Factor;
end;

procedure TForm1.CreateStringGrid(var Grid: TStringGrid; AParent: TWinControl);
var
  i: Integer;
begin
  // Assign the OnDrawCell event handler
  Grid.OnDrawCell := StringGridCSVDrawCell;

  // Set number of columns
  Grid.ColCount := 30;
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
  Grid.Cells[CodeC, 0] := 'CodeC';
  Grid.Cells[filename, 0] := 'Filename';
  Grid.Cells[CodeB, 0] := 'CodeB';

  // Set column widths for result column
  Grid.ColWidths[0] := 150;
  Grid.ColAlignments[0] := taCenter;

  // Set grid options to show lines
  Grid.Options := Grid.Options + [goFixedVertLine, goFixedHorzLine, goVertLine,
    goHorzLine];

  // Responsive layout management
  Grid.Anchors := [akLeft, akTop, akRight, akBottom];

  // Optionally set the grid's parent
  Grid.Parent := AParent;

end;

procedure TForm1.SetIndex;
begin
  Result := 0;
  Date := 2;
  WorkerName := 3;
  EmployeeCode := 4;
  CodeD := 7;
  CostProcessName := 8;
  MoldCode := 10;
  Model := 11;
  LampName := 12;
  PartName := 13;
  ModifyJobNo := 9;
  PartCode := 14;
  PartMaster := 15;
  Start := 17;
  Finish := 18;
  Min := 19;
  MCCode := 20;
  Machmaster := 21;
  MachStart := 22;
  MachDate := 23;
  MachFinish := 24;
  MachMin := 25;
  ATC := 26;
  Remark := 27;
  Status := 29;
  CodeA := 5;
  Shift_n := 1;
  CodeC := 16;
  filename := 28;
  CodeB :=  6;    // code b

end;

procedure TForm1.AdjustLastColumnWidth(Grid: TStringGrid);
var
  TotalWidth, OtherColsWidth, i: Integer;
begin
  OtherColsWidth := 0;
  for i := 0 to Grid.ColCount - 2 do // Exclude the last column
    Inc(OtherColsWidth, Grid.ColWidths[i]);

  TotalWidth := Grid.ClientWidth - OtherColsWidth - Grid.GridLineWidth *
    (Grid.ColCount - 1);
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
  Result := false;
  RegEx := TRegEx.Create('^(?<hours>\d{2})\.(?<minutes>\d{2})$');
  Match := RegEx.Match(TimeStr);

  if Match.Success then
  begin
    if TryStrToInt(Match.Groups['hours'].Value, Hours) and
      TryStrToInt(Match.Groups['minutes'].Value, Minutes) then
    begin
      Result := (Hours >= 0) and (Hours <= 23) and (Minutes >= 0) and
        (Minutes <= 59);
    end;
  end;
end;

procedure MoveAllFiles;
var
  Files: TStringDynArray;
  filename, NewFileName, DateTimeStr: string;
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
      for filename in Files do
      begin
        NewFileName := TPath.Combine(MovePath,
          TPath.GetFileNameWithoutExtension(filename) + '_' + DateTimeStr +
          TPath.GetExtension(filename));
        TFile.Move(filename, NewFileName);
      end;
    end
    else
    begin
      // MovePath is empty, delete all files in the source directory
      for filename in Files do
        TFile.Delete(filename);
    end;
  end;

end;

procedure TForm1.file2Click(Sender: TObject);
begin
  halt;
end;

end.
