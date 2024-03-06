unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Uni, UniProvider, OracleUniProvider, MemDS, Vcl.Grids, Vcl.DBGrids,
  DBAccess, Vcl.ExtDlgs,System.IniFiles ,DateUtils, System.ImageList,
  Vcl.ImgList, Vcl.Buttons, Vcl.Menus, Winapi.Winsock   , System.Types  ,System.IOUtils

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
  UniQuery.SQL.Text := 'select * from DENSO_IF_NEDO_DATA_PLAN';
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
begin
  //wait
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
        MaxCol := Length(HeaderCells);
        StringGridCSV.ColCount := MaxCol;
        StringGridCSV.RowCount := 1;
        // Set the header
        for Col := 0 to High(HeaderCells) do
          StringGridCSV.Cells[Col, 0] := HeaderCells[Col];
        CSVHeaderRead := True;
      end;

      // Skip the header line for all files
      for Row := 1 to CSVLines.Count - 1 do
      begin
        Line := CSVLines[Row];
        var Cells := Line.Split([',']);
        // Add a new row to the grid for each line read
        StringGridCSV.RowCount := StringGridCSV.RowCount + 1;
        // Fill the row with data
        for Col := 0 to High(Cells) do
        begin
          if Col < MaxCol then
            StringGridCSV.Cells[Col, StringGridCSV.RowCount - 1] := Cells[Col];
        end;
      end;
    finally
      CSVLines.Free;
    end;
  end;
end;

end.

