object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 445
  ClientWidth = 870
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object SpeedButtonIMP: TSpeedButton
    Left = 8
    Top = 8
    Width = 73
    Height = 58
    Caption = 'Import'
    OnClick = SpeedButtonIMPClick
  end
  object EditFolderPath: TEdit
    Left = 232
    Top = 21
    Width = 257
    Height = 23
    Enabled = False
    TabOrder = 0
  end
  object OpenFolderPath: TButton
    Left = 511
    Top = 20
    Width = 42
    Height = 25
    Caption = 'Open'
    TabOrder = 1
    OnClick = OpenFolderPathClick
  end
  object ButtonReadCSV: TButton
    Left = 95
    Top = -1
    Width = 66
    Height = 67
    Caption = 'Read'
    TabOrder = 2
    OnClick = ButtonReadClick
  end
  object StringGridCSV: TStringGrid
    Left = 8
    Top = 80
    Width = 854
    Height = 362
    TabOrder = 3
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 426
    Width = 870
    Height = 19
    Panels = <>
    ExplicitLeft = 440
    ExplicitTop = 232
    ExplicitWidth = 0
  end
  object FolderDialog: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <>
    Options = []
    Left = 744
    Top = 8
  end
  object UniConnection: TUniConnection
    Left = 632
    Top = 8
  end
  object UniQuery: TUniQuery
    Connection = UniConnection
    Left = 792
    Top = 8
  end
  object OracleUniProvider: TOracleUniProvider
    Left = 568
    Top = 32
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 688
    Top = 16
  end
end
