object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 299
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 635
    Height = 49
    Align = alTop
    Caption = 'Panel1'
    TabOrder = 0
    object Button1: TButton
      Left = 16
      Top = 9
      Width = 75
      Height = 25
      Caption = 'connect'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 400
      Top = 9
      Width = 75
      Height = 25
      Caption = 'writer'
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 518
      Top = 9
      Width = 75
      Height = 25
      Caption = 'calc'
      TabOrder = 2
      OnClick = Button3Click
    end
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 49
    Width = 635
    Height = 250
    Align = alClient
    DataSource = DataSource1
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object DBNavigator1: TDBNavigator
    Left = 113
    Top = 8
    Width = 240
    Height = 25
    DataSource = DataSource1
    TabOrder = 2
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 'Provider=MSDASQL.1;Persist Security Info=False;Data Source=ole1'
    Provider = 'MSDASQL.1'
    Left = 584
    Top = 72
  end
  object ADOTable1: TADOTable
    Connection = ADOConnection1
    CursorType = ctStatic
    TableName = 'country'
    Left = 584
    Top = 120
  end
  object DataSource1: TDataSource
    DataSet = ADOTable1
    Left = 584
    Top = 160
  end
end
