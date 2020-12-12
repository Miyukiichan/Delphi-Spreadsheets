object Spreadsheet: TSpreadsheet
  Left = 0
  Top = 0
  Caption = 'Spreadsheet'
  ClientHeight = 631
  ClientWidth = 848
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object coordLabel: TLabel
    Left = 302
    Top = 8
    Width = 20
    Height = 19
    Caption = 'A1'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object StringGrid: TStringGrid
    Left = 0
    Top = 35
    Width = 849
    Height = 582
    ColCount = 2000
    RowCount = 2000
    TabOrder = 0
    OnClick = StringGridClick
    OnDrawCell = StringGridDrawCell
    OnMouseDown = StringGridMouseDown
    OnMouseUp = StringGridMouseUp
  end
  object cellEdit: TEdit
    Left = 8
    Top = 8
    Width = 281
    Height = 21
    TabOrder = 1
    OnChange = cellEditChange
    OnKeyPress = cellEditKeyPress
  end
  object statusBar: TStatusBar
    Left = 0
    Top = 612
    Width = 848
    Height = 19
    Panels = <
      item
        Width = 50
      end>
  end
  object MainMenu: TMainMenu
    Left = 792
    Top = 8
    object FileMenu: TMenuItem
      Caption = 'File'
      object NewAction: TMenuItem
        Caption = 'New'
        OnClick = NewActionClick
      end
      object LoadAction: TMenuItem
        Caption = 'Load'
        OnClick = LoadActionClick
      end
      object SaveAction: TMenuItem
        Caption = 'Save'
        OnClick = SaveActionClick
      end
      object SaveAsAction: TMenuItem
        Caption = 'Save As'
        OnClick = SaveAsActionClick
      end
    end
    object EditMenu: TMenuItem
      Caption = 'Edit'
    end
  end
  object openFile: TOpenTextFileDialog
    Left = 736
    Top = 8
  end
  object saveFile: TSaveTextFileDialog
    Left = 688
    Top = 8
  end
  object PopupMenu: TPopupMenu
    Left = 640
    Top = 8
    object TestAction: TMenuItem
      Caption = 'Test Action'
    end
  end
end
