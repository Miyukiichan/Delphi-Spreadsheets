unit MainScreen;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ToolWin, Vcl.ActnMan, Vcl.ActnCtrls,
  Vcl.ActnMenus, Vcl.Grids, Vcl.Menus, Vcl.StdCtrls, Math, Vcl.ComCtrls,
  Vcl.ExtDlgs, System.UITypes, System.Actions, Vcl.ActnList, Vcl.StdActns,
  System.ImageList, Vcl.ImgList, Vcl.VirtualImageList;

type
  TSpreadsheet = class(TForm)
    StringGrid: TStringGrid;
    MainMenu: TMainMenu;
    FileMenu: TMenuItem;
    EditMenu: TMenuItem;
    cellEdit: TEdit;
    coordLabel: TLabel;
    NewAction: TMenuItem;
    LoadAction: TMenuItem;
    SaveAction: TMenuItem;
    SaveAsAction: TMenuItem;
    statusBar: TStatusBar;
    openFile: TOpenTextFileDialog;
    saveFile: TSaveTextFileDialog;
    PopupMenu: TPopupMenu;
    TestAction: TMenuItem;
    procedure StringGridClick(Sender: TObject);
    procedure cellEditKeyPress(Sender: TObject; var Key: Char);
    procedure cellEditChange(Sender: TObject);
    procedure StringGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure StringGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SaveActionClick(Sender: TObject);
    procedure NewActionClick(Sender: TObject);
    procedure LoadActionClick(Sender: TObject);
    procedure SaveAsActionClick(Sender: TObject);
    procedure StringGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    procedure SetCellEdit;
    procedure SetCurrentCells;
    procedure SetCoordLabel(X: Integer; Y: Integer);
    function AlphaIndex(index: Integer): string;
    procedure Save;
    function SaveAs: Boolean;
    procedure Load;
    procedure ClearGrid;
    procedure SetFilePromptDefaults(prompt: TOpenTextFileDialog);
    procedure ErrorMsg(msg: string);
  public
    constructor Create(AOwner: TComponent); override;
    procedure SetPath(p: string);
  end;

var
  Sheet: TSpreadsheet;
  blockChange: Boolean;
  Path: string;

implementation

{$R *.dfm}

{ Sync content of selected cells with edit box if user edits it }
procedure TSpreadsheet.cellEditChange(Sender: TObject);
begin
  if not blockChange then SetCurrentCells;
end;

{ Allow editing to affect selected cells }
procedure TSpreadsheet.cellEditKeyPress(Sender: TObject; var Key: Char);
begin
  blockChange := false;
end;

{ User should be able to edit a cell they click on using the edit box }
procedure TSpreadsheet.StringGridClick(Sender: TObject);
begin
  cellEdit.SetFocus;
end;

{ Make the title row/column grey and center the label text }
procedure TSpreadsheet.StringGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  offset: Integer;
begin
  if ((ACol = 0) or (ARow = 0)) and (ARow <> ACol) then
    with TStringGrid(Sender) do begin
      offset := 27;
      if ARow = 0 then offset := 30;
      Dec(offset, ((Length(Cells[ACol, ARow]) - 1) * 3));
      Canvas.Brush.Color := clGray;
      Rect.Left := Rect.Left - 4;
      Canvas.FillRect(Rect);
      Canvas.TextOut(Rect.Left + offset, Rect.Top + 5, Cells[ACol, ARow]);
    end;

end;

{ Sync cell edit and block changes to it affecting selected cells }
procedure TSpreadsheet.StringGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  blockChange := true;
  SetCellEdit;
  SetCoordLabel(X, Y);
end;

procedure TSpreadsheet.StringGridMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  pnt: TPoint;
begin
  if (Button = mbRight) and (GetCursorPos(pnt)) then
    PopupMenu.Popup(pnt.X, pnt.y);
end;

procedure TSpreadsheet.SaveActionClick(Sender: TObject);
begin
  Save;
end;

{ Prompt the user for a file and then save if successful }
procedure TSpreadsheet.SaveAsActionClick(Sender: TObject);
begin
  if SaveAs then Save;
end;

procedure TSpreadsheet.LoadActionClick(Sender: TObject);
begin
  Load;
end;

procedure TSpreadsheet.NewActionClick(Sender: TObject);
begin
  ClearGrid;
  cellEdit.Text := string.Empty;
  SetPath(string.Empty);
end;

{ Constructor }
constructor TSpreadsheet.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited Create(AOwner);
  blockChange := false;
  for i := 1 to StringGrid.RowCount do StringGrid.Cells[0, i] := IntToStr(i);
  for i := 1 to StringGrid.ColCount do StringGrid.Cells[i, 0] := AlphaIndex(i);
  SetFilePromptDefaults(openFile);
  SetFilePromptDefaults(saveFile);
  SetPath(string.Empty);
  NewAction.ShortCut := ShortCut(Word('N'), [ssCtrl]);
  LoadAction.ShortCut := ShortCut(Word('O'), [ssCtrl]);
  SaveAction.ShortCut := ShortCut(Word('S'), [ssCtrl]);
  SaveAsAction.ShortCut := ShortCut(Word('S'), [ssCtrl, ssShift]);
end;

{ Set the cell editor to be the content of the top left cell in the selection }
procedure TSpreadsheet.SetCellEdit;
var
  coord: TGridCoord;
begin
  coord := StringGrid.Selection.TopLeft;
  cellEdit.Text := StringGrid.Cells[coord.X, coord.Y];
end;

{ Set the values of the selected cells to the cell editor content }
procedure TSpreadsheet.SetCurrentCells;
var
  Selection: TGridRect;
  i: Integer;
  j: Integer;
begin
  Selection := StringGrid.Selection;
  for i := Selection.Left to Selection.Right do begin
    for j := Selection.Top to Selection.Bottom do
        StringGrid.Cells[i, j] := cellEdit.Text;
  end;
end;

{ Write the current cell coordinates to the coordinate label }
procedure TSpreadsheet.SetCoordLabel(X: Integer; Y: Integer);
var
  xStr: string;
  yStr: string;
begin
  StringGrid.MouseToCell(X, Y, X, Y);
  yStr := IntToStr(Y);
  xStr := AlphaIndex(X);
  coordLabel.Caption := Concat(xStr, yStr);
end;

{ Convert the given base 10 number to base 26 where 1 = A and 26 = Z etc }
function TSpreadsheet.AlphaIndex(index: Integer): string;
var
  i: Integer;
  j: Integer;
  Arr: array [0 .. 20] of Char;
  Temp: Integer;
begin
  Result := string.Empty;
  i := 0;
  while Index <> 0 do begin
    Temp := Pred(Index) mod 26;
    Arr[i] := Char(Temp + 65);
    Index := Pred(Index) div 26;
    Inc(i);
  end;
  for j := Pred(i) downto 0 do Result := Concat(Result, Arr[j]);
end;

{ Find the maximum cell coord values and save to csv accordingly }
procedure TSpreadsheet.Save;
var
  max_x: Integer;
  max_y: Integer;
  i: Integer;
  j: Integer;
  found: Boolean;
  f: TextFile;
  line: string;
begin
  if Path = string.Empty then begin
    if not SaveAs then Exit;
  end;
  max_x := 0;
  max_y := 0;
  found := false;
  for i := StringGrid.ColCount downto 1 do begin
    for j := StringGrid.RowCount downto 1 do begin
      if StringGrid.Cells[i, j] <> string.Empty then begin
        max_x := i;
        max_y := j;
        found := true;
        Break;
      end;
    end;
    if found then Break;
  end;
  if found then begin
    found := false;
    for i := StringGrid.RowCount downto max_y do begin
      for j := StringGrid.ColCount downto 1 do begin
        if (StringGrid.Cells[j, i] <> string.Empty) then begin
          max_y := i;
          found := true;
          Break;
        end;
      end;
      if found then Break;
    end;
  end;
  AssignFile(f, Path);
  try
    Rewrite(f);
    for i := 1 to max_y do begin
      line := string.Empty;
      for j := 1 to max_x do begin
        line := Concat(line, StringGrid.Cells[j, i]);
        if j < max_x then line := Concat(line, ',');
      end;
      WriteLn(f, line);
    end;
  except ErrorMsg('Error writing to file');
  end;
  CloseFile(f);
end;

{ Prompt the user for a filename and return a boolean based on success }
function TSpreadsheet.SaveAs: Boolean;
begin
  if not saveFile.Execute then Exit(false);
  SetPath(saveFile.filename);
  Exit(true);
end;

{ Prompt the user for  file then parse it as csv and populate the grid with contents }
procedure TSpreadsheet.Load;
var
  f: TextFile;
  line: string;
  lines: TArray<string>;
  i: Integer;
  counter: Integer;
  filename: string;
begin
  if not openFile.Execute then Exit;
  filename := openFile.filename;
  if not FileExists(filename) then begin
    ErrorMsg('File name entered does not exist on system');
    Exit;
  end;
  SetPath(filename);
  ClearGrid;
  AssignFile(f, Path);
  counter := 1;
  try
    Reset(f);
    while not Eof(f) do begin
      ReadLn(f, line);
      lines := line.Split([',']);
      for i := 0 to Length(lines) do
          StringGrid.Cells[i + 1, counter] := lines[i];
      Inc(counter);
    end;
  except ErrorMsg('Error loading file');
  end;
  CloseFile(f);
end;

procedure TSpreadsheet.ClearGrid;
var
  i: Integer;
  j: Integer;
begin
  for i := 1 to StringGrid.ColCount do begin
    for j := 1 to StringGrid.RowCount do
      if (StringGrid.Cells[i, j] <> string.Empty) then
          StringGrid.Cells[i, j] := string.Empty;
  end;
end;

{ Set the current path for the file being edited and display accordingly in the
  statusbar }
procedure TSpreadsheet.SetPath(p: string);
var
  Text: string;
begin
  Path := p;
  if Path = string.Empty then Text := 'New File'
  else Text := Path;
  statusBar.Panels[0].Text := Text;
end;

{ Sets common options for file dialogs }
procedure TSpreadsheet.SetFilePromptDefaults(prompt: TOpenTextFileDialog);
var
  filter: string;
const
  defExt = 'csv';
begin
  filter := 'Microsoft Excel Comma Separated Values File (*.csv)|*.csv';
  prompt.filter := filter;
  prompt.DefaultExt := defExt;
end;

procedure TSpreadsheet.ErrorMsg(msg: string);
begin
  MessageDlg(msg, mtError, [mbOK], 0);
end;

end.
