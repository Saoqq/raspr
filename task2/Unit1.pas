unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBCtrls, Grids, DBGrids, ADODB, ExtCtrls, StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    ADOConnection1: TADOConnection;
    ADOTable1: TADOTable;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    DataSource1: TDataSource;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
     ADOTable1.Active := true;
end;

procedure TForm1.Button2Click(Sender: TObject);
var SM,Desktop,Document,Text,Cursor,Dispatcher,Arr,Struct,Frm: Variant;

  i,j: Integer; ws: WideString;

begin


  SM:=CreateOleObject('com.sun.star.ServiceManager');

  Desktop:=SM.CreateInstance('com.sun.star.frame.Desktop');

  Document:=Desktop.LoadComponentFromURL(

      'private:factory/swriter', '_blank', 0,

      VarArrayCreate([0, -1], varVariant));

  Text:=Document.GetText;

  Cursor:=Document.CurrentController.ViewCursor;

  Text.InsertString(Cursor, 'Hello, World!'#13, false);

  j:=0; ws:='¹';

  for i:=0 to ADOTable1.FieldCount-1 do

      ws:=ws+#9+ADOTable1.Fields.Fields[i].FieldName;

  ws:=ws+#13;

  Text.InsertString(Cursor, ws, false);

  with ADOTable1 do while not Eof do begin

      j:=j+1;

      ws:=IntToStr(j)+'.';

      for i:=0 to FieldCount-1 do

        ws:=ws+#9+Fields.Fields[i].AsString;

      ws:=ws+#13;

      Text.InsertString(Cursor, ws, false);

      Next

  end;

  ADOTable1.Close;

  Cursor.JumpToFirstPage; Cursor.GoDown(1,false);

  Cursor.GoDown(j+1,true);

  Dispatcher:=SM.CreateInstance('com.sun.star.frame.DispatchHelper');

  Arr:=VarArrayCreate([0, 4], varVariant);

  Struct:=SM.Bridge_GetStruct('com.sun.star.beans.PropertyValue');

  Struct.Name:='Delimiter'; Struct.Value:=chr(9);

  Arr[0]:=Struct;

  Struct:=SM.Bridge_GetStruct('com.sun.star.beans.PropertyValue');

  Struct.Name:='WithHeader'; Struct.Value:=true;

  Arr[1]:=Struct;

  Struct:=SM.Bridge_GetStruct('com.sun.star.beans.PropertyValue');

  Struct.Name:='RepeatHeaderLines'; Struct.Value:=1;

  Arr[2]:=Struct;

  Struct:=SM.Bridge_GetStruct('com.sun.star.beans.PropertyValue');

  Struct.Name:='WithBorder'; Struct.Value:=true;

  Arr[3]:=Struct;

  Struct:=SM.Bridge_GetStruct('com.sun.star.beans.PropertyValue');

  Struct.Name:='DontSplitTable'; Struct.Value:=false;

  Arr[4]:=Struct;

  Frm:=Document.CurrentController.Frame;

  Dispatcher.executeDispatch(Frm, '.uno:ConvertTextToTable', '', 0, Arr);

end;

procedure TForm1.Button3Click(Sender: TObject);

var SM,Desktop,Document,Sheet,Arr: Variant; i,j: Integer;

begin

  SM:=CreateOleObject('com.sun.star.ServiceManager');

  Desktop:=SM.CreateInstance('com.sun.star.frame.Desktop');

  Document:=Desktop.LoadComponentFromURL(

      'private:factory/scalc', '_blank', 0,

      VarArrayCreate([0, -1], varVariant));

  Sheet:=Document.GetSheets.GetByIndex(0);

  Sheet.GetCellByPosition(0,0).String:='Hello, World!';

  Sheet.GetCellByPosition(0,1).String:='¹';

  j:=0;

  for i:=0 to ADOTable1.FieldCount-1 do

      Sheet.GetCellByPosition(i+1,1).String:=ADOTable1.Fields.Fields[i].FieldName;

  Arr := VarArrayCreate([0,ADOTable1.FieldCount,0,0],varOleStr);

  with ADOTable1 do while not Eof do begin

      j:=j+1;

      Arr[0,0] := IntToStr(j);

      for i:=0 to FieldCount-1 do

            Arr[i+1,0] := Fields.Fields[i].AsString;

      Sheet.GetCellRangeByName('A'+IntToStr(j+2)+':'+

            Chr(Ord('A')+FieldCount)+IntToStr(j+2)).SetDataArray(Arr);

      Next

  end;

  ADOTable1.Close

end;

end.
