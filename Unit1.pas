unit Unit1;
привет

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Data.Win.ADODB, ComObj, ExcelXP,
  Vcl.ExtCtrls;

type
  TForm1 = class(TForm)
    ADOConnection1: TADOConnection;
    ADOTable1: TADOTable;
    ADOTable2: TADOTable;
    ADOTable3: TADOTable;
    Button1: TButton;
    ADOQuery1: TADOQuery;
    Timer1: TTimer;
    Button2: TButton;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  WBk: OleVariant;
  Excel: Variant;
  t:TDateTime;
  h,m, fm, col, row, row_count, flag, p_min_i, p_max_i:integer;
  sql:String;
  p_min, p_max:Real;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
t:=now();
button1.Enabled:=false;
button2.Enabled:=false;
ADOConnection1.Connected:=true;
ADOTable1.Active:=true;
ADOTable2.Active:=true;
ADOTable3.Active:=true;
if (ADOTable1.FieldByName('h').AsInteger=ADOTable2.FieldByName('h').AsInteger)
  and (ADOTable1.FieldByName('m').AsInteger=ADOTable2.FieldByName('m').AsInteger) then
  begin
   try
    Excel := GetActiveOleObject('Excel.Application');
   except
    try
     Excel := CreateOleObject('Excel.Application');
    except
     Exception.Create('Error');
    end;
   end;
   WBk:=Excel.WorkBooks.Open('C:\Users\user\Documents\stakan_to_excel.xlsx');
   Excel.Application.EnableEvents := false;
   Excel.Visible :=true;
   h:=ADOTable1.FieldByName('h').AsInteger;
   fm:=ADOTable1.FieldByName('m').AsInteger;
   sql:='';
   col:=6;
   row:=3;
   p_min:=strtofloat(floattostrf(ADOTable3.FieldByName('pmin').AsFloat, ffFixed, 20,2));
   p_max:=strtofloat(floattostrf(ADOTable3.FieldByName('pmax').AsFloat, ffFixed, 20,2));
   row_count:=trunc((strtofloat(floattostrf(ADOTable3.FieldByName('pmax').AsFloat, ffFixed, 20,2))-
    strtofloat(floattostrf(ADOTable3.FieldByName('pmin').AsFloat, ffFixed, 20,2)))*100)+1;
   form1.Caption:=timetostr(now()-t);
   for h := ADOTable1.FieldByName('h').AsInteger to 18 do
    begin
    for m := fm to 59 do
     begin
      sql:='select a.h, a.m, vr_ma, pred from red a where a.h='+inttostr(h)+
           ' and a.m='+inttostr(m);
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Text:=sql;
      ADOQuery1.Active:=true;
      ADOQuery1.First;
      if ADOQuery1.RecordCount>0 then
       begin
        Excel.ActiveSheet.Cells.Item[1,col].value:=ADOQuery1.FieldByName('h').AsString+':'+ADOQuery1.FieldByName('m').AsString;
        while not ADOQuery1.Eof do
         begin
          row:=trunc(p_max*100-strtofloat(floattostrf(ADOQuery1.FieldByName('pred').AsFloat, ffFixed, 20,2))*100)+3;
          Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pred').AsString;
          Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vr_ma').AsInteger;
          ADOQuery1.Next;
          form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
         end;
        ADOQuery1.Active:=false;
        col:=col+1;
        sql:='select a.h, a.m, vg_ma, pgreen from green a where a.h='+inttostr(h)+
             ' and a.m='+inttostr(m);
        ADOQuery1.SQL.Clear;
        ADOQuery1.SQL.Text:=sql;
        ADOQuery1.Active:=true;
        ADOQuery1.First;
        while not ADOQuery1.Eof do
         begin
          row:=trunc(p_max*100-strtofloat(floattostrf(ADOQuery1.FieldByName('pgreen').AsFloat, ffFixed, 20,2))*100)+3;
          Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pgreen').AsString;
          Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vg_ma').AsInteger;
          ADOQuery1.Next;
          form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
         end;
        ADOQuery1.Active:=false;
        col:=col+1;
       end
        else ADOQuery1.Active:=false;
     end;
     fm:=0;
     form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
    end;
  end
   else showmessage('Не совпадает время в бид и оффер.');
form1.Caption:='Прошло:'+timetostr(now()-t);
Excel.Visible :=true;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
if form1.Tag=0 then
 begin
   try
    Excel := GetActiveOleObject('Excel.Application');
   except
    try
     Excel := CreateOleObject('Excel.Application');
    except
     Exception.Create('Error');
    end;
   end;
   WBk:=Excel.WorkBooks.Open('C:\Users\user\Documents\stakan_to_excel.xlsx');
   Excel.Application.EnableEvents := false;
   Excel.Visible :=true;
   form1.Tag:=1;
  t:=now();
  button1.Enabled:=false;
  ADOConnection1.Connected:=true;
  ADOTable1.Active:=true;
  ADOTable1.First;
  h:=ADOTable1.FieldByName('h').AsInteger;
  m:=ADOTable1.FieldByName('m').AsInteger;
  p_max:=strtofloat(floattostrf(ADOTable1.FieldByName('pred').AsFloat, ffFixed, 20,2));
  ADOTable1.Active:=false;
  col:=2;
  row:=1001;
 end;
if flag=0 then
 begin
 timer1.Enabled:=true;
 flag:=1;
 button2.Caption:='Остановить следование за минутными свечами.';
 end
  else
  begin
  timer1.Enabled:=false;
  flag:=0;
  button2.Caption:='Запустить следование за минутными свечами.';
  end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
t:=now();
button1.Enabled:=false;
button2.Enabled:=false;
ADOConnection1.Connected:=true;
ADOTable1.Active:=true;
ADOTable2.Active:=true;
ADOTable3.Active:=true;
if (ADOTable1.FieldByName('h').AsInteger=ADOTable2.FieldByName('h').AsInteger)
  and (ADOTable1.FieldByName('m').AsInteger=ADOTable2.FieldByName('m').AsInteger) then
  begin
   try
    Excel := GetActiveOleObject('Excel.Application');
   except
    try
     Excel := CreateOleObject('Excel.Application');
    except
     Exception.Create('Error');
    end;
   end;
   WBk:=Excel.WorkBooks.Open('C:\Users\user\Documents\stakan_to_excel.xlsx');
   Excel.Application.EnableEvents := false;
   Excel.Visible :=true;
   h:=ADOTable1.FieldByName('h').AsInteger;
   fm:=ADOTable1.FieldByName('m').AsInteger;
   sql:='';
   col:=6;
   row:=3;
   p_min_i:=ADOTable3.FieldByName('pmin').AsInteger;
   p_max_i:=ADOTable3.FieldByName('pmax').AsInteger;
   row_count:=p_max_i-p_min_i+2;
   form1.Caption:=timetostr(now()-t);
   for h := ADOTable1.FieldByName('h').AsInteger to 18 do
    begin
    for m := fm to 59 do
     begin
      sql:='select a.h, a.m, vr_ma, pred from red a where a.h='+inttostr(h)+
           ' and a.m='+inttostr(m);
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Text:=sql;
      ADOQuery1.Active:=true;
      ADOQuery1.First;
      if ADOQuery1.RecordCount>0 then
       begin
        Excel.ActiveSheet.Cells.Item[1,col].value:=ADOQuery1.FieldByName('h').AsString+':'+ADOQuery1.FieldByName('m').AsString;
        while not ADOQuery1.Eof do
         begin
          row:=p_max_i-ADOQuery1.FieldByName('pred').AsInteger+3;
          Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pred').AsString;
          Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vr_ma').AsInteger;
          ADOQuery1.Next;
          form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
         end;
        ADOQuery1.Active:=false;
        col:=col+1;
        sql:='select a.h, a.m, vg_ma, pgreen from green a where a.h='+inttostr(h)+
             ' and a.m='+inttostr(m);
        ADOQuery1.SQL.Clear;
        ADOQuery1.SQL.Text:=sql;
        ADOQuery1.Active:=true;
        ADOQuery1.First;
        while not ADOQuery1.Eof do
         begin
          row:=p_max_i-ADOQuery1.FieldByName('pgreen').AsInteger+3;
          Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pgreen').AsString;
          Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vg_ma').AsInteger;
          ADOQuery1.Next;
          form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
         end;
        ADOQuery1.Active:=false;
        col:=col+1;
       end
        else ADOQuery1.Active:=false;
     end;
     fm:=0;
     form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
    end;
  end
   else showmessage('Не совпадает время в бид и оффер.');
form1.Caption:='Прошло:'+timetostr(now()-t);
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
flag:=0;
form1.Tag:=0;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
 hn, mn:integer;
 tstr:String;
begin
timer1.Enabled:=false;
form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
tstr:=timetostr(now());
delete(tstr,1,pos(':',tstr));
mn:=strtoint(copy(tstr,1,pos(':',tstr)-1));
if mn<>m then
 begin
  sql:='select a.h, a.m, vr_ma, pred from red a where a.h='+inttostr(h)+
       ' and a.m='+inttostr(m);
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Text:=sql;
  ADOQuery1.Active:=true;
  ADOQuery1.First;
  while (ADOQuery1.RecordCount>0) and (mn<>m) do
  begin
    Excel.ActiveSheet.Cells.Item[1,col].value:=ADOQuery1.FieldByName('h').AsString+':'+ADOQuery1.FieldByName('m').AsString;
    while not ADOQuery1.Eof do
     begin
      row:=trunc(p_max*100-strtofloat(floattostrf(ADOQuery1.FieldByName('pred').AsFloat, ffFixed, 20,2))*100)+991;
      Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pred').AsString;
      Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vr_ma').AsInteger;
      ADOQuery1.Next;
      form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
     end;
    ADOQuery1.Active:=false;
    col:=col+1;
    sql:='select a.h, a.m, vg_ma, pgreen from green a where a.h='+inttostr(h)+
         ' and a.m='+inttostr(m);
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Text:=sql;
    ADOQuery1.Active:=true;
    ADOQuery1.First;
    while not ADOQuery1.Eof do
     begin
      row:=trunc(p_max*100-strtofloat(floattostrf(ADOQuery1.FieldByName('pgreen').AsFloat, ffFixed, 20,2))*100)+991;
      Excel.ActiveSheet.Cells.Item[row,1].value:=ADOQuery1.FieldByName('pgreen').AsString;
      Excel.ActiveSheet.Cells.Item[row,col].value:=ADOQuery1.FieldByName('vg_ma').AsInteger;
      ADOQuery1.Next;
      form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
     end;
    ADOQuery1.Active:=false;
    form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
    m:=m+1;
    if m=60 then
     begin
      m:=0;
      h:=h+1;
     end;
    col:=col+1;
    sql:='select a.h, a.m, vr_ma, pred from red a where a.h='+inttostr(h)+
         ' and a.m='+inttostr(m);
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Text:=sql;
    ADOQuery1.Active:=true;
    ADOQuery1.First;
    tstr:=timetostr(now());
    delete(tstr,1,pos(':',tstr));
    mn:=strtoint(copy(tstr,1,pos(':',tstr)-1));
  end;
  ADOQuery1.Active:=false;
 end;
form1.Caption:=inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
if flag=1 then
begin
timer1.Enabled:=true;
form1.Caption:='Ожидание свечи. '+inttostr(h)+':'+inttostr(m)+'; '+timetostr(now()-t);
end;
end;

end.
