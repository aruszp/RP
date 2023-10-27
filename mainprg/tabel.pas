unit tabel;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  JvLabel, JvExControls, JvComponent, JvGradient, StdCtrls, JvExStdCtrls,
  JvEdit, JvXPCore, JvXPButtons, mainlib, ToolEdit, DB;

type
  TForm2060 = class(TForm)
    JvXPButton1: TJvXPButton;
    JvXPButton2: TJvXPButton;
    JvEdit1: TJvEdit;
    JvEdit2: TJvEdit;
    JvEdit3: TJvEdit;
    JvGradient1: TJvGradient;
    JvLabel1: TJvLabel;
    JvLabel2: TJvLabel;
    JvLabel3: TJvLabel;
    JvLabel13: TJvLabel;
    JvLabel14: TJvLabel;
    JvLabel15: TJvLabel;
    JvLabel16: TJvLabel;
    JvLabel21: TJvLabel;
    JvLabel22: TJvLabel;
    JvLabel23: TJvLabel;
    JvLabel24: TJvLabel;
    CheckBox1: TCheckBox;
    JvLabel4: TJvLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    JvXPButton3: TJvXPButton;
    ComboBox1: TComboBox;
    JvLabel5: TJvLabel;
    procedure JvXPButton2Click(Sender: TObject);
   
    procedure FProv;
    procedure SetColor;
    procedure SetCreate;
    procedure FormShow(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure JvXPButton1Click(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure JvXPButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
   TNDEdit2,TNdEdit4:array[1..31] of TComboEdit;
   TNDEdit,TNdEdit3:array[1..31] of TjvEdit;
   TNDLabel:array[1..31] of TLabel;

    { Public declarations }
  end;

var
  Form2060: TForm2060;

implementation

uses pevazp, sptab_grid, vyb_sptabel, MyDATAMODULE;

{$R *.DFM}

procedure TForm2060.JvXPButton2Click(Sender: TObject);
begin
 form2060.close;
end;

procedure TForm2060.FProv;
var x,x2:Real;
    i,ik:integer;
    s0:String;
begin

 for i:=1 to NDayMes(RMes,RGod) do
  begin

   if trim(TNdEdit[i].Text)='' then x:=0 else
    begin
     try

      s0:=trim(TNdEdit[i].Text);
      for ik:=1 to LengTh(s0) do
       begin
        if DecimalSeparator=',' then if s0[ik]='.' then s0[ik]:=',' ;
        if DecimalSeparator='.' then if s0[ik]=',' then s0[ik]:='.' ;
       end;

     TNdEdit[i].Text:=s0;

      x:=StrToFloat(s0);
      if (x<>0) and (trim(TNdEdit2[i].Text)='') then TNdEdit2[i].Text:='Я';
     except
       MessageDlg('Ошибка ввода',mtInformation,[mbOk],0);
       TNdEdit[i].SetFocus;
       exit;
     end;
     end;

    if trim(TNdEdit3[i].Text)='' then x2:=0 else
     begin
      try

      s0:=trim(TNdEdit3[i].Text);
      for ik:=1 to LengTh(s0) do
       begin
        if DecimalSeparator=',' then if s0[ik]='.' then s0[ik]:=',' ;
        if DecimalSeparator='.' then if s0[ik]=',' then s0[ik]:='.' ;
       end;

      TNdEdit3[i].Text:=s0;

       x2:=StrToFloat(s0);
      except
      MessageDlg('Ошибка ввода',mtInformation,[mbOk],0);
       TNdEdit3[i].SetFocus;
       exit;
      end;
    end;
 end;

end;


procedure TForm2060.SetCreate;
var i:integer;
begin
 for i:=1 to 15 do
  begin

   TNdEdit2[i]:=TComboEdit.Create(form2060);
   TNdEdit2[i].Parent:=form2060;
   TNdEdit2[i].Top:=120;
   TNdEdit2[i].Left:=70+i*50;
   TNdEdit2[i].Width:=50;
   TNdEdit2[i].ButtonWidth:=12;

   TNdEdit[i]:=TjvEdit.Create(form2060);
   TNdEdit[i].Parent:=form2060;
   TNdEdit[i].Top:=145;
   TNdEdit[i].Left:=70+i*50;
   TNdEdit[i].Width:=35;

   TNdEdit4[i]:=TComboEdit.Create(form2060);
   TNdEdit4[i].Parent:=form2060;
   TNdEdit4[i].Top:=190;
   TNdEdit4[i].Left:=70+i*50;
   TNdEdit4[i].Width:=50;
   TNdEdit4[i].ButtonWidth:=12;

   TNdEdit3[i]:=TjvEdit.Create(form2060);
   TNdEdit3[i].Parent:=form2060;
   TNdEdit3[i].Top:=215;
   TNdEdit3[i].Left:=70+i*50;
   TNdEdit3[i].Width:=35;

   TNdEdit2[i].OnButtonClick:=Edit2Change;
   TNdEdit2[i].Button.Hint:='Выбор из справочника';
   TNdEdit2[i].Button.ShowHint:=true;
   TNdEdit4[i].Button.Hint:='Выбор из справочника';
   TNdEdit4[i].Button.ShowHint:=true;
   TNdEdit4[i].OnButtonClick:=Edit2Change;
   TNdEdit[i].OnExit:=Edit1Change;
   TNdEdit3[i].OnExit:=Edit1Change;

   TNdLabel[i]:=TLabel.Create(form2060);
   TNdLabel[i].Parent:=form2060;
   TNdLabel[i].Top:=100;
   TNdLabel[i].Left:=70+i*50;
   TNdLabel[i].Caption:=floattostr(i);
   TNdLabel[i].Transparent:=true;

  end;
 for i:=16 to 31 do
  begin


   TNdEdit2[i]:=TComboEdit.Create(form2060);
   TNdEdit2[i].Parent:=form2060;
   TNdEdit2[i].Top:=350;
   TNdEdit2[i].Left:=70+(i-15)*50;
   TNdEdit2[i].Width:=50;
   TNdEdit2[i].Visible:=true;
   TNdEdit2[i].ButtonWidth:=12;

   TNdEdit[i]:=TjvEdit.Create(form2060);
   TNdEdit[i].Parent:=form2060;
   TNdEdit[i].Top:=375;
   TNdEdit[i].Left:=70+(i-15)*50;
   TNdEdit[i].Width:=35;
   TNdEdit[i].Visible:=true;

   TNdEdit4[i]:=TComboEdit.Create(form2060);
   TNdEdit4[i].Parent:=form2060;
   TNdEdit4[i].Top:=415;
   TNdEdit4[i].Left:=70+(i-15)*50;
   TNdEdit4[i].Width:=50;
   TNdEdit4[i].ButtonWidth:=12;

   TNdEdit3[i]:=TjvEdit.Create(form2060);
   TNdEdit3[i].Parent:=form2060;
   TNdEdit3[i].Top:=440;
   TNdEdit3[i].Left:=70+(i-15)*50;
   TNdEdit3[i].Width:=35;

   TNdEdit[i].OnExit:=Edit1Change;
   TNdEdit3[i].OnExit:=Edit1Change;
   TNdEdit2[i].OnButtonClick:=Edit2Change;
   TNdEdit4[i].OnButtonClick:=Edit2Change;
   TNdEdit2[i].Button.Hint:='Выбор из справочника';
   TNdEdit2[i].Button.ShowHint:=true;
   TNdEdit4[i].Button.Hint:='Выбор из справочника';
   TNdEdit4[i].Button.ShowHint:=true;

   TNdLabel[i]:=TLabel.Create(form2060);
   TNdLabel[i].Parent:=form2060;
   TNdLabel[i].Top:=330;
   TNdLabel[i].Left:=70+(i-15)*50;
   TNdLabel[i].Caption:=floattostr(i);
   TNdLabel[i].Transparent:=True;

  end;

  for i:=1 to 31 do
   begin
    TNdEdit[i].Font.Color:=clBlack;
    TNdEdit2[i].Font.Color:=clBlack;
   end;


end;


procedure TForm2060.SetColor;
var i,k:integer;
     td:TDate;
     s:string;
begin

 k:=NDayMes(Trunc(RMes),Trunc(RGod));
 for i:=1 to 31 do
  begin
   if i<=k then
     begin
      TNdLabel[i].Font.Color:=clBlack;
      TNdLabel[i].Font.Name:='Calibri';
      TNdLabel[i].Font.Size:=10;
      TNdEdit[i].Visible:=True;

      TNdEdit[i].Font.Name:='Calibri';
      TNdEdit[i].Font.Size:=10;
      TNdEdit2[i].Font.Name:='Calibri';
      TNdEdit2[i].Font.Size:=10;
      TNdEdit3[i].Font.Name:='Calibri';
      TNdEdit3[i].Font.Size:=10;
      TNdEdit4[i].Font.Name:='Calibri';
      TNdEdit4[i].Font.Size:=10;


      TNdEdit2[i].Visible:=True;

      TNdEdit3[i].Visible:=True;
      TNdEdit4[i].Visible:=True;

      TNdLabel[i].Visible:=True;
      td:=EncodeDate(Trunc(RGod),Trunc(RMes),i);
      s:=NameDay(td);
      TndLabel[i].Caption:=floattostr(i)+' ['+s+']';
      TNdLabel[i].Font.Color:=clBlack;
      TNdEdit[i].Font.Color:=clBlack;
       TNdEdit2[i].Font.Color:=clBlack;
      TNdEdit[i].Alignment:=taCenter;
      TNdEdit2[i].Alignment:=taCenter;

      TNdEdit3[i].Font.Color:=clBlack;
       TNdEdit4[i].Font.Color:=clBlack;
      TNdEdit3[i].Alignment:=taCenter;
      TNdEdit4[i].Alignment:=taCenter;

      if form9.WPRAZDNIK[i] then
       begin
        TNdLabel[i].Font.Color:=clMaroon;
        TNdEdit[i].Font.Color:=clMaroon;
        TNdEdit2[i].Font.Color:=clMaroon;
        TNdEdit3[i].Font.Color:=clMaroon;
        TNdEdit4[i].Font.Color:=clMaroon;
       {
        TNdEdit[i].Color:=$00D8C8FD;
        TNdEdit2[i].Color:=$00D8C8FD;
        TNdEdit3[i].Color:=$00D8C8FD;
        TNdEdit4[i].Color:=$00D8C8FD;
       }

        TNdEdit[i].Color:=$00F8F5FE;
        TNdEdit2[i].Color:=$00F8F5FE;
        TNdEdit3[i].Color:=$00F8F5FE;
        TNdEdit4[i].Color:=$00F8F5FE;


       end;


   //    TNdEdit2[i].Font.Style:=[fsBold];
   //    TNdEdit4[i].Font.Style:=[fsBold];

   //    TNdEdit[i].Font.Style:=[fsBold];
   //    TNdEdit4[i].Font.Style:=[fsBold];

    end
     else
    begin
      TNdEdit[i].Visible:=False;
       TNdEdit2[i].Visible:=False;
       TNdEdit3[i].Visible:=False;
       TNdEdit4[i].Visible:=False;
      TNdLabel[i].Visible:=False;

      TNdEdit[i].Text:='';
      TNdEdit2[i].Text:='';
      TNdEdit3[i].Text:='';
      TNdEdit4[i].Text:='';
    end;
  end;
end;


procedure TForm2060.FormShow(Sender: TObject);
var i:integer;
    xId:Real;
    rtf:boolean;
    s:String;
    x:real;
    a,b:Real;
begin
  SetCreate;
  SetColor;
  rtf:=false;

  


  for i:=1 to NDayMes(RMEs,RGod) do
   begin

    xId:=form9.Table1.fieldbyname('TN'+IntToStr(i)).asFLoat;
    if form1.tabel.Locate('id',xId,[locaseinsensitive]) then
      begin
        x:=form1.tabel.fieldbyname('cas').asFloat;

        if ComboBox1.ItemIndex=1 then
         begin
          a:=INT(x) ;
          b:=DRound(100*(x-a),2);
          x:=DRound(a+b*100/60/100,2);
         end;


        if x<>0 then
         begin
          if DRound(x-int(x),2)=0 then s:=floattostr(x) else s:=floattostrf(x,ffNumber,12,2);
          if ComboBox1.ItemIndex=1 then s:=floattostr(x);
          TNdEdit[i].Text:=s;
         end;

        if form1.sptabel2.Locate('ID',form1.tabel.fieldbyname('kod').asFloat,[loCaseInsensitive]) then
            TNdEdit2[i].Text:=form1.sptabel2KOD.Value else TNdEdit2[i].Text:='??' ;
      end;

    xId:=form9.Table1.fieldbyname('TC'+IntToStr(i)).asFLoat;
    if form1.tabel.Locate('id',xId,[locaseinsensitive]) then
      begin
        rtf:=true;
        x:=form1.tabel.fieldbyname('cas').asFloat;

        if ComboBox1.ItemIndex=1 then
         begin
          a:=INT(x) ;
          b:=DRound(100*(x-a),2);
          x:=DRound(a+b*100/60/100,2);
         end;

        if x<>0 then
         begin
          if DRound(x-int(x),2) =0 then s:=floattostr(x) else s:=floattostrf(x,ffNumber,12,2);
          if ComboBox1.ItemIndex=1 then s:=floattostr(x);
          TNdEdit3[i].Text:=s;
         end;

        if form1.sptabel2.Locate('ID',form1.tabel.fieldbyname('kod').asFloat,[loCaseInsensitive]) then
            TNdEdit4[i].Text:=form1.sptabel2KOD.Value else TNdEdit4[i].Text:='??' ;
      end;

   end;

  CheckBox1.Checked:=rtf;
  CheckBox1Click(nil);


end;

procedure TForm2060.CheckBox1Click(Sender: TObject);
var rtf:boolean;
    i:integer;
begin

    rtf:=CheckBox1.Checked;

    jvLabel15.Visible:=rtf;
    jvLabel16.Visible:=rtf;
    jvLabel23.Visible:=rtf;
    jvLabel24.Visible:=rtf;

    for i:=1 to NDayMes(RMEs,RGod) do
      begin
       TNdEdit3[i].visible:=rtf;
       TNdEdit4[i].visible:=rtf;
      end;
end;

procedure TForm2060.Edit1Change(Sender: TObject);
begin
 FProv;
end;

procedure TForm2060.JvXPButton1Click(Sender: TObject);
var i:integer;
    xId,tId:Real;
    DCAS1,DCAS2:array[1..31] of Real;
    DKOD1,DKOD2:array[1..31] of Real;
    s:String;
    a,b:Real;
begin


    datam.Query1.Close;
    datam.Query1.SQL.Clear;
    datam.Query1.SQL.Add('select max(id) from tabel');
    datam.Query1.Prepare;
    datam.Query1.Open;
    tId:=datam.Query1.Fields[0].AsFloat;


 for i:=1 to 31 do DCAS1[i]:=0;
 for i:=1 to 31 do DCAS2[i]:=0;
 for i:=1 to 31 do DKOD1[i]:=0;
 for i:=1 to 31 do DKOD2[i]:=0;


 for i:=1 to NDayMes(RMEs,RGod) do
  begin

   if Trim(TNdEdit2[i].Text)<>'' then
    begin
     if form1.sptabel2.Locate('KOD',Trim(TNdEdit2[i].Text),[loCaseInsensitive]) then DKOD1[i]:=form1.sptabel2ID.Value
       else
        begin
         MessageDlg('Код не из справочника'+#13+Trim(TNdEdit2[i].Text),mtError,[mbOk],0);
         TNdEdit2[i].SetFocus;
         exit;
        end;
    end;

    if Trim(TNdEdit4[i].Text)<>'' then
     begin
     if form1.sptabel2.Locate('KOD',Trim(TNdEdit4[i].Text),[loCaseInsensitive]) then DKOD2[i]:=form1.sptabel2ID.Value
       else
        begin
         MessageDlg('Код не из справочника'+#13+Trim(TNdEdit4[i].Text),mtError,[mbOk],0);
         TNdEdit4[i].SetFocus;
         exit;
        end;
     end;
   try
    s:=trim(TNdEdit[i].Text);
    if s='' then s:='0';
    DCAS1[i]:=StrToFloat(s);
    s:=trim(TNdEdit3[i].Text);
    if s='' then s:='0';
    DCAS2[i]:=StrToFloat(s);
   except
    MessageDlg('Ошибка ввода времени, число '+floattostr(i),mtWarning,[mbOk],0);
    TNdEdit[i].SetFocus;
    exit;
   end;
  end;


 for i:=1 to NDayMes(RMEs,RGod) do
  begin
   if ComboBox1.ItemIndex=1 then
    begin
     a:=INT(DCAS1[i]) ;
     b:=DRound(100*(DCAS1[i]-a),2);
     DCAS1[i]:=DRound(a+b*60/100/100,2);
    end;

   if ComboBox1.ItemIndex=1 then
    begin
     a:=INT(DCAS2[i]) ;
     b:=DRound(100*(DCAS2[i]-a),2);
     DCAS2[i]:=DRound(a+b*60/100/100,2);
    end;

  end;
  

 for i:=1 to NDayMes(RMEs,RGod) do
  begin

     //добавляем если нет ID

    if (form9.Table1.fieldbyname('TN'+IntToStr(i)).asFLoat=0) and (DKOD1[i]<>0) then
     begin
      form1.tabel.Append;
      tId:=tId+1;
      form1.tabel.FieldByNAme('dat').asDateTime:=EncodeDate(RGod,RMEs,i);
      form1.tabel.FieldByNAme('nls').asFloat:=form9.Table1.Fieldbyname('nls').asFloat;
      form1.tabel.FieldByNAme('ID').asFloat:=tId;
      form1.tabel.post;
      form1.tabel.FlushBuffers;
      form9.table1.Edit;
      form9.table1.FieldByName('TN'+IntToStr(i)).asFloat:=form1.tabelID.Value;
      form9.table1.Post;
     end;

    if (form9.Table1.fieldbyname('TC'+IntToStr(i)).asFLoat=0) and (DKOD2[i]<>0) then
     begin
      form1.tabel.Append;
      tId:=tId+1;
      form1.tabel.FieldByNAme('dat').asDateTime:=EncodeDate(RGod,RMEs,i);
      form1.tabel.FieldByNAme('nls').asFloat:=form9.Table1.Fieldbyname('nls').asFloat;
      form1.tabel.FieldByNAme('ID').asFloat:=tId;
      form1.tabel.post;
      form1.tabel.FlushBuffers;
      form9.table1.Edit;
      form9.table1.FieldByName('TC'+IntToStr(i)).asFloat:=form1.tabelID.Value;
      form9.table1.Post;
     end;



    xId:=form9.Table1.fieldbyname('TN'+IntToStr(i)).asFLoat;
    if form1.tabel.Locate('id',xId,[locaseinsensitive]) then
     begin
      if DKOD1[i]<>0 then
       begin
        form1.tabel.edit;
        form1.tabel.fieldbyname('kod').asFloat:=DKOD1[i];
        form1.tabel.fieldbyname('cas').asFloat:=DCAS1[i];
        form1.tabel.post;
       end
        else
       begin
        form1.tabel.delete;
        form9.table1.edit;
        form9.Table1.fieldbyname('TN'+IntToStr(i)).asFLoat:=0;
        form9.table1.post;
       end ;
     end;


    xId:=form9.Table1.fieldbyname('TC'+IntToStr(i)).asFLoat;
    if form1.tabel.Locate('id',xId,[locaseinsensitive]) then
     begin
      if DKOD2[i]<>0 then
       begin
        form1.tabel.edit;
        form1.tabel.fieldbyname('kod').asFloat:=DKOD2[i];
        form1.tabel.fieldbyname('cas').asFloat:=DCAS2[i];
        form1.tabel.post;
       end
        else
       begin
        form1.tabel.delete;
        form9.table1.edit;
        form9.Table1.fieldbyname('TC'+IntToStr(i)).asFLoat:=0;
        form9.table1.post;
       end;
     end;
  end;

  form2060.close;
end;

procedure TForm2060.Edit2Change(Sender: TObject);
var i,k1,k2:integer;
begin
 k1:=0; k2:=0;
 for i:=1 to NDayMes(RMes,RGod) do
   begin
    if Sender=TNdEdit2[i] then k1:=i;
    if Sender=TNdEdit4[i] then k2:=i;
   end;

 if k1<>0 then
  begin
       form_107:=tform_107.create(self);
       form_107.xIDKOD:=0;
       form_107.xFind:=Trim(TNdEdit2[k1].Text);
       form_107.ShowModal;
       if form_107.xIDKOD<>0 then
        begin
         form1.sptabel2.Locate('ID',form_107.xIDKOD,[loCaseInsensitive])  ;
         TNdEdit2[k1].Text:=form1.sptabel2KOD.asString;
        end;
       form_107.free;
  end;

 if k2<>0 then
  begin
       form_107:=tform_107.create(self);
       form_107.xIDKOD:=0;
       form_107.xFind:=Trim(TNdEdit2[k2].Text);
       form_107.ShowModal;
       if form_107.xIDKOD<>0 then
        begin
         form1.sptabel2.Locate('ID',form_107.xIDKOD,[loCaseInsensitive])  ;
         TNdEdit4[k2].Text:=form1.sptabel2KOD.asString;
        end;
       form_107.free;
  end;


end;

procedure TForm2060.JvXPButton3Click(Sender: TObject);
var i:integer;
begin
 if MessageDlg('Удалить все данные за месяц ?',mtInformation,[mbYes,mbNo],0) = mrNo then exit;
 for i:=1 to NDayMes(RMEs,RGod) do
  begin
   TNdEdit2[i].Text:='';
   TNdEdit4[i].Text:='';
   TNdEdit[i].Text:='' ;
   TNdEdit3[i].Text:=''
   
  end;
end;

end.
