unit uved_app;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  JvExControls, JvComponent, JvXPCore, JvXPButtons, StdCtrls, Mask,
  JvExMask, JvToolEdit, JvBaseEdits, DB, mainlib, ExtCtrls, DBTables;

type
  TForm3304 = class(TForm)
    JvXPButton2: TJvXPButton;
    JvXPButton1: TJvXPButton;
    ComboBox1: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    ComboBox2: TComboBox;
    JvCalcEdit1: TJvCalcEdit;
    Label3: TLabel;
    Label4: TLabel;
    JvXPButton35: TJvXPButton;
    Label5: TLabel;
    Edit2: TEdit;
    Label6: TLabel;
    Edit3: TEdit;
    Label7: TLabel;
    Edit4: TEdit;
    JvXPButton9: TJvXPButton;
    Edit5: TEdit;
    Label8: TLabel;
    RadioGroup1: TRadioGroup;
    ComboBox3: TComboBox;
    JvXPButton15: TJvXPButton;
    Edit1: TEdit;
    Label9: TLabel;
    Panel2: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    ComboBox4: TComboBox;
    Edit6: TEdit;
    JvDateEdit1: TJvDateEdit;
    CheckBox1: TCheckBox;
    tmpdbapp: TTable;
    tmpdbappNLS: TFloatField;
    tmpdbappFIO: TStringField;
    tmpdbappDOLJNOST: TStringField;
    tmpdbappBIRTHDAY: TDateField;
    tmpdbappOKTMO: TStringField;
    tmpdbappSUMMA: TFloatField;
    DataSource1: TDataSource;
    JvXPButton3: TJvXPButton;
    tmpdbappX1: TFloatField;
    tmpdbappX2: TFloatField;
    Edit7: TEdit;
    procedure JvXPButton1Click(Sender: TObject);
    procedure JvXPButton35Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure JvXPButton2Click(Sender: TObject);
    procedure SetTMes;
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure JvXPButton9Click(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure JvXPButton15Click(Sender: TObject);
    procedure ComboBox3Change(Sender: TObject);
    procedure ZapolnKBK;
    procedure JvXPButton3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
   TMes:Integer;
   TRG:Integer;
    { Public declarations }
  end;

var
  Form3304: TForm3304;

implementation

uses kbk_spr, pevazp, MyDATAMODULE, uvedfnd, FWait, FSpr2006, oktmo_ed,
  vybperiod, upl_str_vzn, uved_app_ndfl, vyb_srokndfl;

{$R *.DFM}

procedure TForm3304.SetTMes;
var nkvart,nmes:integer;
begin
 if ComboBox1.ItemIndex=0 then nkvart:=1;
 if ComboBox1.ItemIndex=1 then nkvart:=2;
 if ComboBox1.ItemIndex=2 then nkvart:=3;
 if ComboBox1.ItemIndex=3 then nkvart:=4;

 if ComboBox2.ItemIndex=0 then nmes:=1;
 if ComboBox2.ItemIndex=1 then nmes:=2;
 if ComboBox2.ItemIndex=2 then nmes:=3;

 TMes:=(nkvart-1)*3+nMes;

 if (TMes>=1) and (TMes<=12) then Edit5.Text:=ansilowercase(namemes[TMes]) else Edit5.TExt:='??';

 if Radiogroup1.ItemIndex=3 then ZapolnKBK;

end;

procedure TForm3304.JvXPButton1Click(Sender: TObject);
begin
 form3304.close;
end;

procedure TForm3304.JvXPButton35Click(Sender: TObject);
begin
 form845:=tform845.create(nil);
 if RadioGroup1.ItemIndex=0 then form845.TFILTER:=7;
 if RadioGroup1.ItemIndex=1 then form845.TFILTER:=3;
 if RadioGroup1.ItemIndex=2 then form845.TFILTER:=0;
 if RadioGroup1.ItemIndex=3 then form845.TFILTER:=1;
 if RadioGroup1.ItemIndex=4 then form845.TFILTER:=2;
 if RadioGroup1.ItemIndex=5 then form845.TFILTER:=5;

 form845.ShowModal;
 if trim(form845.TKbk)<>trim(Edit2.Text) then
   begin
    if MessageDlg('Изменить значение КБК '+Edit2.Text+#13+'на новое значение '+form845.TKbk,mtInformation,[mbYes,mbNo],0) = mrYes then
      begin
       Edit2.Text:=form845.TKbk;
      end;
   end;
 form845.free;

end;

procedure TForm3304.FormShow(Sender: TObject);
begin

 RadioGroup1.Items[2]:='НДФЛ '+floattostr(RGod)+'г. по сроку уплаты ';
 RadioGroup1.Items[6]:='НДФЛ '+floattostr(RGod-1)+'г.';


 ComboBox1.ItemIndex:=0;
 ComboBox2.ItemIndex:=0;

 if RMes<=3 then
  begin
   ComboBox1.ItemIndex:=0;
   ComboBox2.ItemIndex:=RMes-1;
  end;
 if (RMes>=4) and (RMes<=6) then
  begin
   ComboBox1.ItemIndex:=1;
   ComboBox2.ItemIndex:=RMes-4;
  end;
 if (RMes>=7) and (RMes<=9) then
  begin
   ComboBox1.ItemIndex:=2;
   ComboBox2.ItemIndex:=RMes-7;
  end;
 if RMes>=10 then
  begin
   ComboBox1.ItemIndex:=3;
   ComboBox2.ItemIndex:=RMes-10;   
  end;

 Edit4.Text:=floattostr(RGod);
 RadioGroup1Click(nil);
 SetTMes;

 tmpdbapp.DatabaseName:=form1.dbdir;
 tmpdbapp.TableName:='tmpdbapp.dbf';
 tmpdbapp.Active:=false;
 if tmpdbapp.Exists then tmpdbapp.DeleteTable;
 tmpdbapp.CreateTable;
 tmpdbapp.Active:=true;

 RadioGroup1.ItemIndex:=TRG;

end;

procedure TForm3304.JvXPButton2Click(Sender: TObject);
var xId,x:Real;
    sMes,sPer:String;
    s:String;
    dd,mm,yy:word;
    ik:integer;
begin

 s:='';

 if ((ComboBox4.ItemIndex=3) or (ComboBox4.ItemIndex=6)) and (CheckBox1.Checked) then //ндфл
  begin
    tmpdbapp.first;
    x:=0;
    while not tmpdbapp.eof do
     begin
      x:=x+tmpdbappSUMMA.Value;
      tmpdbapp.next;
     end;

    s:='и в таблицу уплаты НДФЛ на сумму '+FSTR(x,0) +' руб.'+#13+
        'Дата уплаты НДФЛ '+FormatDateTime('dd.mm.yyyy',jvDateEdit1.Date);

    if DRound(jvCalcEdit1.Value-x,2)<>0 then
      begin
       MessageDlg('Сумма по уведомлению '+floattostr(jvCalcEdit1.Value)+#13+'не равна сумме НДФЛ по всем сотрудникам в таблице уплаты '+floattostr(x)+#13
         +'см. кнопка <Инфо>',mtError,[mbOk],0);
       exit;
      end;
  end;

  if MessageDlg('Добавить информацию в Уведомление'+#13+s,mtInformation,[mbYes,mbNo],0) = mrNo then exit;


 datam.qtmp.close;
 datam.qtmp.sql.clear;
 datam.qtmp.databasename:=form1.dbdir;
 datam.qtmp.sql.add('select max(id) from uvedob');
 datam.qtmp.prepare;
 datam.qtmp.open;
 xId:=datam.qtmp.fields[0].asfloat+1;
 datam.qtmp.close;

 sPer:='??';
 if ComboBox1.ItemIndex=0 then sPer:='21';
 if ComboBox1.ItemIndex=1 then sPer:='31';
 if ComboBox1.ItemIndex=2 then sPer:='33';
 if ComboBox1.ItemIndex=3 then sPer:='34';

 sMes:='??';
 if ComboBox2.ItemIndex=0 then sMes:='01';
 if ComboBox2.ItemIndex=1 then sMes:='02';
 if ComboBox2.ItemIndex=2 then sMes:='03';
 if ComboBox2.ItemIndex=3 then sMes:='04';


 if RadioGroup1.ItemIndex=0 then s:='ОПС, ОМС, ВНиМ';
 if RadioGroup1.ItemIndex=1 then s:='Доп.тариф';
 if RadioGroup1.ItemIndex=2 then s:='НДФЛ срок уплаты';
  if RadioGroup1.ItemIndex=3 then s:='ОПС';
   if RadioGroup1.ItemIndex=4 then s:='ОМС';
    if RadioGroup1.ItemIndex=5 then s:='ВНиМ';
      if RadioGroup1.ItemIndex=6 then s:='НДФЛ';

 form3301.uvedob.append;
 form3301.uvedob.fieldbyname('id').asfloat:=xid;
 form3301.uvedob.fieldbyname('iduved').asfloat:=form3301.uvedID.Value;
 form3301.uvedob.fieldbyname('god').asfloat:=RGod;
 form3301.uvedob.fieldbyname('period').asString:=sPer;
 form3301.uvedob.fieldbyname('mes').asString:=sMes;
 form3301.uvedob.fieldbyname('oktmo').asString:=trim(Combobox3.Text);
 form3301.uvedob.fieldbyname('kpp').asString:=trim(Edit3.Text);
 form3301.uvedob.fieldbyname('kbk').asString:=trim(Edit2.Text);
 form3301.uvedob.fieldbyname('note').asString:=s+'/'+trim(Edit1.Text);
 form3301.uvedob.fieldbyname('summa').asFloat:=jvCalcEdit1.Value;
  form3301.uvedob.fieldbyname('dokupl').asString:=trim(Edit6.Text);
   form3301.uvedob.fieldbyname('datupl').asdateTime:=jvDateEdit1.Date;
     form3301.uvedob.fieldbyname('typeupl').asfloat:=ComboBox4.ItemIndex;
 form3301.uvedob.post;

 if (form3301.uvedobTYPEUPL.Value=3) and (CheckBox1.Checked) then //ндфл
  begin

  datam.qtmp.close;
  datam.qtmp.sql.clear;
  datam.qtmp.databasename:=form1.dbdir;
  datam.qtmp.sql.add('select max(id) from uplatandfl');
  datam.qtmp.prepare;
  datam.qtmp.open;
  xId:=datam.qtmp.fields[0].asfloat+1;
  datam.qtmp.close;

   form1.uplatandfl.DatabaseName:=form1.dbdir;
   form1.uplatandfl.Active:=true;
   form3304.tmpdbapp.first;
   while not form3304.tmpdbapp.eof do
    begin
     if Radiogroup1.ItemIndex=2 then
      begin
       xid:=xid+1;
       DecodeDAte(jvDateEdit1.Date,yy,mm,dd);
       form1.uplatandfl.append;
       form1.uplatandfl.fieldbyname('id').asfloat:=xId;
       form1.uplatandfl.fieldbyname('iduved').asfloat:=form3301.uvedob.fieldbyname('id').asfloat;
       form1.uplatandfl.fieldbyname('idst').asfloat:=0;
       form1.uplatandfl.FieldByName('st').asFloat:=13;
       form1.uplatandfl.fieldbyname('nls').asfloat:=form3304.tmpdbappNLS.Value;
       form1.uplatandfl.fieldbyname('summa').asfloat:=form3304.tmpdbappSumma.Value;
       form1.uplatandfl.fieldbyname('oktmo').asstring:=trim(form3304.tmpdbappoktmo.Value);
       form1.uplatandfl.FieldByName('DAT').asDateTime:=jvDateEdit1.Date;
       form1.uplatandfl.fieldbyname('mes').asfloat:=mm;
       form1.uplatandfl.fieldbyname('god').asfloat:=yy;
       form1.uplatandfl.FieldByName('NPLPOR').asString:=trim(edit6.text);
       form1.uplatandfl.post;
      end;

     if Radiogroup1.ItemIndex=6 then
      begin
       FOR IK:=1 TO 2 DO
        BEGIN
         xid:=xid+1;
         DecodeDAte(jvDateEdit1.Date,yy,mm,dd);
         form1.uplatandfl.append;
         form1.uplatandfl.fieldbyname('id').asfloat:=xId;
         form1.uplatandfl.fieldbyname('iduved').asfloat:=form3301.uvedob.fieldbyname('id').asfloat;
         form1.uplatandfl.fieldbyname('idst').asfloat:=0;
         form1.uplatandfl.FieldByName('st').asFloat:=13;
         form1.uplatandfl.fieldbyname('nls').asfloat:=form3304.tmpdbappNLS.Value;
         if ik=1 then form1.uplatandfl.fieldbyname('summa').asfloat:=form3304.tmpdbappx1.Value;
         if ik=2 then form1.uplatandfl.fieldbyname('summa').asfloat:=form3304.tmpdbappx2.Value;
         if ik=1 then form1.uplatandfl.fieldbyname('ifns').asString:='2022';
         form1.uplatandfl.fieldbyname('oktmo').asstring:=trim(form3304.tmpdbappoktmo.Value);
         form1.uplatandfl.FieldByName('DAT').asDateTime:=jvDateEdit1.Date;
         form1.uplatandfl.fieldbyname('mes').asfloat:=mm;
         form1.uplatandfl.fieldbyname('god').asfloat:=yy;
         form1.uplatandfl.FieldByName('NPLPOR').asString:=trim(edit6.text);
         form1.uplatandfl.post;
        END;
      end;


     form3304.tmpdbapp.next;
    end;
    form1.uplatandfl.Active:=false;
  end;



 if ((form3301.uvedobTYPEUPL.Value=1) or (form3301.uvedobTYPEUPL.Value=2)) and (CheckBox1.Checked) then //страховые
  begin
    datam.qtmp.close;
    datam.qtmp.sql.clear;
    datam.qtmp.databasename:=form1.dbdir;
    datam.qtmp.sql.add('select max(id) from uplstr2012');
    datam.qtmp.prepare;
    datam.qtmp.open;
    xId:=datam.qtmp.fields[0].asfloat+1;
    datam.qtmp.close;

    form111.uplstr2012.append;
    form111.uplstr2012.fieldbyname('id').asfloat:=xId;
    form111.uplstr2012.fieldbyname('iduved').asfloat:=form3301.uvedob.fieldbyname('id').asFloat;
    form111.uplstr2012.fieldbyname('idst').asfloat:=0;
    form111.uplstr2012.fieldbyname('pod').asfloat:=0;
    form111.uplstr2012.fieldbyname('summa2').asfloat:=0;
    form111.uplstr2012.fieldbyname('typesum2').asfloat:=0;
    form111.uplstr2012.fieldbyname('summa').asfloat:=form3301.uvedob.fieldbyname('summa').asFloat;
    if form3301.uvedobTYPEUPL.Value=1 then form111.uplstr2012.fieldbyname('type').asfloat:=1;  //страховые
    if form3301.uvedobTYPEUPL.Value=2 then form111.uplstr2012.fieldbyname('type').asfloat:=9;  //доп.тариф
    form111.uplstr2012.fieldbyname('nplpor').asstring:=form3301.uvedob.fieldbyname('dokupl').asString;
    form111.uplstr2012.fieldbyname('dat').asdatetime:=form3301.uvedob.fieldbyname('datupl').asdatetime;
    DecodeDAte(form3301.uvedob.fieldbyname('datupl').asdatetime,yy,mm,dd);
    form111.uplstr2012.fieldbyname('mes').asfloat:=mm;
    form111.uplstr2012.fieldbyname('god').asfloat:=yy;
    form111.uplstr2012.Post;
    MessageDlg('Вместе с Уведомлением добавлена информация в таблицу уплаты страховых взносов',mtinformation,[mbOk],0);
  end;

 form3304.close;

end;

procedure TForm3304.ComboBox1Change(Sender: TObject);
begin
 SetTMes;
end;

procedure TForm3304.ComboBox2Change(Sender: TObject);
begin
 SetTMes;
end;

procedure TForm3304.JvXPButton9Click(Sender: TObject);
var x,x1,x2,x0:Real;
    s,soktmo:String;
    tD1,tD2,tDat:TDate;
    TOK:integer;
    sops,soms,sfss:real;
    gok:boolean;
    gm1,gm2,smes:integer;
    mD1,mD2,mD3:TDate;
    RTF,RTFb:boolean;
    xsumma:Real;
begin

 tmpdbapp.Active:=false;
 if tmpdbapp.Exists then tmpdbapp.DeleteTable;
 tmpdbapp.CreateTable;
 tmpdbapp.Active:=true;


 SetTMes;

 if NOT ((TMes>=1) and (TMes<=12)) then
  begin
   MessageDlg('Ошибка месяца отчета',mtWarning,[mbOk],0);
   exit;
  end;

 s:='??';
 if RadioGroup1.ItemIndex=0 then s:='страховые взносы ОПС, ОМС, ВНиМ';
 if RadioGroup1.ItemIndex=1 then s:='страховые взносы по доп.тарифу';
 if RadioGroup1.ItemIndex=2 then s:='НДФЛ со сроком уплаты '+trim(Edit5.Text)+#13+'ОКТМО='+trim(ComboBox3.Text);
 if RadioGroup1.ItemIndex=6 then s:='НДФЛ остатки '+floattostr(RGod-1)+'г.'+#13+'ОКТМО='+trim(ComboBox3.Text);

  if RadioGroup1.ItemIndex=3 then s:='страховые взносы ОПС';
   if RadioGroup1.ItemIndex=4 then s:='страховые взносы ОМС';
    if RadioGroup1.ItemIndex=5 then s:='страховые взносы ВНиМ';



 if RadioGroup1.ItemIndex=2 then
  begin
   form3305:=tform3305.create(nil);
 //  form3305.tSrok2:=false;
   form3305.showmodal;
   gOk:=form3305.gOk;
   if gOk then
    begin
     mD1:=form3305.tD1;
     mD2:=form3305.tD2;
     mD3:=form3305.tD3;
   //  showmessage(datetostr(mD1)+#13+Datetostr(mD2)+#13+Datetostr(mD3));
    end;
   form3305.free;
   if not gOk then exit;
  end
   else
     if MessageDlg('Рассчитать '+s,mtinformation,[mbYes,mbNo],0) = mrNo then exit;;


  if (RadioGroup1.ItemIndex=0) OR  ((RadioGroup1.ItemIndex>=3) and (RadioGroup1.ItemIndex<=5)) then
       begin
        form1.SpravNDFLStr(false,sops,soms,sfss,smes);
        if RadioGroup1.ItemIndex=0 then x:=sops+soms+sfss;
        if RadioGroup1.ItemIndex=3 then x:=sops;
        if RadioGroup1.ItemIndex=4 then x:=soms;
        if RadioGroup1.ItemIndex=5 then x:=sfss;
        if (smes>=1) and (smes<=12) then edit1.text:=namemes[smes] else edit1.text:='';
        jvCalcEdit1.Value:=DRound(x,2);
        exit;
       end;

       if RadioGroup1.ItemIndex=1 then
        begin
         form58:=TForm58.Create(Self);
         form58.Label2.Visible:=false;
         form58.ComboBox2.Visible:=false;
         form58.ShowModal;
         gm1:=form58.fm1;
         gm2:=form58.fm1;
         gOk:=form58.TOk;
         form58.free;
         if not gOk then exit;
         edit1.text:=namemes[gm1];
        end;




 form1.kart.IndexName:='FAM';
 form1.kart.first;
 x:=0;
 while not form1.kart.eof do
  begin

     if RadioGroup1.ItemIndex=1 then
       begin
        ZapolnEsn2010(form1.kartNLS.Value);
        x:=x+DPFR2013[2,gm1,0]+DPFR2013[4,gm1,0];
       end;


     if RadioGroup1.ItemIndex=2 then   //ндфл
       begin
        edit1.text:=namemes[TMes];
        tD1:=EncodeDate(RGod,1,1);
        tD2:=EncodeDate(RGod,12,31);
        datam.qtmp.close;
        datam.qtmp.databasename:=form1.dbdir;
        datam.qtmp.sql.clear;
        datam.qtmp.sql.add('select * from sdoxod where nls='+floattostr(form1.kartnls.value));
        datam.qtmp.sql.add('and dat>='+#39+FormatDateTime('dd.mm.yyyy',tD1)+#39);
        datam.qtmp.sql.add('and dat<='+#39+FormatDateTime('dd.mm.yyyy',tD2)+#39);
        datam.qtmp.prepare;
        datam.qtmp.open;
        datam.qtmp.first;
        TOK:=Form_58.ZapolnDOK(form1.kartNLS.Value);
        while not datam.qtmp.eof do
         begin
           soktmo:=form_58.FGetOktmo(Trunc(datam.qtmp.fieldbyname('mes').asFloat),Trunc(datam.qtmp.fieldbyname('god').asFloat));
           tDat:=form_58.FGetDatPerecisl(datam.qtmp.fieldbyname('dat').asdatetime);
           RTF:=true;
           if datam.qtmp.FieldByName('god').asfloat<=2022 then
             begin
               RTF:=false;

              {
               datam.qtmpstaj.close;
               datam.qtmpstaj.DatabaseName:=form1.dbdir;
               datam.qtmpstaj.sql.clear;
               datam.qtmpstaj.sql.add('select datprov from obrt2new where nls='+floattostr(form1.kartnls.value)+' and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
               datam.qtmpstaj.prepare;
               datam.qtmpstaj.open;
               if datam.qtmpstaj.recordcount>=1 then
                begin
                 if datam.qtmpstaj.fieldbyname('datprov').asdatetime<EncodeDate(RGod,1,1) then RTF:=false;
                end;
               datam.qtmpstaj.close;
              }
             end;
           if (tDAt=mD3) and (trim(soktmo)=trim(ComboBox3.Text)) and (RTF) then
             begin
              x:=x+datam.qtmp.fieldbyname('nalog').asfloat;
              if not tmpdbapp.Locate('nls',form1.kartnls.value,[locaseinsensitive]) then
                begin
                 tmpdbapp.Append;
                 tmpdbapp.fieldbyname('nls').asfloat:=form1.kartnls.value;
                 tmpdbapp.fieldbyname('fio').asstring:=form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.';
                 tmpdbapp.fieldbyname('doljnost').asstring:=form1.kartDOLGNOST.Value;
                 tmpdbapp.fieldbyname('oktmo').asstring:=trim(soktmo);
                 tmpdbapp.fieldbyname('birthday').asdatetime:=form1.kartBIRTHDAY.Value;
                 tmpdbapp.fieldbyname('summa').asfloat:=datam.qtmp.fieldbyname('nalog').asfloat;
                 tmpdbapp.post;
                end
                  else
                 begin
                  tmpdbapp.edit;
                  tmpdbapp.fieldbyname('summa').asfloat:=tmpdbapp.fieldbyname('summa').asfloat+datam.qtmp.fieldbyname('nalog').asfloat;
                  tmpdbapp.post;
                 end;

             end;
          datam.qtmp.next;
         end;
        
       end;


     if RadioGroup1.ItemIndex=6 then   //ндфл остатки
       begin
        edit1.text:=floattostr(RGod-1)+'г.';
        x:=0;
        form1.kart.first;
        while not form1.kart.eof do
         begin
           TOK:=Form_58.ZapolnDOK(form1.kartNLS.Value);
           soktmo:=form_58.FGetOktmo(1,RGod);
           if trim(soktmo)=trim(ComboBox3.Text) then
            begin
               xsumma:=0;
               datam.qtmpstaj.close;
               datam.qtmpstaj.DatabaseName:=form1.DBDIR;
               datam.qtmpstaj.SQL.Clear;
               datam.qtmpstaj.sql.add('select summa from uplatandfl where nls='+floattostr(form1.kartnls.value));
               datam.qtmpstaj.sql.add('and mes=0 and god='+floattostr(RGod));
               datam.qtmpstaj.Prepare;
               datam.qtmpstaj.open;
               if datam.qtmpstaj.RecordCount>=1 then xsumma:=datam.qtmpstaj.fieldbyname('summa').asfloat;
               datam.qtmpstaj.close;


                //*****
                 datam.qtmp.close;
                 datam.qtmp.databasename:=form1.dbdir;          //выплата 2022 года в 2023 году разбивка остатка
                 datam.qtmp.sql.clear;
                 datam.qtmp.sql.add('select * from sdoxod where nls='+floattostr(form1.kartnls.value));
                 datam.qtmp.sql.add('and dat>='+#39+FormatDateTime('dd.mm.yyyy',EncodeDate(RGod,1,1))+#39);
                 datam.qtmp.sql.add('and dat<='+#39+FormatDateTime('dd.mm.yyyy',EncodeDate(RGod,12,31))+#39);
                 datam.qtmp.sql.add('and GOD<'+floattostr(RGod));
                 datam.qtmp.prepare;
                 datam.qtmp.open;
                 datam.qtmp.first;
                 x1:=0; x2:=0;
                 while not datam.qtmp.eof do  //сколько доход выплачен в 2023 году
                   begin
                    datam.qtmpstaj.close;
                    datam.qtmpstaj.DatabaseName:=form1.dbdir;
                    datam.qtmpstaj.sql.clear;
                    datam.qtmpstaj.sql.add('select datprov from obrt2new where nls='+floattostr(form1.kartnls.value)+' and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
                    datam.qtmpstaj.prepare;
                    datam.qtmpstaj.open;
                    RTFb:=true;
                     if datam.qtmpstaj.recordcount>=1 then
                      begin
                       if datam.qtmpstaj.fieldbyname('datprov').asdatetime<EncodeDate(RGod,1,1) then RTFb:=false;
                      end;
                    datam.qtmpstaj.close;

                    if RTFb then x2:=x2+datam.qtmp.fieldbyname('nalog').asfloat;


                  datam.qtmp.next;
                end;
       //*****

             x0:=xsumma; //x - всего остаток 2022г, x2 - доход ндфл выплачен в 2023 за 2022, остальное x1 просто остаток 2022
      //      ShowMessage(floattostr(x1)+#13+floattostr(x2)+#13+floattostr(xsumma)+#13+floattostr(x0)) ;

             if x0>=x2 then x1:=DRound(x0-x2,2) else
              begin
               x1:=0;
               x2:=x0;
              end;
    //        ShowMessage(floattostr(x1)+#13+floattostr(x2)+#13+floattostr(xsumma)+#13+floattostr(x0)) ;


               if xsumma<>0 then
                begin
                 tmpdbapp.Append;
                 tmpdbapp.fieldbyname('nls').asfloat:=form1.kartnls.value;
                 tmpdbapp.fieldbyname('fio').asstring:=form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.';
                 tmpdbapp.fieldbyname('doljnost').asstring:=form1.kartDOLGNOST.Value;
                 tmpdbapp.fieldbyname('oktmo').asstring:=trim(soktmo);
                 tmpdbapp.fieldbyname('birthday').asdatetime:=form1.kartBIRTHDAY.Value;
                 tmpdbapp.fieldbyname('summa').asfloat:=x1+x2;
                 tmpdbapp.fieldbyname('x1').asfloat:=x1;
                 tmpdbapp.fieldbyname('x2').asfloat:=x2;
                 x:=x+x1+x2;
                 tmpdbapp.post;
                end;
             end;
          form1.kart.next;
         end;
       end;


   form1.kart.next;
  end;

  jvCalcEdit1.Value:=DRound(x,2);

  if RadioGroup1.ItemIndex=2 then
   begin
     Edit1.Text:=formatdatetime('dd.mm.yyyy',mD3);
     MessageDlg('Сформировано по сроку уплаты НДФЛ '+formatdatetime('dd.mm.yyyy',mD3)+#13+
       'Выплата дохода период '+Formatdatetime('dd.mm.yyyy',mD1)+' - '+Formatdatetime('dd.mm.yyyy',mD2) + #13+
                    'Список сотрудников можно посмотреть и отредактировать через кнопку ИНФО',mtInformation,[mbok],0);
   end;

  if RadioGroup1.ItemIndex=6 then
   begin
    if x<>0 then
      begin
       MessageDlg('Сформировано уплата остатков НДФЛ за предыдущий год'+ #13+
        'Список сотрудников можно посмотреть и отредактировать через кнопку ИНФО',mtInformation,[mbok],0);
      end
       else
         MessageDlg('Сумма остатков на начало года = 0 рублей'+#13+'В случае наличия остатков в декабре предыдущего года выполнить <Перенос остатков> в блоке <Перечисление НДФЛ>',mtInformation,[mbOk],0);

   end;

end;

procedure TForm3304.RadioGroup1Click(Sender: TObject);
begin
 edit1.text:='';

 form1.kbk.filter:='';
 form1.kbk.filtered:=false;
 Combobox3.Items.Clear;
 Edit3.Text:='';
 jvXPButton3.Visible:=false;
 Label9.Caption:='Период' ;

 Edit7.Text:=RadioGroup1.Items[RadioGroup1.ItemIndex];


 if Radiogroup1.ItemIndex=0 then
  begin
    ComboBox4.ItemIndex:=1;
    JvXPButton15.Visible:=false;
    ComboBox3.Items.Add(form1.config2OKTMO.Value);
    Combobox3.ItemIndex:=0;
    Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([7,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;
  if Radiogroup1.ItemIndex=1 then
   begin
     ComboBox4.ItemIndex:=2;
     JvXPButton15.Visible:=false;
     ComboBox3.Items.Add(form1.config2OKTMO.Value);
     Combobox3.ItemIndex:=0;
     Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([3,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;
  if (Radiogroup1.ItemIndex=2) or (Radiogroup1.ItemIndex=6) then
   begin
     if Radiogroup1.ItemIndex=2 then Label9.Caption:='Срок уплаты';
     if Radiogroup1.ItemIndex=6 then Label9.Caption:='Уплата';
     ComboBox4.ItemIndex:=3;
     jvXPButton3.Visible:=true;
     JvXPButton15.Visible:=true;
     ZapolnKBK;
     Combobox3.ItemIndex:=0;
     Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([0,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;

 if Radiogroup1.ItemIndex=3 then
  begin
    ComboBox4.ItemIndex:=1;
    JvXPButton15.Visible:=false;
    ComboBox3.Items.Add(form1.config2OKTMO.Value);
    Combobox3.ItemIndex:=0;
    Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([1,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;

 if Radiogroup1.ItemIndex=4 then
  begin
    ComboBox4.ItemIndex:=1;
    JvXPButton15.Visible:=false;
    ComboBox3.Items.Add(form1.config2OKTMO.Value);
    Combobox3.ItemIndex:=0;
    Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([2,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;

 if Radiogroup1.ItemIndex=5 then
  begin
    ComboBox4.ItemIndex:=1;
    JvXPButton15.Visible:=false;
    ComboBox3.Items.Add(form1.config2OKTMO.Value);
    Combobox3.ItemIndex:=0;
    Edit3.Text:=form1.config2KPP.Value;
    if form1.kbk.Locate('id;god',VarArrayOf([5,RGod]),[loCaseInsensitive]) then Edit2.Text:=form1.kbkKBK.Value else Edit2.Text:='?';
  end;





end;

procedure TForm3304.ZapolnKBK;
var TOK,i:Integer;
    rtf:boolean;
    soktmo:string;
begin
 Combobox3.Items.Clear;
 Combobox3.Items.Add(form1.config2OKTMO.Value);
 form1.kart.first;
 while not form1.kart.eof do
  begin
   datam.kart2.locate('nls',form1.kartnls.value,[locaseinsensitive]);

    if trim(datam.kart2OKTMO.Value)<>'' then
        begin
         rtf:=false;
         for i:=0 to ComboBox3.Items.Count do
          begin
           if ComboBox3.Items[i]=trim(datam.kart2OKTMO.Value) then rtf:=true;
          end;
         if not rtf then CombObox3.Items.Add(trim(datam.kart2OKTMO.Value));
        end;
   TOK:=Form_58.ZapolnDOK(form1.kartNLS.Value);  //=0 - нет записей в таблице ОКТМО один только ОКТМО, =1 - есть
   if TOK<>0 then
    begin
     soktmo:=form_58.FGetOktmo(TMes,RGod);
      rtf:=false;
         for i:=0 to ComboBox3.Items.Count do
          begin
           if ComboBox3.Items[i]=trim(soktmo) then rtf:=true;
          end;
         if not rtf then CombObox3.Items.Add(trim(soktmo));
    end;
   form1.kart.next;
  end;

  if ComboBox3.Items.Count>1 then MessageDlg('В данном отчетном периоде более одного ОКТМО по сотрудникам',mtinformation,[mbOk],0);

end;

procedure TForm3304.JvXPButton15Click(Sender: TObject);
begin

 if not form_58.oktmo.Locate('oktmo',trim(ComboBox3.Text),[loCaseInsensitive]) then
  begin
   form_58.oktmo.append;
   form_58.oktmo.fieldbyname('oktmo').asString:=trim(ComboBox3.Text);
   form_58.oktmo.post;
  end;

 form807:=Tform807.Create(nil);
 form807.ShowModal;
 form807.free;
end;

procedure TForm3304.ComboBox3Change(Sender: TObject);
begin
  if RadioGroup1.ItemIndex=2 then
    begin
     if form_58.oktmo.Locate('oktmo',trim(Combobox3.Text),[loCaseInsensitive]) then Edit3.Text:=form_58.oktmo.fieldbyname('kpp').asString;
    end;
end;

procedure TForm3304.JvXPButton3Click(Sender: TObject);
var x:real;
begin
 form3328:=tform3328.create(nil);
 form3328.showmodal;
 form3328.free;

 tmpdbapp.first;
 x:=0;
 while not tmpdbapp.eof do
  begin
    x:=x+tmpdbappSUMMA.Value;
   tmpdbapp.next;
  end;
  jvCalcEdit1.Value:=DRound(x,2);
end;

procedure TForm3304.FormCreate(Sender: TObject);
begin
 TRG:=0;
end;

end.
