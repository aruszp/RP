unit FSpr2006;
                        
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  OffBtn, RXCtrls, StdCtrls, Mask, ToolEdit, CurrEdit, Buttons,
  JvExControls, JvComponent, JvXPCore, JvXPButtons,DB,mainlib,ComObj, nalkarta,
  DBTables, Grids, DBGrids,Reindex, ExtCtrls, Filectrl,ShellApi, Menus,
  JvGradient,excellib;

type
  TForm_58 = class(TForm)
    JvXPButton1: TJvXPButton;
    JvXPButton5: TJvXPButton;
    podpndfl: TTable;
    podpndflFIO: TStringField;
    podpndflDOKUM: TStringField;
    podpndflAGENT: TFloatField;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    JvXPButton6: TJvXPButton;
    JvXPButton7: TJvXPButton;
    tmprepdox: TTable;
    DataSource2: TDataSource;
    DBGrid2: TDBGrid;
    Button1: TButton;
    Panel1: TPanel;
    Edit2: TEdit;
    ComboBox2: TComboBox;
    CheckBox1: TCheckBox;
    ComboBox3: TComboBox;
    JvXPButton8: TJvXPButton;
    DateEdit2: TDateEdit;
    ComboBox4: TComboBox;
    Panel2: TPanel;
    DateEdit1: TDateEdit;
    ComboBox1: TComboBox;
    Edit1: TEdit;
    RxCalcEdit1: TRxCalcEdit;
    JvXPButton3: TJvXPButton;
    JvXPButton2: TJvXPButton;
    JvXPButton4: TJvXPButton;
    Edit3: TEdit;
    JvXPButton10: TJvXPButton;
    oktmo: TTable;
    Edit4: TEdit;
    Edit5: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    oktmooktmo: TStringField;
    oktmoifns: TStringField;
    Edit6: TEdit;
    k6ndfl: TTable;
    k6ndflgod: TFloatField;
    k6ndflmes: TFloatField;
    k6ndflkolvo: TFloatField;
    k6ndflsumma: TFloatField;
    k6ndflndfl: TFloatField;
    k6ndfloktmo: TStringField;
    k6ndfldat: TDateField;
    JvXPButton9: TJvXPButton;
    DataSource3: TDataSource;
    ndflr5: TTable;
    ndflr5nls: TFloatField;
    ndflr5god: TFloatField;
    ndflr5suderj: TFloatField;
    ndflr5sfix: TFloatField;
    ndflr5snuderj: TFloatField;
    ndflr5dat: TDateField;
    ndflr5num: TStringField;
    ndflr5ifns: TStringField;
    ndflr5sisc: TFloatField;
    ndflr5sud: TFloatField;
    ndfl6: TTable;
    ndfl6ob: TTable;
    ndfl6ID: TFloatField;
    ndfl6GOD: TFloatField;
    ndfl6MES: TFloatField;
    ndfl6OKTMO: TStringField;
    ndfl6P020: TFloatField;
    ndfl6P025: TFloatField;
    ndfl6P030: TFloatField;
    ndfl6P040: TFloatField;
    ndfl6P045: TFloatField;
    ndfl6P050: TFloatField;
    ndfl6P0202: TFloatField;
    ndfl6P0252: TFloatField;
    ndfl6P0302: TFloatField;
    ndfl6P0402: TFloatField;
    ndfl6P0452: TFloatField;
    ndfl6P0502: TFloatField;
    ndfl6P060: TFloatField;
    ndfl6P070: TFloatField;
    ndfl6P080: TFloatField;
    ndfl6P090: TFloatField;
    ndfl6obID: TFloatField;
    ndfl6obID2: TFloatField;
    ndfl6obSUMMA: TFloatField;
    ndfl6obNDFL: TFloatField;
    ndfl6obDAT1: TDateField;
    ndfl6obDAT2: TDateField;
    ndfl6obDAT3: TDateField;
    JvXPButton12: TJvXPButton;
    RadioGroup1: TRadioGroup;
    JvXPButton13: TJvXPButton;
    ndfl6P0203: TFloatField;
    ndfl6P0253: TFloatField;
    ndfl6P0303: TFloatField;
    ndfl6P0403: TFloatField;
    ndfl6P0453: TFloatField;
    ndfl6P0503: TFloatField;
    oktmoKPP: TStringField;
    Edit7: TEdit;
    Edit8: TEdit;
    JvXPButton14: TJvXPButton;
    JvXPButton15: TJvXPButton;
    JvXPButton16: TJvXPButton;
    oktmoNOTE: TStringField;
    PopupMenu2: TPopupMenu;
    N20171720181: TMenuItem;
    N20171: TMenuItem;
    reorg: TTable;
    reorgkod: TStringField;
    reorginn: TStringField;
    reorgKPP: TStringField;
    Edit99: TEdit;
    JvXPButton17: TJvXPButton;
    reorgpriznak: TStringField;
    JvXPButton18: TJvXPButton;
    Edit10: TEdit;
    Edit9: TEdit;
    JvXPButton19: TJvXPButton;
    PopupMenu3: TPopupMenu;
    N21711566021020181: TMenuItem;
    N5711566021020181: TMenuItem;
    JvGradient1: TJvGradient;
    RxLabel6: TRxLabel;
    Label1: TLabel;
    JvGradient2: TJvGradient;
    RxLabel3: TRxLabel;
    RxLabel2: TRxLabel;
    RxLabel19: TRxLabel;
    RxLabel1: TRxLabel;
    RxLabel4: TRxLabel;
    RxLabel13: TRxLabel;
    RxLabel16: TRxLabel;
    RxLabel14: TRxLabel;
    RxLabel17: TRxLabel;
    JvGradient3: TJvGradient;
    RxLabel15: TRxLabel;
    RxLabel10: TRxLabel;
    RxLabel9: TRxLabel;
    RxLabel18: TRxLabel;
    RxLabel8: TRxLabel;
    RxLabel7: TRxLabel;
    RxLabel5: TRxLabel;
    RxLabel20: TRxLabel;
    ndfl6P0602: TFloatField;
    ndfl6P0603: TFloatField;
    ndfl6P0702: TFloatField;
    ndfl6P0703: TFloatField;
    ndfl6P0802: TFloatField;
    ndfl6P0803: TFloatField;
    ndfl6P0902: TFloatField;
    ndfl6P0903: TFloatField;
    ndfl6P113: TFloatField;
    ndfl6P1132: TFloatField;
    ndfl6P1133: TFloatField;
    ndfl6KBK: TStringField;
    JvXPButton11: TJvXPButton;
    ndfl6P180: TFloatField;
    ndfl6P1802: TFloatField;
    ndfl6P1803: TFloatField;
    tItog: TTable;
    tItogndfl: TFloatField;
    tItogdat3: TDateField;
    ndflr5DOXNEUD: TFloatField;
    ndflr5PR1: TFloatField;
    ndflr5PR2: TFloatField;
    ndfl6P1151: TFloatField;
    ndfl6P1152: TFloatField;
    ndfl6P1153: TFloatField;
    ndfl6P1211: TFloatField;
    ndfl6P1212: TFloatField;
    ndfl6P1213: TFloatField;
    ndfl6P1421: TFloatField;
    ndfl6P1422: TFloatField;
    ndfl6P1423: TFloatField;
    ndfl6P1551: TFloatField;
    ndfl6P1552: TFloatField;
    ndfl6P1553: TFloatField;
    ndflr5SPRIB: TFloatField;
    JvXPButton20: TJvXPButton;
    JvXPButton21: TJvXPButton;
    ndfl6DEC22: TFloatField;
    Table1: TTable;
    N20231: TMenuItem;
    Label2: TLabel;
    RxLabel11: TRxLabel;
    RxLabel12: TRxLabel;
    procedure FormActivate(Sender: TObject);
    procedure JvXPButton1Click(Sender: TObject);
    procedure JvXPButton2Click(Sender: TObject);
    procedure JvXPButton3Click(Sender: TObject);
    procedure V(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure JvXPButton5Click(Sender: TObject);
    procedure JvXPButton6Click(Sender: TObject);
    procedure JvXPButton7Click(Sender: TObject);
    procedure JvXPButton8Click(Sender: TObject);
    procedure JvXPButton10Click(Sender: TObject);
    function  FPrazdnik(xDat:TDate):TDate ;
      function  FPrazdnik2(xDat:TDate):TDate ;
         function FPrazdnikEndGod(xDat:TDate):TDate ;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction); //первый рабочий день после даты xDat
    function FSetGni(xoktmo:string):String;
    function FSetKPP(xoktmo:string):String;

    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox4Change(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure PSpr2016(wnls:Real);
    procedure PSpr2017(wnls:Real);
    procedure PSpr2016_st35(wnls:Real);
    procedure JvXPButton9Click(Sender: TObject);
    procedure JvXPButton11Click(Sender: TObject);
    procedure FDatNacisl(tkoddox:String;wm,wg:Integer;datvyp:TDate; var datnacisl,datuderj,datperecisl:TDate);
    procedure JvXPButton12Click(Sender: TObject);
    procedure FormRepDatNew(formtype,tnevyp:integer) ;
    procedure JvXPButton13Click(Sender: TObject);
    procedure JvXPButton14Click(Sender: TObject);
    procedure JvXPButton15Click(Sender: TObject);

    procedure FormRepSrok(tTypeDat:Integer); //по сроку уплаты НДФЛ
      procedure FormRepSrok22(tExcel:Boolean;tTypeDat:Integer); //по сроку уплаты НДФЛ

    function FGetUderjNdfl(tNls:Real;tDatEnd:TDate):Real;
    procedure FNdfl6raspred(tV:Boolean);
    procedure DelIdVyplNull(tNls:Real);
    function SmenaKodAvans(mes,god,nls:Real):Boolean;
        function SmenaKodAvansFil(mes,god,nls:Real):Boolean;     //с учетом подразделения
    procedure JvXPButton16Click(Sender: TObject);
    procedure N20171Click(Sender: TObject);
    procedure N20171720181Click(Sender: TObject);
    procedure JvXPButton4Click(Sender: TObject);
    procedure JvXPButton17Click(Sender: TObject);
    function FKodNdfl6():String;
    procedure ComboBox3Change(Sender: TObject);
    procedure JvXPButton18Click(Sender: TObject);
    procedure JvXPButton19Click(Sender: TObject);
    procedure PSpr2019(wnls:Real);
    procedure PSpr20192(wnls:Real;tFakt:boolean);
    procedure RepVypl(pMes,pGod:integer);
    procedure N21711566021020181Click(Sender: TObject);
    procedure N5711566021020181Click(Sender: TObject);

    procedure P3Vyvod(E2:oleVariant;klist:integer;Znacenie:String;nStroka:Integer;StartCol,EndCol:string)      ;  //вывод через ячейку

    procedure IsprDat31122022;

    procedure Ispr2NDFLDec2022 ;

    procedure PVyvod(E:oleVariant;nList:Integer;Adr:String;Znacenie:String;NeVyvod:String);

    function ZapolnDOK(tNls:Real):Integer;

    function FOktmo(wm,wg:integer;sOktmo:String):boolean;

     function FGetOktmo(wm,wg:integer):String;
      function FGetDatPerecisl(datvyp:TDate):TDate;
    procedure RxLabel10Click(Sender: TObject); // да/нет октмо в данном периоде предварительно ZapolnDOK
    procedure ProcIsprOKTMO;

    procedure ZapolnSpr2021(FType:Integer;sKpp:String);
    procedure JvXPButton20Click(Sender: TObject);
    procedure JvXPButton21Click(Sender: TObject);

    procedure NewSpr2023;
    procedure N20231Click(Sender: TObject);


  private
    { Private declarations }

  public
   
    RKOD,RINN,RKPP,RPRIZNAK:String;
    DOK2:array[1..10] of TDate;
    SOK2:array[1..10] of String;
    NUMLIST:Integer;
   
    SKBK:String;
     xMes1,xMes2,xMes21:Integer;
       PRNUMDAT:Boolean;
   FXML:TextFile;
   EREESTR:Variant;
   itdoxod,itisc,ituderj,itupl,itfix,itlis,itneud,itprib:Real;
   itkolvo:integer;

    tIsprDec2022:boolean;

    { Public declarations }
  end;

var
  Form_58: TForm_58;

implementation

uses pevazp, inf_basa, kartrab, password, MyDATAMODULE, kladr, podpndfled,
  status, FWait, Fspr2006_6Ndfl, K6_ndfl, FSpr2006_r5, Fspr2006_v6,
  vybperiod, Ed_stroka, oktmo_ed, uplatadoxod, kartobrt, Soobsenie, reorg,
  NPVed, vvodnpackapfr, vvod_npved, memoedit, vyb_srokndfl, ObrabDec2022,
  vyb_d1_d2, uplatandfl, ndfl6_report;

{$R *.DFM}


procedure TForm_58.NewSpr2023; //перерасчпределяет коды доходов и вычетов для Справок 2-НДФЛ с 2023 года
var i,j,k,ik:Integer;
    x,xv,xn,xtmp:real;
    dd,mm,yy:Word;
    sSoob:String;
    sKoddox:String;
    Brtf,rtf:boolean;
    st:real;
    xMes,xGod:Integer;
    DVicet:array[1..12,1..2] of Real;
    _DST20:array[1..6] of Real;
    pvic:real;
    skodv:String;
    TOK:integer;
    DKorr:array[1..12] of Real;
begin
 for i:=1 to 6 do _DST20[i]:=0;
 for i:=1 to 13 do for j:=1 to 10 do qqMT[j,i]:=0;
 for i:=1 to 12 do begin DDoxod9[i]:=0; DPn9[i]:=0 end;
 DIsc[13]:=0; DNal[13]:=0;

 TOK:=form_58.ZapolnDOK(form1.kartNls.Value);   //изменение ОКТМО

 for i:=1 to 12 do DKorr[i]:=0;  //корректировка вычетов
 datam.qtmpstaj.close;
 datam.qtmpstaj.DatabaseName:=form1.dbdir;
 datam.qtmpstaj.sql.clear;
 datam.qtmpstaj.sql.add('select * from zakrmes where nls='+floattostr(form1.kartnls.value));
 datam.qtmpstaj.sql.add('and type=95');
 datam.qtmpstaj.sql.add('and wg='+floattostr(RGod));
 datam.qtmpstaj.prepare;
 datam.qtmpstaj.open;
 if datam.qtmpstaj.RecordCount>=1 then
  begin
   for i:=1 to 12 do DKorr[i]:=datam.qtmpstaj.fieldbyname('n'+floattostr(i)).asfloat;
  end;

 datam.qtmpstaj.close;
 datam.qtmpstaj.DatabaseName:=form1.dbdir;
 datam.qtmpstaj.sql.clear;
 datam.qtmpstaj.sql.add('select * from sdoxod where nls='+floattostr(form1.kartnls.value));
 datam.qtmpstaj.sql.add('and dat>='+#39+FormatdateTime('dd.mm.yyyy',EncodeDate(RGod,1,1))+#39);
 datam.qtmpstaj.sql.add('and dat<='+#39+FormatdateTime('dd.mm.yyyy',EncodeDate(RGod,12,31))+#39);
 datam.qtmpstaj.prepare;
 datam.qtmpstaj.open;
 datam.qtmpstaj.first;
 for i:=1 to 12 do for j:=1 to 2 do DVicet[i,j]:=0;
 while not datam.qtmpstaj.eof do
  begin
   xMes:=datam.qtmpstaj.fieldbyname('mes').asinteger;
   xGod:=datam.qtmpstaj.fieldbyname('god').asinteger;

   Brtf:=true;
   DecodeDate(datam.qtmpstaj.fieldbyname('dat').asdatetime,yy,mm,dd);

   s:=form_58.FGetOktmo(mm,yy);
   if trim(s)<>trim(form_58.ComboBox1.Text) then Brtf:=false;


   if (datam.qtmpstaj.fieldbyname('god').asfloat<=2022) and (datam.qtmpstaj.fieldbyname('tavans').asfloat=1) then Brtf:=false; //аванс 2022г декабрь
   if datam.qtmpstaj.fieldbyname('kodnac').asfloat<>0 then
    begin
     form1.NACISL.locate('kod',datam.qtmpstaj.fieldbyname('kodnac').asfloat,[locaseinsensitive]);
     sKoddox:=form1.NACISLKODDOX.Value;
     if form1.NACISLpn.Value=1 then Brtf:=false;
    end
     else
    begin
     sKoddox:='2000'; 
    end;
    k:=0;
    if Brtf then
     begin
      rtf:=false;
      for i:=1 to 10 do
       begin
        if not rtf then
         begin
          if sKoddox=qqD[i] then begin k:=i; rtf:=true; end;
          if (qqD[i]='') and (not rtf) then begin k:=i; qqD[i]:=sKoddox; rtf:=true; end;
         end;
       end;
        if k<>0 then
          begin
           qqMT[k,mm]:=qqMt[k,mm]+datam.qtmpstaj.fieldbyname('sdoxod').asfloat;
           DIsc[13]:=DIsc[13]+datam.qtmpstaj.fieldbyname('nalog').asfloat;
           DNal[13]:=DNal[13]+datam.qtmpstaj.fieldbyname('sdoxod').asfloat-datam.qtmpstaj.fieldbyname('rvicet').asfloat;
           if xGod=RGod then DVicet[xMes,2]:=DVicet[xMes,2]+datam.qtmpstaj.fieldbyname('rvicet').asfloat
                           else DVicet[xMes,1]:=DVicet[xMes,1]+datam.qtmpstaj.fieldbyname('rvicet').asfloat;;

          end
           else
             ShowMessage(form1.kartfam.value+#13+'Ошибка распределения по кодам дохода');
     end;

   datam.qtmpstaj.next;
  end;

  for i:=1 to 12 do
   begin

    x:=DVicet[i,1]; //по sdoxod вычеты месяца i распределить предыдущий год
    if Dround(x,2)<>0 then
     begin
      RGod:=RGod-1;
      for ik:=1 to 3 do
        begin
          if ik=1 then begin pvic:=FVic1(i); skodv:=form1.sTmpString; end;
          if ik=2 then pvic:=FKartVic2(i,skodv) ;
          if ik=3 then pvic:=FIjdev2(i,skodv);

           rtf:=false;
           for j:=1 to 6 do
            begin
             if not rtf then
              begin
               if sKodv=DS20[j] then begin k:=j; rtf:=true; end;
               if (DS20[j]='') and (not rtf) then begin k:=j; DS20[j]:=sKodv; rtf:=true; end;
              end;
             end;
           if x>=pvic then
             begin
              _DST20[k]:=_DST20[k]+pvic;
              x:=x-pvic;
             end
              else
             begin
              _DST20[k]:=_DST20[k]+x;
              x:=0;
             end ;
        end;
      RGod:=RGod+1;
     end;

    x:=DVicet[i,2]; //по sdoxod вычеты месяца i распределить текущий год !
   // x - фактические вычеты в месяце i уже с учетом корректировки даже
   for j:=1 to 6 do                       //
    begin
     if (DKorr[i]<>0) and (DS20[j]<>'') then
       begin
        DST20[j,i]:=DST20[j,i]+DKorr[i];  DKorr[i]:=0;
       end ;
    end;


    for j:=1 to 6 do
     begin
      if (DST20[j,i]<>0) and (DS20[j]<>'') then //фактически применены в месяце i
       begin
      //  ShowMessage(floattostr(DST20[j,i])+#13+floattostr(i)+' '+floattostr(j)+' '+DS20[j]);
        if x>=DST20[j,i] then
          begin
            _DST20[j]:=_DST20[j]+DST20[j,i]; x:=x-DST20[j,i]
          end
           else
          begin
           _DST20[j]:=_DST20[j]+x; x:=0;
          end;

       end;
     end;
   end;
  xv:=0;
  for j:=1 to 6 do
   begin
    xv:=xv+_DST20[j];
    DST20[j,13]:=_DST20[j];
   end;

 for j:=1 to 10 do   //перенос вычетов по мат.помощи если в след.месяце выплата но работает толкьо внутри года и если одна выплата !!!!
  begin
   if (qqD[j]<>'') and (qqV[j]<>'') then
    begin
      datam.qtmpstaj.first;
      while not datam.qtmpstaj.eof do
       begin
          if datam.qtmpstaj.fieldbyname('kodnac').asfloat<>0 then
           begin
            form1.NACISL.locate('kod',datam.qtmpstaj.fieldbyname('kodnac').asfloat,[locaseinsensitive]);
            sKoddox:=form1.NACISLKODDOX.Value;
            if sKoddox=qqD[j] then
             begin
              DecodeDate(datam.qtmpstaj.fieldbyname('dat').asdatetime,yy,mm,dd);
              xMes:=datam.qtmpstaj.fieldbyname('mes').asinteger;
              xGod:=datam.qtmpstaj.fieldbyname('god').asinteger;
              if (xGod=yy) and (mm<>xmes) then
               begin
                if (qqDV[j,xmes]<>0) and (DRound(qqDV[j,mm],2)=0) then
                  begin
                   qqDV[j,mm]:=qqDV[j,xMes];
                   qqDV[j,xmes]:=0;
                  end;
               end;
             end;
           end;
        datam.qtmpstaj.next;
       end;
    end;
  end;

  if trim(form1.kartIMVYC_KOD.Value)<>'' then for i:=1 to 12 do
     begin
      xv:=xv+DImvyc[i];
     end;
  xn:=0;
  xtmp:=0;
  for j:=1 to 10 do for i:=1 to 12 do
   begin
    xn:=xn+qqMT[j,i];
    qqMT[j,13]:=qqMT[j,13]+qqMT[j,i];
    if (qqV[j]<>'') and (qqDV[j,i]<>0) then xtmp:=xtmp+qqDV[j,i]; //вычеты мат.помощь
   end;



   if DRound(xtmp,2)<>0 then
    begin
      //showmessage(floattostr(xtmp));
      xv:=xv+xtmp;
      DNal[13]:=DNal[13]-xtmp; //база
    end;


   if form1.kartSTATUS.Value='2' then st:=30 else st:=13;

   rtf:=false;
   sSoob:='';
   if Dround(DNal[13],2)<>DRound(xn-xv,2) then
    begin
     rtf:=true;
     sSoob:=sSoob+'База <> Доход - Вычеты'+#13+#13+floattostr(DNal[13])+#13+floattostr(xn)+#13+floattostr(xv);
    end;
   if ABS(Dround(DNal[13]*st/100,0)-DRound(DIsc[13],0))>1 then
    begin
     rtf:=true;
     sSoob:=sSoob+'База '+floattostr(DNal[13])+' * 0,13 <> НДФЛ '+floattostr(DIsc[13])+#13;
    end;

   if rtf then ShowMessage(sSoob+#13+form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+'Вызвано из Fspr2006.NewSpr2013');


end;


procedure TForm_58.RepVypl(pMes,pGod:Integer);
var ds,s,fNameXLS:String;
    tDat:TDate;
    i,m:integer;
    oldnls:real;
    nstroka:integer;
    NewValueArray: OLEVariant;
    E:Variant;
    PCol:array[1..31] of integer;
    DDAT:array[1..31] of TDate;
    DSDOXOD,DVICET,DNALOG:array[1..31] of Real;
    rtf:boolean;
    x1,x2,x3,x4:real;
begin
 ds:='repdec22.dbf';
  if table1.Active then table1.Active:=False;
  if table1.Exists then table1.DeleteTable;
  table1.TableName:=ds;
  table1.DatabaseName:=form1.DBDIR;
  table1.Exclusive:=True;
  table1.TableType:=ttDBase;
  table1.FieldDefs.Clear;
  table1.FieldDefs.Add('NLS',ftInteger,0,false);

  for i:=1 to 31 do
   begin
  //  table1.FieldDefs.Add('D_'+InttoStr(i),ftDate,0,false);   {дата}
    table1.FieldDefs.Add('P_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('K_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('S_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('V_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('N_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('IT_'+IntTostr(i),ftFloat,0,false); {оклад}
   end;
    i:=100;
    table1.FieldDefs.Add('P_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('S_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('V_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('N_'+IntTostr(i),ftFloat,0,false); {оклад}
    table1.FieldDefs.Add('IT_'+IntTostr(i),ftFloat,0,false); {оклад}

  table1.CreateTable;
  table1.Open;

   datam.qtmpstaj.close;
   datam.qtmpstaj.databasename:=form1.dbdir;
   datam.qtmpstaj.sql.clear;
   datam.qtmpstaj.sql.add('select k.fam,k.im,k.ot, s.*, o.datprov from kart k, sdoxod s, obrt2new o where k.nls=s.nls and k.nls=o.nls and s.idvypl=o.id');
   datam.qtmpstaj.sql.add('and s.sdoxod<>0 and mes='+floattostr(pMes)+' and god='+floattostr(pGod));
   datam.qtmpstaj.sql.add('order by o.datprov');
   datam.qtmpstaj.prepare;
   datam.qtmpstaj.open;
   datam.qtmpstaj.first;
   for i:=1 to 31 do DDAT[i]:=EncodeDate(1990,1,1);
   while not datam.qtmpstaj.eof do
    begin
      tDat:=datam.qtmpstaj.fieldbyname('datprov').asdatetime;

      m:=0; rtf:=false;
      for i:=1 to 31 do
       begin
        if DDAT[i]=tDat then
         begin
          m:=i;
          rtf:=true;
         end;
        if (DDAT[i]=EncodeDate(1990,1,1)) and (not rtf) then
          begin
           DDAT[i]:=tDat;
           rtf:=true;
           m:=i;
          end;
       end;

    //  ShowMessage(datetostr(tDat)+#13+floattostr(datam.qtmpstaj.fieldbyname('nls').asfloat)+#13+floattostr(datam.qtmpstaj.fieldbyname('sdoxod').asfloat));

      table1.first;
        if not table1.locate('nls;p_'+inttostr(m),VarArrayOf([datam.qtmpstaj.fieldbyname('nls').asfloat,0]),[locaseinsensitive]) then
         begin
          table1.append;
          for i:=1 to 31 do table1.fieldbyname('P_'+Inttostr(i)).asfloat:=0;
          table1.fieldbyname('nls').asfloat:=datam.qtmpstaj.fieldbyname('nls').asfloat;
          table1.post;
         end;
          table1.edit;
          table1.fieldbyname('P_'+Inttostr(m)).asfloat:=1;
          table1.fieldbyname('k_'+inttostr(m)).asfloat:=DRound(datam.qtmpstaj.fieldbyname('kodnac').asfloat,2);
          table1.fieldbyname('s_'+inttostr(m)).asfloat:=DRound(datam.qtmpstaj.fieldbyname('sdoxod').asfloat,2);
          table1.fieldbyname('v_'+inttostr(m)).asfloat:=DRound(datam.qtmpstaj.fieldbyname('rvicet').asfloat,2);
          table1.fieldbyname('n_'+inttostr(m)).asfloat:=DRound(datam.qtmpstaj.fieldbyname('nalog').asfloat,2);
          table1.fieldbyname('it_'+inttostr(m)).asfloat:=DRound(datam.qtmpstaj.fieldbyname('sdoxod').asfloat-datam.qtmpstaj.fieldbyname('nalog').asfloat,2);
          table1.post;

        table1.first;
        if table1.locate('nls',datam.qtmpstaj.fieldbyname('nls').asfloat,[locaseinsensitive]) then
         begin
          table1.edit;
          table1.fieldbyname('s_'+inttostr(100)).asfloat:=DRound(table1.fieldbyname('s_'+inttostr(100)).asfloat+datam.qtmpstaj.fieldbyname('sdoxod').asfloat,2);
          table1.fieldbyname('v_'+inttostr(100)).asfloat:=DRound(table1.fieldbyname('v_'+inttostr(100)).asfloat+datam.qtmpstaj.fieldbyname('rvicet').asfloat,2);
          table1.fieldbyname('n_'+inttostr(100)).asfloat:=DRound(table1.fieldbyname('n_'+inttostr(100)).asfloat+datam.qtmpstaj.fieldbyname('nalog').asfloat,2);
          table1.fieldbyname('it_'+inttostr(100)).asfloat:=DRound(table1.fieldbyname('it_'+inttostr(100)).asfloat+datam.qtmpstaj.fieldbyname('sdoxod').asfloat-datam.qtmpstaj.fieldbyname('nalog').asfloat,2);
          table1.post;
         end;


      datam.qtmpstaj.next;
    end;

    table1.close;

    datam.qtmp.close;
    datam.qtmp.databasename:=form1.dbdir;
    datam.qtmp.sql.clear;
    datam.qtmp.sql.add('select k.fam,k.im,k.ot, s.* from kart k, repdec22 s where k.nls=s.nls order by k.fam,k.im,k.ot');
    datam.qtmp.prepare;
    datam.qtmp.open;


  fNameXLS:=GetNameXlsn('Отчет_','repd')+'.xls';

 if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\repdec22.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLs) then
   begin
    MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
    exit;
   end;

 E:=CreateOleObject('Excel.Application');
 E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
 E.Visible:=True;
 try
  E.Application.WindowState:=2;
//  E.ActiveWindow.DisplayGridlines:=False;
 except
 end;

 E.ActiveWorkbook.Sheets.Item[1].Range['A1'].Value:=form1.config2NAME.Value+', ИНН '+form1.config2INN.value;
 E.ActiveWorkbook.Sheets.Item[1].Range['A2'].Value:='Отчет по выплате заработной платы за '+ansilowercase(namemes[pMEs])+' '+floattostr(pGod)+' г.';

   datam.qtmp.first;
   oldnls:=-1; nstroka:=4;
   for i:=1 to 31 do
     begin
      PCol[i]:=0;
      DSDOXOD[i]:=0;
      DVICET[i]:=0;
      DNALOG[i]:=0;
     end;

    for i:=1 to 31 do
      begin
       E.ActiveWorkBook.Sheets.Item[1].Range[GetAddr(5*i-3)+IntToStr(3)].Value:=formatdatetime('dd.mm.yyyy',DDAT[i])+' ['+inttostr(i)+']';
      end;

   NewValueArray := VarArrayCreate([1, 1, 1, 160], varVariant);
   while not datam.qtmp.eof do
    begin
      nstroka:=nstroka+1;
      for i:=1 to 160 do  newValueArray[1,i]:=null;
      if oldnls<>datam.qtmp.fieldbyname('nls').asfloat then
       begin
        form1.kart.locate('nls',datam.qtmp.fieldbyname('nls').asfloat,[locaseinsensitive]);
        newValueArray[1,1]:=form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.';
        E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka),GetAddr(160)+IntToStr(Nstroka)].Interior.Color:=WColorRep ;
       end;
       for i:=1 to 31 do
        begin
         if datam.qtmp.fieldbyname('P_'+inttostr(i)).asfloat=1 then
          begin
           PCol[i]:=1; //отображать
           if datam.qtmp.fieldbyname('k_'+inttostr(i)).asfloat=0 then s:='Оклад / тариф'
            else
              begin
               if form1.nacisl.locate('kod',datam.qtmp.fieldbyname('k_'+inttostr(i)).asfloat,[locaseinsensitive]) then s:=form1.nacislname.value else s:='???';
              end;
           newValueArray[1,i*5-3]:=ansilowercase(s);
           newValueArray[1,i*5-3+1]:=datam.qtmp.fieldbyname('s_'+inttostr(i)).asfloat;
           newValueArray[1,i*5-3+2]:=datam.qtmp.fieldbyname('v_'+inttostr(i)).asfloat;
           newValueArray[1,i*5-3+3]:=datam.qtmp.fieldbyname('n_'+inttostr(i)).asfloat;
           newValueArray[1,i*5-3+4]:=datam.qtmp.fieldbyname('it_'+inttostr(i)).asfloat;
           DSDOXOD[i]:=DSDOXOD[i]+datam.qtmp.fieldbyname('s_'+inttostr(i)).asfloat;
           DVICET[i]:=DVICET[i]+datam.qtmp.fieldbyname('v_'+inttostr(i)).asfloat;
           DNALOG[i]:=DNALOG[i]+datam.qtmp.fieldbyname('n_'+inttostr(i)).asfloat;

           if datam.qtmp.fieldbyname('s_100').asfloat<>0 then
            begin
             newValueArray[1,157]:=datam.qtmp.fieldbyname('s_100').asfloat;
             newValueArray[1,158]:=datam.qtmp.fieldbyname('v_100').asfloat;
             newValueArray[1,159]:=datam.qtmp.fieldbyname('n_100').asfloat;
             newValueArray[1,160]:=datam.qtmp.fieldbyname('it_100').asfloat;
            end;

         if not form1.TCheckOO then
           E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(NStroka),'FD'+IntToStr(NStroka)]:=NewValuearray
            else
              excellib.PExcelVyvod(E,1,'A'+IntToStr(NStroka),'FD'+IntToStr(NStroka),NewValuearray);

            for m:=7 to 12 do E.ActiveWorkbook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)+':FD'+IntToStr(Nstroka)].Borders[m].LineStyle:=1;

          end;
        end;

      oldnls:=datam.qtmp.fieldbyname('nls').asfloat;
     datam.qtmp.next;
    end;

     nstroka:=nstroka+1;
     for i:=1 to 160 do  newValueArray[1,i]:=null;

     x1:=0; x2:=0; x3:=0; x4:=0;
     for i:=1 to 31 do
      begin
        newValueArray[1,i*5-3+1]:=DSDOXOD[i];
        x1:=x1+DSDOXOD[i];
        newValueArray[1,i*5-3+2]:=DVICET[i];
        x2:=x2+DVICET[i];
        newValueArray[1,i*5-3+3]:=DNALOG[i];
        x3:=x3+DNALOG[i];
        newValueArray[1,i*5-3+4]:=DSDOXOD[i]-DNALOG[i];
        x4:=x4+DSDOXOD[i]-DNALOG[i];
      end;
      newValueArray[1,157]:=DRound(x1,2);
      newValueArray[1,158]:=DRound(x2,2);
      newValueArray[1,159]:=DRound(x3,2);
      newValueArray[1,160]:=DRound(x4,2);



       if not form1.TCheckOO then
           E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(NStroka),'FD'+IntToStr(NStroka)]:=NewValuearray
            else
              excellib.PExcelVyvod(E,1,'A'+IntToStr(NStroka),'FD'+IntToStr(NStroka),NewValuearray);

     for i:=1 to 31 do
      begin
       if PCol[i]=0 then
         begin
          E.ActiveWorkbook.Sheets.Item[1].Columns[i*5-3].Hidden:=true;
          E.ActiveWorkbook.Sheets.Item[1].Columns[i*5-3+1].Hidden:=true;
          E.ActiveWorkbook.Sheets.Item[1].Columns[i*5-3+2].Hidden:=true;
          E.ActiveWorkbook.Sheets.Item[1].Columns[i*5-3+3].Hidden:=true;
          E.ActiveWorkbook.Sheets.Item[1].Columns[i*5-3+4].Hidden:=true;
         end;
      end;
     E.Application.WindowState:=-4137;
     E:=UnAssigned;

end;


procedure TForm_58.Ispr2NDFLDec2022 ;   //испарвление 2-нДЛФ за декабрь 2022 - выплата з/п в январе исключить из справок за 2022г
var i,j:integer;
    skod:string;
    xv:real;
begin
   datam.qtmp.close;
   datam.qtmp.sql.clear;
   datam.qtmp.databasename:=form1.dbdir;
   datam.qtmp.sql.add('select * from sdoxod where nls='+floattostr(form1.kartnls.value));
   datam.qtmp.sql.add('and tavans<>1 and mes=12 and god=2022 and dat>='+#39+'01.01.2023'+#39);
   datam.qtmp.prepare;
   datam.qtmp.Open;
   datam.qtmp.first;
   while not datam.qtmp.eof do
     begin
      if datam.qtmp.FieldByName('kodnac').asfloat=0 then skod:='2000' else
       begin
        if form1.nacisl.locate('kod',datam.qtmp.fieldbyname('kodnac').asfloat,[locaseinsensitive]) then skod:=form1.NACISLKODDOX.Value;
       end;

      for j:=1 to 10 do
       begin
        if qqD[j]=skod then
         begin
          qqMT[j,12]:=DRound(qqMT[j,12]-datam.qtmp.fieldbyname('sdoxod').asfloat,2); //доход по коду декабрь
          qqMT[j,13]:=DRound(qqMT[j,13]-datam.qtmp.fieldbyname('sdoxod').asfloat,2); //доход по коду год
         end;
       end;
      if DRound(datam.qtmp.fieldbyname('rvicet').asfloat,2)<>0 then
        begin
         xv:=DRound(datam.qtmp.fieldbyname('rvicet').asfloat,2);   // учитывает только стандартные вычеты а если имущественный нет, доработать нужно булет
        // Showmessage(' уменьшаемы вычеты на '+floattostr(xv));
         for j:=6 downto 1 do
          begin
         //  showmessage('j='+inttostr(j)+' было '+DS20[j]+' '+floattostr(DST20[j,13]));

           if (DS20[j]<>'') and (DST20[j,13]<>0) then
            begin
             if xv<=DST20[j,13] then
               begin
                DST20[j,13]:=DRound(DST20[j,13]-xv,2);
                xv:=0;
               end
                else
              begin
               xv:=DRound(xv-DST20[j,13],2);
               DST20[j,13]:=0;
              end;
            end;
            // showmessage('j='+inttostr(j)+' стало '+DS20[j]+' '+floattostr(DST20[j,13]));
          end;
        end;

      DISC[13]:=DRound(DISC[13]-datam.qtmp.fieldbyname('nalog').asfloat,2);
      DNal[13]:=DRound(DNal[13]-(datam.qtmp.fieldbyname('sdoxod').asfloat-datam.qtmp.fieldbyname('rvicet').asfloat),2);  //база
     datam.qtmp.Next;
   end;

 //  for j:=1 to 6 do  showmessage('ИТОГ: '+DS20[j]+' '+floattostr(DST20[j,13]));



end;

procedure TForm_58.IsprDat31122022; //испарвляет дату выплаты дохода с 31.12.22 на 30.12.22
begin
 datam.qTmp.Close;
 datam.qTmp.DatabaseName:=form1.DBDIR;
 datam.qTmp.SQL.Clear;
 datam.qtmp.sql.add('update sdoxod set dat='+#39+'30.12.2022'+#39);
 datam.qtmp.sql.add('where dat='+#39+'31.12.2022'+#39);
 datam.qTmp.prepare;
 datam.qTmp.ExecSQL;
 datam.qtmp.close;
end;


function TForm_58.FGetOktmo(wm,wg:integer):String; //
var rtf:boolean;
    i:integer;
    tDat:TDate;
    s:String;
begin
 tDat:=EncodeDAte(wg,wm,1);
 rtf:=false;
 s:='';
 for i:=1 to 10 do
   begin
    SOK2[i]:=trim(SOK2[i]);
    if (SOK2[i]<>'') and (DOK2[i]<=tDAt) then
      begin
       s:=trim(SOK2[i])
      end;
   end;
 FGetOktmo:=s;
end;


function TForm_58.FOktmo(wm,wg:integer;sOktmo:String):boolean; // да/нет октмо в данном периоде предварительно ZapolnDOK
var rtf:boolean;
    i:integer;
    tDat:TDate;
begin
 tDat:=EncodeDAte(wg,wm,1);
 rtf:=false;
 for i:=1 to 10 do
   begin
    if (SOK2[i]<>'') and (DOK2[i]<=tDAt) then
      begin
       if (trim(sOktmo)=trim(SOK2[i])) then rtf:=true else rtf:=false;
      end;
   end;
 FOktmo:=rtf;
end;


function TForm_58.ZapolnDOK(tNls:Real):Integer;  //заполняет DOK[i] - да/нет с нужным ОКТМО в месяце i
var i,j:integer;
    tDat:Tdate;
    s:String;
    rtf,rtf2:boolean;
    TOK:Integer;
begin
 datam.qTmp2.Close;
 datam.qTmp2.DatabaseName:=form1.DBDIR;
 datam.qTmp2.SQL.Clear;
 datam.qTmp2.sql.add('select * from oktmonls where nls='+floattostr(tNls)+' order by dat');
 datam.qTmp2.Prepare;
 datam.qTmp2.Open;
 datam.qTmp2.first;
 datam.kart2.locate('nls',form1.kartNls.Value,[loCaseInsensitive]);

 for i:=1 to 10 do SOK2[i]:='';
 for i:=1 to 10 do DOK2[i]:=EncodeDate(1990,1,1);

 SOK2[1]:=form1.config2OKTMO.Value;
 if datam.kart2OKTMO.Value<>'' then SOK2[1]:=datam.kart2OKTMO.Value;

 TOk:=0; j:=1;
 while not datam.qTmp2.eof do
  begin
    rtf2:=true;
    if (j=1) and (trim(datam.qTmp2.fieldbyname('oktmo').asString)=SOK2[1]) then rtf2:=false; //не добавляем в список, т.к. дублирует первый организацию

    if rtf2 then
     begin
      j:=j+1;
      DOK2[j]:=datam.qTmp2.fieldbyname('dat').asDatetime;
      SOK2[j]:=trim(datam.qTmp2.fieldbyname('oktmo').asString);
     end;

   datam.qTmp2.Next;
  end;


 if j=1 then
  begin
   TOK:=0;   //нет записей в базе oktmo по сотруднику
  end
   else
  begin
   TOK:=1;      //есть записи
  end;

 datam.qTmp2.Close;

 ZapolnDOK:=TOK;

end;


procedure TForm_58.P3Vyvod(E2:oleVariant;klist:integer;Znacenie:String;nStroka:Integer;StartCol,EndCol:string)      ;  //вывод через ячейку
var i,j,k1,k2,nCount:integer;
    sAddr,s:String;
    NMASV: OLEVariant;
    k:integer;
    MyRange:OleVariant;
begin
 if trim(znacenie)='' then exit;
 k1:=excellib.GetNumCol(startcol);
 k2:=excellib.GetNumCol(endcol);


 nCount:=k2-k1+1;

 NMASV:= VarArrayCreate([1, 1, 1, 120], varVariant);

 for i:=1 to 120 do NMASV[1,i]:='';

 k:=k1;
 for i:=1 to length(znacenie) do
  begin
   NMASV[1,i*3-2]:=Copy(Znacenie,i,1);
   if form1.TCheckOO then E2.ActiveWorkBook.Sheets.Item[kList].Range[GetAddr(k)+inttostr(nstroka)].Value:=Copy(Znacenie,i,1);  //OpenOffice  массив не работает
   k:=k+3;
  end;
  if (not form1.TCheckOO) then E2.ActiveWorkBook.Sheets.Item[kList].Range[StartCol+IntToStr(nStroka)+':'+EndCol+IntToStr(nStroka)]:=NMASV; //Excel массивом

end;


procedure TForm_58.PVyvod(E:oleVariant;nList:Integer;Adr:String;Znacenie:String;NeVyvod:String);
begin
 if trim(Znacenie)<>NeVyVod then E.ActiveWorkBook.Sheets.Item[nList].Range[Adr].Value:=Znacenie;
end;





procedure TForm_58.PSpr2019(wnls:Real);
var fNameXLS,fName:String;
    E:OleVAriant;
    xRegion:Real;
    dd,mm,yy:Word;
    s1:String;
    nd,nc,i,j,k:Integer;
    x:Real;
    soktmo:String;
    sKpp:String;
    nSpr:Real;
    nczap:integer;
    nDatSpr:TDate;
    xFam,xIm,xOt:String;
    Nz0,Nz:Integer;
    RtfI:Boolean;
    TOk:Real;
    NSUVED:integer;

begin

    fNameXls:=GetNameXlsn('2НДФЛ','nd')+'.xls';

    fName:='2NDFL2022.xls';
    if RGod>=2023 then fName:='2NDFL2023.xls';

    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\'+fNAme,GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      exit;
     end;

   form102.RxLabel1.Caption:='Заполнение формы Excel Открытие шаблона'   ;
   form102.ProgressBar1.Position:=0;
   form102.Show;
   form102.Refresh;
   Nz0:=100;
   Nz:=0;

    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);

    try
    // E.Visible:=False;
     E.Application.WindowState:=2;
    except
    end;
    
    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;


     s:=Floattostr(RGod);
   //  E.ActiveWorkBook.Sheets.Item[1].Range['AP10']:=Copy(s,1,1);
   //  E.ActiveWorkBook.Sheets.Item[1].Range['AS10']:=Copy(s,2,1);

   //  E.ActiveWorkBook.Sheets.Item[1].Range['AV10']:=Copy(s,3,1);
   //  E.ActiveWorkBook.Sheets.Item[1].Range['AY10']:=Copy(s,4,1);

 // ??????????  E.ActiveWorkBook.Sheets.Item[1].Range['BK10']:=RPriznak;

     nSpr:=INT(rxCalcEdit1.Value);
     rxCalcEdit1.Value:=rxCAlcEdit1.Value+1;
     nDatSpr:=DateEdit1.DAte;

     //****
     if PRNUMDAT THEN
      BEGIN
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select num,dat from spr2006 where nls='+FloatToStr(form1.kartNls.Value));
       datam.Query1.SQL.Add('and GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       if datam.Query1.RecordCount>0 then
         begin
          nSpr:=datam.Query1.Fields[0].asFloat;
          nDatSpr:=datam.Query1.Fields[1].asDateTime;
          rxCalcEdit1.Value:=rxCAlcEdit1.Value-1;
         end;
       datam.Query1.Close;
      END;
     //**

     s:=DopolnCHR(Floattostr(nSpr),'-',6);
     P3Vyvod(E,1,s,10,'P','AH');

     P3Vyvod(E,2,s,7,'AA','AS');
     P3Vyvod(E,3,s,7,'AA','AS');

   {
     PVyVod(E,1,'J10',Copy(s,1,1),'0');
     PVyVod(E,1,'M10',Copy(s,2,1),'0');
     PVyVod(E,1,'P10',Copy(s,3,1),'0');
     PVyVod(E,1,'S10',Copy(s,4,1),'0');
     PVyVod(E,1,'V10',Copy(s,5,1),'0');
     PVyVod(E,1,'Y10',Copy(s,6,1),'0');
     PVyVod(E,1,'AB10',Copy(s,7,1),'0');

   }
     s:=trim(edit3.text); //коррект номер
     PVyVod(E,1,'CD10',Copy(s,1,1),'0');;
     PVyVod(E,1,'CG10',Copy(s,2,1),'0');;


     s:=DopolnCH(IntToStr(NUMLIST),'0',2);
     P3Vyvod(E,1,s,4,'BR','BX');
     NUMLIST:=NUMLIST+1;

     s:=DopolnCH(IntToStr(NUMLIST),'0',2);
     P3Vyvod(E,2,s,4,'BR','BX');
     NUMLIST:=NUMLIST+1;

     s:=DopolnCHR(form1.config2INN.asString,'-',11);
     P3Vyvod(E,1,s,1,'AK','BR');

   {
     PVyVod(E,1,'AK1',Copy(s,1,1),'0');
     PVyVod(E,1,'AN1',Copy(s,2,1),'0');
     PVyVod(E,1,'AQ1',Copy(s,3,1),'0');
     PVyVod(E,1,'AT1',Copy(s,4,1),'0');
     PVyVod(E,1,'AW1',Copy(s,5,1),'0');
     PVyVod(E,1,'AZ1',Copy(s,6,1),'0');
     PVyVod(E,1,'BC1',Copy(s,7,1),'0');
     PVyVod(E,1,'BF1',Copy(s,8,1),'0');
     PVyVod(E,1,'BI1',Copy(s,9,1),'0');
     PVyVod(E,1,'BL1',Copy(s,10,1),'0');
     PVyVod(E,1,'BO1',Copy(s,11,1),'0');
     PVyVod(E,1,'BR1',Copy(s,12,1),'0');
   }

  {
     if form_58.RKOD<>'' then
      begin
       PVyVod(E,1,'AE20',form_58.RKOD,'');

       s:=DopolnCH(trim(form_58.RINN),'0',9);
       PVyVod(E,1,'BI20',Copy(s,1,1),'0');
       PVyVod(E,1,'BL20',Copy(s,2,1),'0');
       PVyVod(E,1,'BO20',Copy(s,3,1),'0');
       PVyVod(E,1,'BR20',Copy(s,4,1),'0');
       PVyVod(E,1,'BU20',Copy(s,5,1),'0');
       PVyVod(E,1,'BX20',Copy(s,6,1),'0');
       PVyVod(E,1,'CA20',Copy(s,7,1),'0');
       PVyVod(E,1,'CD20',Copy(s,8,1),'0');
       PVyVod(E,1,'CG20',Copy(s,9,1),'0');
       PVyVod(E,1,'CJ20',Copy(s,10,1),'0');

       s:=DopolnCH(trim(form_58.RKPP),'0',8);
       PVyVod(E,1,'CP20',Copy(s,1,1),'0');
       PVyVod(E,1,'CS20',Copy(s,2,1),'0');
       PVyVod(E,1,'CV20',Copy(s,3,1),'0');
       PVyVod(E,1,'CY20',Copy(s,4,1),'0');
       PVyVod(E,1,'DB20',Copy(s,5,1),'0');
       PVyVod(E,1,'DE20',Copy(s,6,1),'0');
       PVyVod(E,1,'DH20',Copy(s,7,1),'0');
       PVyVod(E,1,'DK20',Copy(s,8,1),'0');
       PVyVod(E,1,'DN20',Copy(s,9,1),'0');

      end;
  }
   nz:=10;
   form102.ProgressBar1.Position:=Trunc(100*nz/nz0);
   form102.RxLabel1.Caption:='Заполнение данных организации'   ;
   form102.Refresh;

    s:=trim(form1.config2NAME.Value);
  {
    P3Vyvod(E,1,Copy(s,1,40),13,'A','DN');
    P3Vyvod(E,1,Copy(s,41,40),15,'A','DN');
    P3Vyvod(E,1,Copy(s,81,40),17,'A','DN');
  }
   {
    Pvyvod(E,1,'A13',Copy(s,1,1),'');
    PvyVod(E,1,'D13',Copy(s,2,1),'');
    PvyVod(E,1,'G13',Copy(s,3,1),'');
    PvyVod(E,1,'J13',Copy(s,4,1),'');
    PvyVod(E,1,'M13',Copy(s,5,1),'');
    PvyVod(E,1,'P13',Copy(s,6,1),'');
    PvyVod(E,1,'S13',Copy(s,7,1),'');
    PvyVod(E,1,'V13',Copy(s,8,1),'');
    PvyVod(E,1,'Y13',Copy(s,9,1),'');
    PvyVod(E,1,'AB13',Copy(s,10,1),'');
    PvyVod(E,1,'AE13',Copy(s,11,1),'');
    PvyVod(E,1,'AH13',Copy(s,12,1),'');
    PvyVod(E,1,'AK13',Copy(s,13,1),'');
    PvyVod(E,1,'AN13',Copy(s,14,1),'');
    PvyVod(E,1,'AQ13',Copy(s,15,1),'');
    PvyVod(E,1,'AT13',Copy(s,16,1),'');
    PvyVod(E,1,'AW13',Copy(s,17,1),'');
    PvyVod(E,1,'AZ13',Copy(s,18,1),'');
    PvyVod(E,1,'BC13',Copy(s,19,1),'');
    PvyVod(E,1,'BF13',Copy(s,20,1),'');
    PvyVod(E,1,'BI13',Copy(s,21,1),'');
    PvyVod(E,1,'BL13',Copy(s,22,1),'');
    PvyVod(E,1,'BO13',Copy(s,23,1),'');
    PvyVod(E,1,'BR13',Copy(s,24,1),'');
    PvyVod(E,1,'BU13',Copy(s,25,1),'');
    PvyVod(E,1,'BX13',Copy(s,26,1),'');
    PvyVod(E,1,'CA13',Copy(s,27,1),'');
    PvyVod(E,1,'CD13',Copy(s,28,1),'');
    PvyVod(E,1,'CG13',Copy(s,29,1),'');
    PvyVod(E,1,'CJ13',Copy(s,30,1),'');
    PvyVod(E,1,'CM13',Copy(s,31,1),'');
    PvyVod(E,1,'CP13',Copy(s,32,1),'');
    PvyVod(E,1,'CS13',Copy(s,33,1),'');
    PvyVod(E,1,'CV13',Copy(s,34,1),'');
    PvyVod(E,1,'CY13',Copy(s,35,1),'');
    PvyVod(E,1,'DB13',Copy(s,36,1),'');
    PvyVod(E,1,'DE13',Copy(s,37,1),'');
    PvyVod(E,1,'DH13',Copy(s,38,1),'');
    PvyVod(E,1,'DK13',Copy(s,39,1),'');
    PvyVod(E,1,'DN13',Copy(s,40,1),'');

    PvyVod(E,1,'A15',Copy(s,41,1),'');
    PvyVod(E,1,'D15',Copy(s,42,1),'');
    PvyVod(E,1,'G15',Copy(s,43,1),'');
    PvyVod(E,1,'J15',Copy(s,44,1),'');
    PvyVod(E,1,'M15',Copy(s,45,1),'');
    PvyVod(E,1,'P15',Copy(s,46,1),'');
    PvyVod(E,1,'S15',Copy(s,47,1),'');
    PvyVod(E,1,'V15',Copy(s,48,1),'');
    PvyVod(E,1,'Y15',Copy(s,49,1),'');
    PvyVod(E,1,'AB15',Copy(s,50,1),'');
    PvyVod(E,1,'AE15',Copy(s,51,1),'');
    PvyVod(E,1,'AH15',Copy(s,52,1),'');
    PvyVod(E,1,'AK15',Copy(s,53,1),'');
    PvyVod(E,1,'AN15',Copy(s,54,1),'');
    PvyVod(E,1,'AQ15',Copy(s,55,1),'');
    PvyVod(E,1,'AT15',Copy(s,56,1),'');
    PvyVod(E,1,'AW15',Copy(s,57,1),'');
    PvyVod(E,1,'AZ15',Copy(s,58,1),'');
    PvyVod(E,1,'BC15',Copy(s,59,1),'');
    PvyVod(E,1,'BF15',Copy(s,60,1),'');
    PvyVod(E,1,'BI15',Copy(s,61,1),'');
    PvyVod(E,1,'BL15',Copy(s,62,1),'');
    PvyVod(E,1,'BO15',Copy(s,63,1),'');
    PvyVod(E,1,'BR15',Copy(s,64,1),'');
    PvyVod(E,1,'BU15',Copy(s,65,1),'');
    PvyVod(E,1,'BX15',Copy(s,66,1),'');
    PvyVod(E,1,'CA15',Copy(s,67,1),'');
    PvyVod(E,1,'CD15',Copy(s,68,1),'');
    PvyVod(E,1,'CG15',Copy(s,69,1),'');
    PvyVod(E,1,'CJ15',Copy(s,70,1),'');
    PvyVod(E,1,'CM15',Copy(s,71,1),'');
    PvyVod(E,1,'CP15',Copy(s,72,1),'');
    PvyVod(E,1,'CS15',Copy(s,73,1),'');
    PvyVod(E,1,'CV15',Copy(s,74,1),'');
    PvyVod(E,1,'CY15',Copy(s,75,1),'');
    PvyVod(E,1,'DB15',Copy(s,76,1),'');
    PvyVod(E,1,'DE15',Copy(s,77,1),'');
    PvyVod(E,1,'DH15',Copy(s,78,1),'');
    PvyVod(E,1,'DK15',Copy(s,79,1),'');
    PvyVod(E,1,'DN15',Copy(s,80,1),'');
  }


     datam.kart2.Locate('nls',form1.kartnls.value,[locaseinsensitive]);




     if Trim(datam.kart2OKTMO.Value)='' then soktmo:=form1.config2OKTMO.Value else soktmo:=datam.kart2OKTMO.Value;

     //****
     RtfI:=true;
     TOK:=form_58.ZapolnDOK(form1.kartNls.Value);   //изменение ОКТМО
       if TOK=1 then
         begin
          s:=FGetOktmo(1,RGod);
          for i:=2 to 12 do if FGEtOktmo(i,RGod)<>s then RtfI:=false;
          if RtfI then soktmo:=s else
              begin
               soktmo:=trim(combobox1.text);
               MessageDlg(form1.kartfam.value+' обнаружено изменение ОКТМО в течение года'+#13+
                  'Справка формируется по ОКТМО '+soktmo+#13+
                    'Для формирования по другому ОКТМО выберите значение из списка КОД ОКТМО и повторите',mtInformation,[mbOk],0);
              end;   //октмо на начало года, изменений в течение года не было
         end;


     s:=DopolnCHR(soktmo,'-',10);
   //  P3Vyvod(E,1,s,22,'P','AT');


   {
     PVyVod(E,1,'P22',Copy(s,1,1),'0');
     PVyVod(E,1,'S22',Copy(s,2,1),'0');
     PVyVod(E,1,'V22',Copy(s,3,1),'0');
     PVyVod(E,1,'Y22',Copy(s,4,1),'0');
     PVyVod(E,1,'AB22',Copy(s,5,1),'0');
     PVyVod(E,1,'AE22',Copy(s,6,1),'0');
     PVyVod(E,1,'AH22',Copy(s,7,1),'0');
     PVyVod(E,1,'AK22',Copy(s,8,1),'0');
     PVyVod(E,1,'AN22',Copy(s,9,1),'0');
     PVyVod(E,1,'AQ22',Copy(s,10,1),'0');
     PVyVod(E,1,'AT22',Copy(s,11,1),'0');
     }

     if oktmo.Locate('oktmo',soktmo,[locaseinsensitive]) then
      begin
       sKpp:=oktmoKpp.Value;
      end
       else
        skpp:=form1.config2KPP.Value;
      s:=DopolnCHR(sKPP,'-',8);
      P3Vyvod(E,1,s,4,'AK','BI');

     {
      PVyVod(E,1,'AK4',Copy(s,1,1),'0');
      PVyVod(E,1,'AN4',Copy(s,2,1),'0');
      PVyVod(E,1,'AQ4',Copy(s,3,1),'0');
      PVyVod(E,1,'AT4',Copy(s,4,1),'0');
      PVyVod(E,1,'AW4',Copy(s,5,1),'0');
      PVyVod(E,1,'AZ4',Copy(s,6,1),'0');
      PVyVod(E,1,'BC4',Copy(s,7,1),'0');
      PVyVod(E,1,'BF4',Copy(s,8,1),'0');
      PVyVod(E,1,'BI4',Copy(s,9,1),'0');
     }

     s:=form_58.FSetGni(soktmo);
   //  P3Vyvod(E,1,s,10,'DE','DN');
    {
     E.ActiveWorkBook.Sheets.Item[1].Range['DE10']:=Copy(s,1,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['DH10']:=Copy(s,2,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['DK10']:=Copy(s,3,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['DN10']:=Copy(s,4,1);
    }

     //**

     s:=trim(form1.kartKODDOC.Value)  ;
     if length(s)=1 then s:=' '+s;
     E.ActiveWorkBook.Sheets.Item[1].Range['AB27'].Value:=Copy(s,1,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['AE27'].Value:=Copy(s,2,1);

     s:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;
   //  s:=DopolnCH(trim(s),' ',18);

     P3Vyvod(E,1,s,27,'BL','DN');

   {
     PVyvod(E,1,'BL36',Copy(s,1,1),'');
     PVyvod(E,1,'BO36',Copy(s,2,1),'');
     PVyvod(E,1,'BR36',Copy(s,3,1),'');
     PVyvod(E,1,'BU36',Copy(s,4,1),'');
     PVyvod(E,1,'BX36',Copy(s,5,1),'');
     PVyvod(E,1,'CA36',Copy(s,6,1),'');
     PVyvod(E,1,'CD36',Copy(s,7,1),'');
     PVyvod(E,1,'CG36',Copy(s,8,1),'');
     PVyvod(E,1,'CJ36',Copy(s,9,1),'');
     PVyvod(E,1,'CM36',Copy(s,10,1),'');
     PVyvod(E,1,'CP36',Copy(s,11,1),'');
     PVyvod(E,1,'CS36',Copy(s,12,1),'');
     PVyvod(E,1,'CV36',Copy(s,13,1),'');
     PVyvod(E,1,'CY36',Copy(s,14,1),'');
     PVyvod(E,1,'DB36',Copy(s,15,1),'');
     PVyvod(E,1,'DE36',Copy(s,16,1),'');
     PVyvod(E,1,'DH36',Copy(s,17,1),'');
     PVyvod(E,1,'DK36',Copy(s,18,1),'');
     PVyvod(E,1,'DN36',Copy(s,19,1),'');
    }


     //**

   //  s:=DopolnCH(trim(form1.config2TEL.Value),' ',19);
     s:=trim(form1.config2TEL.Value);

 //   P3Vyvod(E,1,s,22,'BI','DN');

    {
     PVyvod(E,1,'BI22',Copy(s,1,1),'');
     PVyvod(E,1,'BL22',Copy(s,2,1),'');
     PVyvod(E,1,'BO22',Copy(s,3,1),'');
     PVyvod(E,1,'BR22',Copy(s,4,1),'');
     PVyvod(E,1,'BU22',Copy(s,5,1),'');
     PVyvod(E,1,'BX22',Copy(s,6,1),'');
     PVyvod(E,1,'CA22',Copy(s,7,1),'');
     PVyvod(E,1,'CD22',Copy(s,8,1),'');
     PVyvod(E,1,'CG22',Copy(s,9,1),'');
     PVyvod(E,1,'CJ22',Copy(s,10,1),'');
     PVyvod(E,1,'CM22',Copy(s,11,1),'');
     PVyvod(E,1,'CP22',Copy(s,12,1),'');
     PVyvod(E,1,'CS22',Copy(s,13,1),'');
     PVyvod(E,1,'CV22',Copy(s,14,1),'');
     PVyvod(E,1,'CY22',Copy(s,15,1),'');
     PVyvod(E,1,'DB22',Copy(s,16,1),'');
     PVyvod(E,1,'DE22',Copy(s,17,1),'');
     PVyvod(E,1,'DH22',Copy(s,18,1),'');
     PVyvod(E,1,'DK22',Copy(s,19,1),'');
     PVyvod(E,1,'DN22',Copy(s,20,1),'');
    }

   nz:=20;
   form102.ProgressBar1.Position:=Trunc(100*nz/nz0);
   form102.RxLabel1.Caption:='Заполнение данных организации'   ;
   form102.Refresh;

     s:=DopolnCH(form1.kartINN.asString,'0',11);
     P3Vyvod(E,1,s,16,'AE','BL');
    {
     PVyVod(E,1,'CG24',Copy(s,1,1),'0');
     PVyVod(E,1,'CJ24',Copy(s,2,1),'0');
     PVyVod(E,1,'CM24',Copy(s,3,1),'0');
     PVyVod(E,1,'CP24',Copy(s,4,1),'0');
     PVyVod(E,1,'CS24',Copy(s,5,1),'0');
     PVyVod(E,1,'CV24',Copy(s,6,1),'0');
     PVyVod(E,1,'CY24',Copy(s,7,1),'0');
     PVyVod(E,1,'DB24',Copy(s,8,1),'0');
     PVyVod(E,1,'DE24',Copy(s,9,1),'0');
     PVyVod(E,1,'DH24',Copy(s,10,1),'0');
     PVyVod(E,1,'DK24',Copy(s,11,1),'0');
     PVyVod(E,1,'DN24',Copy(s,12,1),'0');
    }

     s:=DopolnCH(form1.kartSTRANA.Value,'0',2);
     P3Vyvod(E,1,s,25,'DH','DN');
   {
     E.ActiveWorkBook.Sheets.Item[1].Range['DH34']:=Copy(s,1,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['DK34']:=Copy(s,2,1);
     E.ActiveWorkBook.Sheets.Item[1].Range['DN34']:=Copy(s,3,1);
   }

    s:=trim(form1.kartFAM.Value);
    P3Vyvod(E,1,s,18,'P','BU');


   { PVyvod(E,1,'P27',Copy(s,1,1),'');
    PVyvod(E,1,'S27',Copy(s,2,1),'');
    PVyvod(E,1,'V27',Copy(s,3,1),'');
    PVyvod(E,1,'Y27',Copy(s,4,1),'');
    PVyvod(E,1,'AB27',Copy(s,5,1),'');
    PVyvod(E,1,'AE27',Copy(s,6,1),'');
    PVyvod(E,1,'AH27',Copy(s,7,1),'');
    PVyvod(E,1,'AK27',Copy(s,8,1),'');
    PVyvod(E,1,'AN27',Copy(s,9,1),'');
    PVyvod(E,1,'AQ27',Copy(s,10,1),'');
    PVyvod(E,1,'AT27',Copy(s,11,1),'');
    PVyvod(E,1,'AW27',Copy(s,12,1),'');
    PVyvod(E,1,'AZ27',Copy(s,13,1),'');
    PVyvod(E,1,'BC27',Copy(s,14,1),'');
    PVyvod(E,1,'BF27',Copy(s,15,1),'');
    PVyvod(E,1,'BI27',Copy(s,16,1),'');
    PVyvod(E,1,'BL27',Copy(s,17,1),'');
    PVyvod(E,1,'BO27',Copy(s,18,1),'');
    PVyvod(E,1,'BR27',Copy(s,19,1),'');
    PVyvod(E,1,'BU27',Copy(s,20,1),'');
   }


    s:=trim(form1.kartIM.Value);
    P3Vyvod(E,1,s,20,'P','CD');

{    PVyvod(E,1,'P29',Copy(s,1,1),'');
    PVyvod(E,1,'S29',Copy(s,2,1),'');
    PVyvod(E,1,'V29',Copy(s,3,1),'');
    PVyvod(E,1,'Y29',Copy(s,4,1),'');
    PVyvod(E,1,'AB29',Copy(s,5,1),'');
    PVyvod(E,1,'AE29',Copy(s,6,1),'');
    PVyvod(E,1,'AH29',Copy(s,7,1),'');
    PVyvod(E,1,'AK29',Copy(s,8,1),'');
    PVyvod(E,1,'AN29',Copy(s,9,1),'');
    PVyvod(E,1,'AQ29',Copy(s,10,1),'');
    PVyvod(E,1,'AT29',Copy(s,11,1),'');
    PVyvod(E,1,'AW29',Copy(s,12,1),'');
    PVyvod(E,1,'AZ29',Copy(s,13,1),'');
    PVyvod(E,1,'BC29',Copy(s,14,1),'');
    PVyvod(E,1,'BF29',Copy(s,15,1),'');
    PVyvod(E,1,'BI29',Copy(s,16,1),'');
    PVyvod(E,1,'BL29',Copy(s,17,1),'');
    PVyvod(E,1,'BO29',Copy(s,18,1),'');
    PVyvod(E,1,'BR29',Copy(s,19,1),'');
    PVyvod(E,1,'BU29',Copy(s,20,1),'');
}
    s:=trim(form1.kartOT.Value);
    P3Vyvod(E,1,s,22,'P','BU');

 {
    PVyvod(E,1,'P31',Copy(s,1,1),'');
    PVyvod(E,1,'S31',Copy(s,2,1),'');
    PVyvod(E,1,'V31',Copy(s,3,1),'');
    PVyvod(E,1,'Y31',Copy(s,4,1),'');
    PVyvod(E,1,'AB31',Copy(s,5,1),'');
    PVyvod(E,1,'AE31',Copy(s,6,1),'');
    PVyvod(E,1,'AH31',Copy(s,7,1),'');
    PVyvod(E,1,'AK31',Copy(s,8,1),'');
    PVyvod(E,1,'AN31',Copy(s,9,1),'');
    PVyvod(E,1,'AQ31',Copy(s,10,1),'');
    PVyvod(E,1,'AT31',Copy(s,11,1),'');
    PVyvod(E,1,'AW31',Copy(s,12,1),'');
    PVyvod(E,1,'AZ31',Copy(s,13,1),'');
    PVyvod(E,1,'BC31',Copy(s,14,1),'');
    PVyvod(E,1,'BF31',Copy(s,15,1),'');
    PVyvod(E,1,'BI31',Copy(s,16,1),'');
    PVyvod(E,1,'BL31',Copy(s,17,1),'');
    PVyvod(E,1,'BO31',Copy(s,18,1),'');
    PVyvod(E,1,'BR31',Copy(s,19,1),'');
    PVyvod(E,1,'BU31',Copy(s,20,1),'');
 }


     datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
     if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=7) then
                   E.ActiveWorkBook.Sheets.Item[1].Range['AB25'].Value:=datam.kart2STATUS2.AsString
                         else E.ActiveWorkBook.Sheets.Item[1].Range['AB25'].Value:='1';



      s:=Formatdatetime('dd.mm.yyyy',nDatSpr);
     // P3Vyvod(E,1,s,77,'BC','CD');
      E.ActiveWorkBook.Sheets.Item[1].Range['BR74'].Value:=s;

     { E.ActiveWorkBook.Sheets.Item[1].Range['BC77']:=Copy(s,1,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BF77']:=Copy(s,2,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BL77']:=Copy(s,4,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BO77']:=Copy(s,5,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BU77']:=Copy(s,7,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BX77']:=Copy(s,8,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['CA77']:=Copy(s,9,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['CD77']:=Copy(s,10,1);
     }

         E.ActiveWorkBook.Sheets.Item[2].Range['BR75'].Value:=s;
         E.ActiveWorkBook.Sheets.Item[3].Range['BR75'].Value:=s;


     s:=Formatdatetime('dd.mm.yyyy',form1.kartBIRTHDAY.Value);
     P3Vyvod(E,1,s,25,'BC','CD');
    {
      E.ActiveWorkBook.Sheets.Item[1].Range['BC34']:=Copy(s,1,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BF34']:=Copy(s,2,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BL34']:=Copy(s,4,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BO34']:=Copy(s,5,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BU34']:=Copy(s,7,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['BX34']:=Copy(s,8,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['CA34']:=Copy(s,9,1);
      E.ActiveWorkBook.Sheets.Item[1].Range['CD34']:=Copy(s,10,1);
    }


    {
       E.ActiveWorkBook.Sheets.Item[1].Range['AO20']:=form1.kartKODDOC.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BU20']:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;
    }

    nz:=30;
   form102.ProgressBar1.Position:=Trunc(100*nz/nz0);
   form102.RxLabel1.Caption:='Заполнение данных сотрудника'   ;
   form102.Refresh;

       s:=Floattostr(podpndflAgent.Value);
   //    E.ActiveWorkBook.Sheets.Item[1].Range['D64']:=Copy(s,1,1);

    GetFIO(podpndflFIO.Value,xFam,xIm,xOt);

    s:=xFam;
  //  P3Vyvod(E,1,s,67,'A','BF');
   {
    PVyvod(E,1,'A67',Copy(s,1,1),'');
    PVyvod(E,1,'D67',Copy(s,2,1),'');
    PVyvod(E,1,'G67',Copy(s,3,1),'');
    PVyvod(E,1,'J67',Copy(s,4,1),'');
    PVyvod(E,1,'M67',Copy(s,5,1),'');
    PVyvod(E,1,'P67',Copy(s,6,1),'');
    PVyvod(E,1,'S67',Copy(s,7,1),'');
    PVyvod(E,1,'V67',Copy(s,8,1),'');
    PVyvod(E,1,'Y67',Copy(s,9,1),'');
    PVyvod(E,1,'AB67',Copy(s,10,1),'');
    PVyvod(E,1,'AE67',Copy(s,11,1),'');
    PVyvod(E,1,'AH67',Copy(s,12,1),'');
    PVyvod(E,1,'AK67',Copy(s,13,1),'');
    PVyvod(E,1,'AN67',Copy(s,14,1),'');
    PVyvod(E,1,'AQ67',Copy(s,15,1),'');
    PVyvod(E,1,'AT67',Copy(s,16,1),'');
    PVyvod(E,1,'AW67',Copy(s,17,1),'');
    PVyvod(E,1,'AZ67',Copy(s,18,1),'');
    PVyvod(E,1,'BC67',Copy(s,19,1),'');
    PVyvod(E,1,'BF67',Copy(s,20,1),'');
   }

    s:=xIm;
 //   P3Vyvod(E,1,s,69,'A','BF');
  {
    PVyvod(E,1,'A69',Copy(s,1,1),'');
    PVyvod(E,1,'D69',Copy(s,2,1),'');
    PVyvod(E,1,'G69',Copy(s,3,1),'');
    PVyvod(E,1,'J69',Copy(s,4,1),'');
    PVyvod(E,1,'M69',Copy(s,5,1),'');
    PVyvod(E,1,'P69',Copy(s,6,1),'');
    PVyvod(E,1,'S69',Copy(s,7,1),'');
    PVyvod(E,1,'V69',Copy(s,8,1),'');
    PVyvod(E,1,'Y69',Copy(s,9,1),'');
    PVyvod(E,1,'AB69',Copy(s,10,1),'');
    PVyvod(E,1,'AE69',Copy(s,11,1),'');
    PVyvod(E,1,'AH69',Copy(s,12,1),'');
    PVyvod(E,1,'AK69',Copy(s,13,1),'');
    PVyvod(E,1,'AN69',Copy(s,14,1),'');
    PVyvod(E,1,'AQ69',Copy(s,15,1),'');
    PVyvod(E,1,'AT69',Copy(s,16,1),'');
    PVyvod(E,1,'AW69',Copy(s,17,1),'');
    PVyvod(E,1,'AZ69',Copy(s,18,1),'');
    PVyvod(E,1,'BC69',Copy(s,19,1),'');
    PVyvod(E,1,'BF69',Copy(s,20,1),'');
   }
    s:=xOt;
  //  P3Vyvod(E,1,s,71,'A','BF');
   {
    PVyvod(E,1,'A71',Copy(s,1,1),'');
    PVyvod(E,1,'D71',Copy(s,2,1),'');
    PVyvod(E,1,'G71',Copy(s,3,1),'');
    PVyvod(E,1,'J71',Copy(s,4,1),'');
    PVyvod(E,1,'M71',Copy(s,5,1),'');
    PVyvod(E,1,'P71',Copy(s,6,1),'');
    PVyvod(E,1,'S71',Copy(s,7,1),'');
    PVyvod(E,1,'V71',Copy(s,8,1),'');
    PVyvod(E,1,'Y71',Copy(s,9,1),'');
    PVyvod(E,1,'AB71',Copy(s,10,1),'');
    PVyvod(E,1,'AE71',Copy(s,11,1),'');
    PVyvod(E,1,'AH71',Copy(s,12,1),'');
    PVyvod(E,1,'AK71',Copy(s,13,1),'');
    PVyvod(E,1,'AN71',Copy(s,14,1),'');
    PVyvod(E,1,'AQ71',Copy(s,15,1),'');
    PVyvod(E,1,'AT71',Copy(s,16,1),'');
    PVyvod(E,1,'AW71',Copy(s,17,1),'');
    PVyvod(E,1,'AZ71',Copy(s,18,1),'');
    PVyvod(E,1,'BC71',Copy(s,19,1),'');
    PVyvod(E,1,'BF71',Copy(s,20,1),'');
   }
    s:=podpndflDOKUM.VAlue;
 //   P3Vyvod(E,1,s,75,'A','DN');
   {
    PVyvod(E,1,'A75',Copy(s,1,1),'');
    PVyvod(E,1,'D75',Copy(s,2,1),'');
    PVyvod(E,1,'G75',Copy(s,3,1),'');
    PVyvod(E,1,'J75',Copy(s,4,1),'');
    PVyvod(E,1,'M75',Copy(s,5,1),'');
    PVyvod(E,1,'P75',Copy(s,6,1),'');
    PVyvod(E,1,'S75',Copy(s,7,1),'');
    PVyvod(E,1,'V75',Copy(s,8,1),'');
    PVyvod(E,1,'Y75',Copy(s,9,1),'');
    PVyvod(E,1,'AB75',Copy(s,10,1),'');
    PVyvod(E,1,'AE75',Copy(s,11,1),'');
    PVyvod(E,1,'AH75',Copy(s,12,1),'');
    PVyvod(E,1,'AK75',Copy(s,13,1),'');
    PVyvod(E,1,'AN75',Copy(s,14,1),'');
    PVyvod(E,1,'AQ75',Copy(s,15,1),'');
    PVyvod(E,1,'AT75',Copy(s,16,1),'');
    PVyvod(E,1,'AW75',Copy(s,17,1),'');
    PVyvod(E,1,'AZ75',Copy(s,18,1),'');
    PVyvod(E,1,'BC75',Copy(s,19,1),'');
    PVyvod(E,1,'BF75',Copy(s,20,1),'');
    PVyvod(E,1,'BI75',Copy(s,21,1),'');
    PVyvod(E,1,'BL75',Copy(s,22,1),'');
    PVyvod(E,1,'BO75',Copy(s,23,1),'');
    PVyvod(E,1,'BR75',Copy(s,24,1),'');
    PVyvod(E,1,'BU75',Copy(s,25,1),'');
    PVyvod(E,1,'BX75',Copy(s,26,1),'');
    PVyvod(E,1,'CA75',Copy(s,27,1),'');
    PVyvod(E,1,'CD75',Copy(s,28,1),'');
    PVyvod(E,1,'CG75',Copy(s,29,1),'');
    PVyvod(E,1,'CJ75',Copy(s,30,1),'');
    PVyvod(E,1,'CM75',Copy(s,31,1),'');
    PVyvod(E,1,'CP75',Copy(s,32,1),'');
    PVyvod(E,1,'CS17',Copy(s,33,1),'');
    PVyvod(E,1,'CV75',Copy(s,34,1),'');
    PVyvod(E,1,'CY75',Copy(s,35,1),'');
    PVyvod(E,1,'DB75',Copy(s,36,1),'');
    PVyvod(E,1,'DE75',Copy(s,37,1),'');
    PVyvod(E,1,'DH75',Copy(s,38,1),'');
    PVyvod(E,1,'DK75',Copy(s,39,1),'');
    PVyvod(E,1,'DN75',Copy(s,40,1),'');
   }


     s:=DopolnCHR(form_58.sKBK ,'-',19);
     P3Vyvod(E,1,s,30,'AC','CH');
     P3Vyvod(E,2,s,9,'BH','DM');
     P3Vyvod(E,3,s,9,'BH','DM');

       if form1.kartSTATUS.Value='2' then
        begin
         // E.ActiveWorkBook.Sheets.Item[2].Range['CC9']:='3';
         // E.ActiveWorkBook.Sheets.Item[2].Range['CF9']:='0';
          E.ActiveWorkBook.Sheets.Item[1].Range['CU29'].Value:='3';
          E.ActiveWorkBook.Sheets.Item[1].Range['CX29'].Value:='0';
         {  E.ActiveWorkBook.Sheets.Item[3].Range['DN9']:='3';
          E.ActiveWorkBook.Sheets.Item[3].Range['DQ9']:='0';
         }
        end
           else
             begin
              // E.ActiveWorkBook.Sheets.Item[2].Range['CC9']:='1';
              //  E.ActiveWorkBook.Sheets.Item[2].Range['CF9']:='3';
                E.ActiveWorkBook.Sheets.Item[1].Range['CU29'].Value:='1';
                E.ActiveWorkBook.Sheets.Item[1].Range['CX29'].Value:='3';
              {
               E.ActiveWorkBook.Sheets.Item[2].Range['DN9']:='1';
               E.ActiveWorkBook.Sheets.Item[2].Range['DQ9']:='3';
               E.ActiveWorkBook.Sheets.Item[3].Range['DN9']:='1';
               E.ActiveWorkBook.Sheets.Item[3].Range['DQ9']:='3';
             }
             end;

       ObrabObrtNalKart;

       if not RtfI then //было изменение в течение года обнуляем месяцы с другими октмо
         begin
          ProcIsprOKTMO;
         end;


     if tIsprDec2022 then Ispr2NDFLDec2022; //исправляем декабрь 2022


    if RGod>=2023 then NewSpr2023; //перераспределяем коды доходов, вычетов с 2023г


  k:=0;
  for j:=1 to 10 do
    begin
     if qqD[j]<>'' then k:=j;  //кол-во кодов для дивидендов 13% добавить чтобы
    end;
  if RGod>=2015 then
   begin
    if DKOD9[1]<>'' then
     begin
      qqD[k+1]:=DKOD9[1];
      for i:=1 to 12 do qqMT[k+1,i]:=DDOX9[1,i];
      for i:=1 to 12 do qqMT[k+1,13]:=qqMT[k+1,13]+DDOX9[1,i];
     end;
   end;



     nc:=0;nd:=2;


   nz:=40;
   form102.ProgressBar1.Position:=Trunc(nz);
   form102.RxLabel1.Caption:='Заполнение сумм'   ;
   form102.Refresh;

     for i:=1 to 12 do
      begin

       if i mod 3 = 0 then
         begin
          nz:=40+i*5;
          form102.ProgressBar1.Position:=Trunc(nz);
          form102.RxLabel1.Caption:='Заполнение сумм'   ;
          form102.Refresh;
         end;

       for j:=1 to 10 do
        begin
          if (qqD[j]<>'') and (qqMT[j,i]<>0) then
           begin
               nc:=nc+1;

               if nc=16 then   //третий лист
                begin
                 nd:=3;
                 nc:=1;
                end;

               if i<10 then s:='0'+IntToStr(i) else s:=IntToStr(i);
               P3Vyvod(E,nd,s,nc*4+8,'K','N');
               {
               PVyVod(E,nd,'K'+IntToStr(nc*4+8),Copy(s,1,1),'0');
               PVyVod(E,nd,'N'+IntToStr(nc*4+8),Copy(s,2,1),'0');
               }
                 s:=mainlib.DopolnCH(Floattostr(FRub(qqMT[j,i])),'-',14)+'.'+SKop(qqMT[j,i]);;
                 P3Vyvod(E,nd,s,nc*4+8,'BH','DG');
                {
                 PVyVod(E,nd,'BH'+IntToStr(nc*4+8),Copy(s,1,1),'-');
                 PVyVod(E,nd,'BK'+IntToStr(nc*4+8),Copy(s,2,1),'-');
                 PVyVod(E,nd,'BN'+IntToStr(nc*4+8),Copy(s,3,1),'-');
                 PVyVod(E,nd,'BQ'+IntToStr(nc*4+8),Copy(s,4,1),'-');
                 PVyVod(E,nd,'BT'+IntToStr(nc*4+8),Copy(s,5,1),'-');
                 PVyVod(E,nd,'BW'+IntToStr(nc*4+8),Copy(s,6,1),'-');
                 PVyVod(E,nd,'BZ'+IntToStr(nc*4+8),Copy(s,7,1),'-');
                 PVyVod(E,nd,'CC'+IntToStr(nc*4+8),Copy(s,8,1),'-');
                 PVyVod(E,nd,'CF'+IntToStr(nc*4+8),Copy(s,9,1),'-');
                 PVyVod(E,nd,'CI'+IntToStr(nc*4+8),Copy(s,10,1),'-');
                 PVyVod(E,nd,'CL'+IntToStr(nc*4+8),Copy(s,11,1),'-');
                 PVyVod(E,nd,'CO'+IntToStr(nc*4+8),Copy(s,12,1),'-');
                 PVyVod(E,nd,'CR'+IntToStr(nc*4+8),Copy(s,13,1),'-');
                 PVyVod(E,nd,'CU'+IntToStr(nc*4+8),Copy(s,14,1),'-');
                 PVyVod(E,nd,'CX'+IntToStr(nc*4+8),Copy(s,15,1),'-');
                 PVyVod(E,nd,'DD'+IntToStr(nc*4+8),Copy(s,17,1),'-');
                 PVyVod(E,nd,'DG'+IntToStr(nc*4+8),Copy(s,18,1),'-');
                }
                 s:=DopolnCH(qqD[j],' ',3);
                 P3Vyvod(E,nd,s,nc*4+8,'AD','AM');
                {
                PVyVod(E,nd,'AD'+IntToStr(nc*4+8),Copy(s,1,1),'0');
                 PVyVod(E,nd,'AG'+IntToStr(nc*4+8),Copy(s,2,1),'0');
                 PVyVod(E,nd,'AJ'+IntToStr(nc*4+8),Copy(s,3,1),'0');
                 PVyVod(E,nd,'AM'+IntToStr(nc*4+8),Copy(s,4,1),'0');   ,
                 }
                 if (qqV[j]<>'') and (qqDV[j,i]<>0) then
                  begin
                    s:=DopolnCH(qqV[j],' ',2);
                    P3Vyvod(E,nd,s,nc*4+10,'AD','AJ');
                   {
                    E.ActiveWorkBook.Sheets.Item[nd].Range['AD'+IntToStr(nc*4+10)]:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['AG'+IntToStr(nc*4+10)]:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['AJ'+IntToStr(nc*4+10)]:=Copy(s,3,1);
                   }
                    s:=mainlib.DopolnCH(Floattostr(FRub(qqDV[j,i])),'-',13)+'.'+SKop(qqDV[j,i]);;
                    P3Vyvod(E,nd,s,nc*4+10,'BK','DG');
                   { E.ActiveWorkBook.Sheets.Item[nd].Range['BK'+IntToStr(nc*4+10)]:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['BN'+IntToStr(nc*4+10)]:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['BQ'+IntToStr(nc*4+10)]:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['BT'+IntToStr(nc*4+10)]:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['BW'+IntToStr(nc*4+10)]:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['BZ'+IntToStr(nc*4+10)]:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CC'+IntToStr(nc*4+10)]:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CF'+IntToStr(nc*4+10)]:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CI'+IntToStr(nc*4+10)]:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CL'+IntToStr(nc*4+10)]:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CO'+IntToStr(nc*4+10)]:=Copy(s,11,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CR'+IntToStr(nc*4+10)]:=Copy(s,12,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CU'+IntToStr(nc*4+10)]:=Copy(s,13,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['CX'+IntToStr(nc*4+10)]:=Copy(s,14,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['DD'+IntToStr(nc*4+10)]:=Copy(s,16,1);
                    E.ActiveWorkBook.Sheets.Item[nd].Range['DG'+IntToStr(nc*4+10)]:=Copy(s,17,1);
                   }
                  end;




           end;
        end;
       end;

      if nd=2 then
         begin
          E.ActiveWorkBook.Sheets.Item[3].Visible:=false ;  //третий лист
          if nc<>15 then
            begin
             nc:=nc+1  ;
            // E.ActiveWorkBook.Sheets.Item[nd].Range['K'+IntToStr(nc*4+8)+':CX68']:='';
            end;
         end;

      for j:=1 to 6 do
      begin
        if (DS20[j]<>'') and (DST20[j,13]<>0) then
         begin
            s:=DS20[j];
            if j<=3 then
             begin
              {
              E.ActiveWorkBook.Sheets.Item[1].Range['D'+IntToStr(j*2+50)]:=Copy(s,1,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['G'+IntToStr(j*2+50)]:=Copy(s,2,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['J'+IntToStr(j*2+50)]:=Copy(s,3,1);
              }
              P3Vyvod(E,1,s,j*2+42,'D','J');

              s:=mainlib.DopolnCH(Floattostr(FRub(DST20[j,13])),'-',7)+'.'+SKop(DST20[j,13]);;
              P3Vyvod(E,1,s,j*2+42,'R','AV');
              {
              E.ActiveWorkBook.Sheets.Item[1].Range['R'+IntToStr(j*2+50)]:=Copy(s,1,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['U'+IntToStr(j*2+50)]:=Copy(s,2,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['X'+IntToStr(j*2+50)]:=Copy(s,3,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AA'+IntToStr(j*2+50)]:=Copy(s,4,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AD'+IntToStr(j*2+50)]:=Copy(s,5,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AG'+IntToStr(j*2+50)]:=Copy(s,6,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AJ'+IntToStr(j*2+50)]:=Copy(s,7,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AM'+IntToStr(j*2+50)]:=Copy(s,8,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AS'+IntToStr(j*2+50)]:=Copy(s,10,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['AV'+IntToStr(j*2+50)]:=Copy(s,11,1);
              }
             end
              else
               begin
               P3Vyvod(E,1,s,(j-3)*2+42,'BN','BT') ;
               {
               E.ActiveWorkBook.Sheets.Item[1].Range['BN'+IntToStr((j-3)*2+50)]:=Copy(s,1,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['BQ'+IntToStr((j-3)*2+50)]:=Copy(s,2,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['BT'+IntToStr((j-3)*2+50)]:=Copy(s,3,1);
               }
               s:=mainlib.DopolnCH(Floattostr(FRub(DST20[j,13])),'-',7)+'.'+SKop(DST20[j,13]);;
               P3Vyvod(E,1,s,(j-3)*2+42,'CB','DF');
              {
               E.ActiveWorkBook.Sheets.Item[1].Range['CB'+IntToStr((j-3)*2+50)]:=Copy(s,1,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CE'+IntToStr((j-3)*2+50)]:=Copy(s,2,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CH'+IntToStr((j-3)*2+50)]:=Copy(s,3,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CK'+IntToStr((j-3)*2+50)]:=Copy(s,4,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CN'+IntToStr((j-3)*2+50)]:=Copy(s,5,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CQ'+IntToStr((j-3)*2+50)]:=Copy(s,6,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CT'+IntToStr((j-3)*2+50)]:=Copy(s,7,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['CW'+IntToStr((j-3)*2+50)]:=Copy(s,8,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['DC'+IntToStr((j-3)*2+50)]:=Copy(s,10,1);
               E.ActiveWorkBook.Sheets.Item[1].Range['DF'+IntToStr((j-3)*2+50)]:=Copy(s,11,1);
               }
              end;




         end;
      end;

    nz:=70;
   form102.ProgressBar1.Position:=Trunc(nz);
   form102.RxLabel1.Caption:='Вычеты'   ;
   form102.Refresh;


   NSUVED:=54;  //номер строки


    if form1.kartIMVYC_SUMM.Value<>0 then
     begin

      x:=0;
      for j:=1 to 12 do x:=x+DImVyc[j];
      if x<>0 then
       begin

        s:=form1.kartIMVYC_KOD.Value;
        E.ActiveWorkBook.Sheets.Item[1].Range['BN50'].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['BQ50'].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['BT50'].Value:=Copy(s,3,1);

        s:=form1.kartImVyc_Num.Value;
        E.ActiveWorkBook.Sheets.Item[1].Range['BX'+IntToStr(NSUVED)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CA'+IntToStr(NSUVED)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CD'+IntToStr(NSUVED)].Value:=Copy(s,3,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CG'+IntToStr(NSUVED)].Value:=Copy(s,4,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CJ'+IntToStr(NSUVED)].Value:=Copy(s,5,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CM'+IntToStr(NSUVED)].Value:=Copy(s,6,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CP'+IntToStr(NSUVED)].Value:=Copy(s,7,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CS'+IntToStr(NSUVED)].Value:=Copy(s,8,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CV'+IntToStr(NSUVED)].Value:=Copy(s,9,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CY'+IntToStr(NSUVED)].Value:=Copy(s,10,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DB'+IntToStr(NSUVED)].Value:=Copy(s,11,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DE'+IntToStr(NSUVED)].Value:=Copy(s,12,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DH'+IntToStr(NSUVED)].Value:=Copy(s,13,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DK'+IntToStr(NSUVED)].Value:=Copy(s,14,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DN'+IntToStr(NSUVED)].Value:=Copy(s,15,1);

        E.ActiveWorkBook.Sheets.Item[1].Range['AB'+IntToStr(NSUVED)].Value:='1';  //код вида уведомления имущ. вычет

        s:=Trim(form1.kartIMVYC_gni.Value);
        E.ActiveWorkBook.Sheets.Item[1].Range['DE'+IntToStr(NSUVED+2)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DH'+IntToStr(NSUVED+2)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DK'+IntToStr(NSUVED+2)].Value:=Copy(s,3,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DN'+IntToStr(NSUVED+2)].Value:=Copy(s,4,1);

        s:=FormatDAtetime('dd.mm.yyyy',form1.kartImVyc_Dat.Value);
        E.ActiveWorkBook.Sheets.Item[1].Range['AB'+IntToStr(NSUVED+2)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AE'+IntToStr(NSUVED+2)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AK'+IntToStr(NSUVED+2)].Value:=Copy(s,4,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(NSUVED+2)].Value:=Copy(s,5,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AT'+IntToStr(NSUVED+2)].Value:=Copy(s,7,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AW'+IntToStr(NSUVED+2)].Value:=Copy(s,8,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ'+IntToStr(NSUVED+2)].Value:=Copy(s,9,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['BC'+IntToStr(NSUVED+2)].Value:=Copy(s,10,1);
        NSUVED:=58;
        //**

        E.ActiveWorkBook.Sheets.Item[1].Range['CP47'].Value:=x;
        s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',7)+'.'+SKop(x);;
              E.ActiveWorkBook.Sheets.Item[1].Range['CB50'].Value:=Copy(s,1,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CE50'].Value:=Copy(s,2,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CH50'].Value:=Copy(s,3,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CK50'].Value:=Copy(s,4,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CN50'].Value:=Copy(s,5,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CQ50'].Value:=Copy(s,6,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CT50'].Value:=Copy(s,7,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['CW50'].Value:=Copy(s,8,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['DC50'].Value:=Copy(s,10,1);
              E.ActiveWorkBook.Sheets.Item[1].Range['DF50'].Value:=Copy(s,11,1);
       end;
       
     end;



   //******* Уведомление фикс платежи

   if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
     begin
      if trim(ndflr5num.Value)<>'' then
       begin
        s:=trim(ndflr5num.Value);
        E.ActiveWorkBook.Sheets.Item[1].Range['BX'+IntToStr(NSUVED)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CA'+IntToStr(NSUVED)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CD'+IntToStr(NSUVED)].Value:=Copy(s,3,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CG'+IntToStr(NSUVED)].Value:=Copy(s,4,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CJ'+IntToStr(NSUVED)].Value:=Copy(s,5,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CM'+IntToStr(NSUVED)].Value:=Copy(s,6,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CP'+IntToStr(NSUVED)].Value:=Copy(s,7,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CS'+IntToStr(NSUVED)].Value:=Copy(s,8,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CV'+IntToStr(NSUVED)].Value:=Copy(s,9,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['CY'+IntToStr(NSUVED)].Value:=Copy(s,10,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DB'+IntToStr(NSUVED)].Value:=Copy(s,11,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DE'+IntToStr(NSUVED)].Value:=Copy(s,12,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DH'+IntToStr(NSUVED)].Value:=Copy(s,13,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DK'+IntToStr(NSUVED)].Value:=Copy(s,14,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DN'+IntToStr(NSUVED)].Value:=Copy(s,15,1);



        E.ActiveWorkBook.Sheets.Item[1].Range['AB'+IntToStr(NSUVED)].Value:='3';  //код вида уведомления


        s:=Trim(ndflr5ifns.Value);
        E.ActiveWorkBook.Sheets.Item[1].Range['DE'+IntToStr(NSUVED+2)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DH'+IntToStr(NSUVED+2)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DK'+IntToStr(NSUVED+2)].Value:=Copy(s,3,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['DN'+IntToStr(NSUVED+2)].Value:=Copy(s,4,1);

        s:=FormatDAtetime('dd.mm.yyyy',ndflr5dat.Value);
        E.ActiveWorkBook.Sheets.Item[1].Range['AB'+IntToStr(NSUVED+2)].Value:=Copy(s,1,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AE'+IntToStr(NSUVED+2)].Value:=Copy(s,2,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AK'+IntToStr(NSUVED+2)].Value:=Copy(s,4,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(NSUVED+2)].Value:=Copy(s,5,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AT'+IntToStr(NSUVED+2)].Value:=Copy(s,7,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AW'+IntToStr(NSUVED+2)].Value:=Copy(s,8,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ'+IntToStr(NSUVED+2)].Value:=Copy(s,9,1);
        E.ActiveWorkBook.Sheets.Item[1].Range['BC'+IntToStr(NSUVED+2)].Value:=Copy(s,10,1);
     end;
    end;

   //****


    x:=0;
    for j:=1 to 10 do x:=x+qqMT[j,13];    //общий доход
     s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',14)+'.'+SKop(x);;
       P3Vyvod(E,1,s,32,'X','BW');
                  {
                    PVyVod(E,1,'O40',Copy(s,1,1),'-');
                    PVyVod(E,1,'R40',Copy(s,2,1),'-');
                    PVyVod(E,1,'U40',Copy(s,3,1),'-');
                    PVyVod(E,1,'X40',Copy(s,4,1),'-');
                    PVyVod(E,1,'AA40',Copy(s,5,1),'-');
                    PVyVod(E,1,'AD40',Copy(s,6,1),'-');
                    PVyVod(E,1,'AG40',Copy(s,7,1),'-');
                    PVyVod(E,1,'AJ40',Copy(s,8,1),'-');
                    PVyVod(E,1,'AM40',Copy(s,9,1),'-');
                    PVyVod(E,1,'AP40',Copy(s,10,1),'-');
                    PVyVod(E,1,'AS40',Copy(s,11,1),'-');
                    PVyVod(E,1,'AV40',Copy(s,12,1),'-');
                    PVyVod(E,1,'AY40',Copy(s,13,1),'-');
                    PVyVod(E,1,'BB40',Copy(s,14,1),'-');
                    PVyVod(E,1,'BE40',Copy(s,15,1),'-');
                    PVyVod(E,1,'BK40',Copy(s,17,1),'-');
                    PVyVod(E,1,'BN40',Copy(s,18,1),'-');
                  }


    nz:=90;
   form102.ProgressBar1.Position:=Trunc(nz);
   form102.RxLabel1.Caption:='Итоговые суммы'   ;
   form102.Refresh;

    x:=DNal[13];
    for j:=1 to 12 do x:=x+DDoxod9[j];   //база
      s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',14)+'.'+SKop(x);;
       P3Vyvod(E,1,s,34,'X','BW');
                  {
                    PVyVod(E,1,'O42',Copy(s,1,1),'-');
                    PVyVod(E,1,'R42',Copy(s,2,1),'-');
                    PVyVod(E,1,'U42',Copy(s,3,1),'-');
                    PVyVod(E,1,'X42',Copy(s,4,1),'-');
                    PVyVod(E,1,'AA42',Copy(s,5,1),'-');
                    PVyVod(E,1,'AD42',Copy(s,6,1),'-');
                    PVyVod(E,1,'AG42',Copy(s,7,1),'-');
                    PVyVod(E,1,'AJ42',Copy(s,8,1),'-');
                    PVyVod(E,1,'AM42',Copy(s,9,1),'-');
                    PVyVod(E,1,'AP42',Copy(s,10,1),'-');
                    PVyVod(E,1,'AS42',Copy(s,11,1),'-');
                    PVyVod(E,1,'AV42',Copy(s,12,1),'-');
                    PVyVod(E,1,'AY42',Copy(s,13,1),'-');
                    PVyVod(E,1,'BB42',Copy(s,14,1),'-');
                    PVyVod(E,1,'BE42',Copy(s,15,1),'-');
                    PVyVod(E,1,'BK42',Copy(s,17,1),'-');
                    PVyVod(E,1,'BN42',Copy(s,18,1),'-');
                  }



    x:=DIsc[13];
    for j:=1 to 13 do x:=x+DPn9[j];
    s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',10);
                 P3Vyvod(E,1,s,36,'X','BB');
                  {
                    PVyVod(E,1,'O44',Copy(s,1,1),'-');
                    PVyVod(E,1,'R44',Copy(s,2,1),'-');
                    PVyVod(E,1,'U44',Copy(s,3,1),'-');
                    PVyVod(E,1,'X44',Copy(s,4,1),'-');
                    PVyVod(E,1,'AA44',Copy(s,5,1),'-');
                    PVyVod(E,1,'AD44',Copy(s,6,1),'-');
                    PVyVod(E,1,'AG44',Copy(s,7,1),'-');
                    PVyVod(E,1,'AJ44',Copy(s,8,1),'-');
                    PVyVod(E,1,'AM44',Copy(s,9,1),'-');
                    PVyVod(E,1,'AP44',Copy(s,10,1),'-');
                    PVyVod(E,1,'AS44',Copy(s,11,1),'-');
                   }

                     P3Vyvod(E,1,s,36,'CJ','DN');
                   { PVyVod(E,1,'O46',Copy(s,1,1),'-');
                    PVyVod(E,1,'R46',Copy(s,2,1),'-');
                    PVyVod(E,1,'U46',Copy(s,3,1),'-');
                    PVyVod(E,1,'X46',Copy(s,4,1),'-');
                    PVyVod(E,1,'AA46',Copy(s,5,1),'-');
                    PVyVod(E,1,'AD46',Copy(s,6,1),'-');
                    PVyVod(E,1,'AG46',Copy(s,7,1),'-');
                    PVyVod(E,1,'AJ46',Copy(s,8,1),'-');
                    PVyVod(E,1,'AM46',Copy(s,9,1),'-');
                    PVyVod(E,1,'AP46',Copy(s,10,1),'-');
                    PVyVod(E,1,'AS46',Copy(s,11,1),'-');
                   }




    x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'')+FUplataNdfl(form1.kartNLS.Value,RGod,9,'');
    if not RtfI then x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,30,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,9,soktmo);

    s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',10);       //перечислено
                  P3Vyvod(E,1,s,40,'X','BB');
                  {
                    PVyVod(E,1,'CJ42',Copy(s,1,1),'-');
                    PVyVod(E,1,'CM42',Copy(s,2,1),'-');
                    PVyVod(E,1,'CP42',Copy(s,3,1),'-');
                    PVyVod(E,1,'CS42',Copy(s,4,1),'-');
                    PVyVod(E,1,'CV42',Copy(s,5,1),'-');
                    PVyVod(E,1,'CY42',Copy(s,6,1),'-');
                    PVyVod(E,1,'DB42',Copy(s,7,1),'-');
                    PVyVod(E,1,'DE42',Copy(s,8,1),'-');
                    PVyVod(E,1,'DH42',Copy(s,9,1),'-');
                    PVyVod(E,1,'DK42',Copy(s,10,1),'-');
                    PVyVod(E,1,'DN42',Copy(s,11,1),'-');
                  }

    datam.query1.close;
    datam.query1.sql.clear;
    datam.query1.sql.add('select * from uplatandfl where nls='+floattostr(form1.kartnls.value));
    datam.query1.sql.add('and type=2 and summa>0 and god='+floattostr(RGod));
    datam.query1.Prepare;
    datam.Query1.open;
    datam.query1.first;
     if datam.query1.RecordCount=1 then //аванс
      begin
        x:=DRound(datam.query1.FieldByName('summa').asFloat,2);
        s:=mainlib.DopolnCH(Floattostr(FRub(x)),'-',10);       //аванс плат
                    E.ActiveWorkBook.Sheets.Item[1].Range['X38'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AA38'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AD38'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AG38'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AJ38'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AM38'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AP38'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AS38'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AV38'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AY38'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BB38'].Value:=Copy(s,11,1);
        {
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59']:=datam.query1.FieldByName('numuved').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59']:=datam.query1.FieldByName('ifns').asString;
        DecodeDate(datam.query1.FieldByName('datuved').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59']:=s1;
        }
      end;
    datam.query1.close;

    if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
     begin

     if ndflr5PR2.Value=1 then
       begin
         s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('sud').asFloat)),'-',10);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CJ36'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CM36'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CP36'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CS36'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CV36'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CY36'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DB36'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DE36'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DH36'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DK36'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DN36'].Value:=Copy(s,11,1);
       end;

      if ndflr5PR1.Value=1 then
       begin
        s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('sisc').asFloat)),'-',10);
                    E.ActiveWorkBook.Sheets.Item[1].Range['X36'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AA36'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AD36'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AG36'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AJ36'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AM36'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AP36'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AS36'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AV36'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AY36'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BB36'].Value:=Copy(s,11,1);
       end;


         s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('suderj').asFloat)),'-',10);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CJ40'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CM40'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CP40'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CS40'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CV40'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CY40'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DB40'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DE40'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DH40'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DK40'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DN40'].Value:=Copy(s,11,1);


      s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('snuderj').asFloat)),'-',14);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BF70'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BI70'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BL70'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BO70'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BR70'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BU70'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BX70'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CA70'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CD70'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CG70'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CJ70'].Value:=Copy(s,11,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CM70'].Value:=Copy(s,12,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CP70'].Value:=Copy(s,13,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CS70'].Value:=Copy(s,14,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CV70'].Value:=Copy(s,15,1);



     s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('sfix').asFloat)),'-',10);
                    E.ActiveWorkBook.Sheets.Item[1].Range['X38'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AA38'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AD38'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AG38'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AJ38'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AM38'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AP38'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AS38'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AV38'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['AY38'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BB38'].Value:=Copy(s,11,1);

      s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('sprib').asFloat)),'-',10);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CJ38'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CM38'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CP38'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CS38'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CV38'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CY38'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DB38'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DE38'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DH38'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DK38'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DN38'].Value:=Copy(s,11,1);


      s:=mainlib.DopolnCH(Floattostr(FRub(ndflr5.FieldByName('doxneud').asFloat)),'-',14)+'.'+SKop(ndflr5.FieldByName('doxneud').asFloat);;
                    E.ActiveWorkBook.Sheets.Item[1].Range['BF68'].Value:=Copy(s,1,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BI68'].Value:=Copy(s,2,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BL68'].Value:=Copy(s,3,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BO68'].Value:=Copy(s,4,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BR68'].Value:=Copy(s,5,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BU68'].Value:=Copy(s,6,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['BX68'].Value:=Copy(s,7,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CA68'].Value:=Copy(s,8,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CD68'].Value:=Copy(s,9,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CG68'].Value:=Copy(s,10,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CJ68'].Value:=Copy(s,11,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CM68'].Value:=Copy(s,12,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CP68'].Value:=Copy(s,13,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CS68'].Value:=Copy(s,14,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['CV68'].Value:=Copy(s,15,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DB68'].Value:=Copy(s,17,1);
                    E.ActiveWorkBook.Sheets.Item[1].Range['DE68'].Value:=Copy(s,18,1);



      {if ndflr5.FieldByName('sfix').asFloat<>0 then
       begin
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59']:=ndflr5.FieldByName('num').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59']:=ndflr5.FieldByName('ifns').asString;
        DecodeDate(ndflr5.FieldByName('dat').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59']:=s1;
       end;
      }
     end;



   form102.Close;

   try
    E.DisplayAlerts:=false;
    E.WorkBooks[1].Save;
   except
   end;

    if E.ActiveWorkBook.Sheets.Item[3].Visible then
     begin
      s:=DopolnCH(IntToStr(NUMLIST),'0',2);
      P3Vyvod(E,3,s,4,'BR','BX');
      NUMLIST:=NUMLIST+1;
     end;

    E.Visible:=True;

    E.WindowState:=-4137 ;
    E:=UnAssigned;

 {

   x:=0;                         //35% отдельная справка
   for i:=1 to 12 do x:=x+DDoxod35[i];
   if x<>0 then
     begin
      rxCalcEdit1.Value:=rxCalcEdit1.Value+1;
      PSpr2016_st35(wnls);
     end;
 }
end;






function TForm_58.FKodNdfl6():String;
var sv:String;
begin
   sv:='000';
    if ComboBox3.ItemIndex=0 then sv:='120';
    if ComboBox3.ItemIndex=1 then sv:='124';
    if ComboBox3.ItemIndex=2 then sv:='125';
    if ComboBox3.ItemIndex=3 then sv:='126';
    if ComboBox3.ItemIndex=4 then sv:='213';
    if ComboBox3.ItemIndex=5 then sv:='214';
    if ComboBox3.ItemIndex=6 then sv:='215';
    if ComboBox3.ItemIndex=7 then sv:='216';
    if ComboBox3.ItemIndex=8 then sv:='220';
    if ComboBox3.ItemIndex=9 then sv:='320';
    if ComboBox3.ItemIndex=10 then sv:='335';

 FKodNdfl6:=sv;
end;


function TForm_58.SmenaKodAvans(mes,god,nls:Real):Boolean;
var xavans,x,y,z:Real;
           x2,y2:real;
  oldKod:Real;
  rtf:Boolean;
  xidnew,xIdvypl:real;
  xDAt:TDate;
begin

     rtf:=false;

          datam.qtmp3.close;
          datam.qtmp3.databasename:=form1.dbdir;
          datam.qtmp3.sql.clear;
          datam.qtmp3.sql.add('select max(id) from sdoxod');
          datam.qtmp3.prepare;
          datam.qtmp3.Open;
          xidnew:=datam.qtmp3.Fields[0].asfloat;
          datam.qtmp3.close;

          if form1.sdoxod.locate('nls;mes;god;kodnac',VarArrayOf([nls,mes,god,0]),[loCaseInsensitive]) then
           begin
            xIdvypl:=form1.sdoxodidvypl.Value;
            xDat:=form1.sdoxodDAT.Value
           end   ;


          //не учитывает что может быть два аванса
          datam.qtmp2.close;
          datam.qtmp2.Databasename:=form1.dbdir;
          datam.qtmp2.sql.clear;
          datam.qtmp2.sql.add('select sum(sdoxod) from sdoxod where nls='+floattostr(nls));
          datam.qtmp2.sql.add('and mes='+floattostr(mes));
          datam.qtmp2.sql.add('and god='+floattostr(god));
          datam.qtmp2.sql.add('and kodnac=0');               //сумма аванса
          datam.qtmp2.prepare;
          datam.qtmp2.open;
          xavans:=datam.qtmp2.fields[0].asfloat;
          datam.qtmp2.close;

        //  Showmessage(floattostr(xavans));

  datam.qtmp2.close;
  datam.qtmp2.Databasename:=form1.dbdir;
  datam.qtmp2.sql.clear;
  datam.qtmp2.sql.add('select o.kod,sum(o.kr),sum(o.snalog) from obrt1new o where o.wm='+floattostr(mes));
  datam.qtmp2.sql.add('and o.wg='+floattostr(god));
  datam.qtmp2.sql.add('and o.nls='+floattostr(nls));
  datam.qtmp2.sql.add('group by o.kod');
  datam.qtmp2.prepare;
  datam.qtmp2.open;


  //сначала проверка что хватает начислений с кодом 2000
  z:=0;
  datam.qtmp2.First;
  while not datam.qtmp2.eof do
   begin
      if form1.NACISL.Locate('kod',datam.qtmp2.fields[0].asFloat,[loCaseInsensitive]) then
       begin
        if form1.NACISLKODDOX.Value='2000' then
         begin
          x:=datam.qtmp2.Fields[1].asFloat;
          if form1.nacislRK.Value then x:=DRound(x,2)+DRound(x*form1.configRK.Value/100,2);

          datam.qtmp3.close;
          datam.qtmp3.Databasename:=form1.DBDIR;
          datam.qtmp3.sql.clear;
          datam.qtmp3.sql.add('select sum(sdoxod) from sdoxod where nls='+floattostr(nls));
          datam.qtmp3.sql.add('and mes='+floattostr(mes));
          datam.qtmp3.sql.add('and god='+floattostr(god));
          datam.qtmp3.sql.add('and kodnac='+floattostr(form1.nacislKod.Value));
          datam.qtmp3.prepare;
          datam.qtmp3.open;
          y:=datam.qtmp3.fields[0].asfloat;
          datam.qtmp3.close;

          z:=z+x-y;

         end;

       end;
     datam.qtmp2.next;
   end;

  //теперь замена если хватает

 // showmessage('Fspr2006-> z='+floattostr(z)+#13+'xavans='+floattostr(xavans));

  if z>=xavans then
   begin

     datam.qtmp2.First;
     while not datam.qtmp2.eof do
      begin
         if form1.NACISL.Locate('kod',datam.qtmp2.fields[0].asFloat,[loCaseInsensitive]) then
          begin
           if form1.NACISLKODDOX.Value='2000' then
            begin
             x:=datam.qtmp2.Fields[1].asFloat;
             if form1.nacislRK.Value then x:=DRound(x,2)+DRound(x*form1.configRK.Value/100,2);
             x2:=datam.qtmp2.Fields[2].asFloat;
             datam.qtmp3.close;
             datam.qtmp3.Databasename:=form1.DBDIR;
             datam.qtmp3.sql.clear;
             datam.qtmp3.sql.add('select sum(sdoxod),sum(nalog) from sdoxod where nls='+floattostr(nls));
             datam.qtmp3.sql.add('and mes='+floattostr(mes));
             datam.qtmp3.sql.add('and god='+floattostr(god));
             datam.qtmp3.sql.add('and kodnac='+floattostr(form1.nacislKod.Value));
             datam.qtmp3.prepare;
             datam.qtmp3.open;
             y:=datam.qtmp3.fields[0].asfloat;
             y2:=datam.qtmp3.fields[1].asfloat;
             datam.qtmp3.close;

             if not rtf then   //
               begin
                  if not form1.sdoxod.locate('nls;mes;god;kodnac',VarArrayOf([nls,mes,god,0]),[loCaseInsensitive]) then
                     begin
                       xidnew:=xidnew+1;
                       form1.sdoxod.Append;
                       form1.sdoxod.fieldbyname('id').asFloat:=xidnew;
                       form1.sdoxod.fieldbyname('tavans').asFloat:=1;
                       form1.sdoxod.fieldbyname('idvypl').asFloat:=xIdVypl;            //xIdvypl=id obrt2 !!!
                       form1.sdoxod.fieldbyname('nls').asFloat:=nls;
                       form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                       form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
                       form1.sdoxod.fieldbyname('mes').asFloat:=mes;
                       form1.sdoxod.fieldbyname('god').asFloat:=god;
                       form1.sdoxod.fieldbyname('sdoxod').asFloat:=0;
                       form1.sdoxod.fieldbyname('nalog').asFloat:=0;
                       form1.sdoxod.post;
                      // ShowMessage('Создание sdoxod');
                    end;

                  if (DRound((x-y)-(x2-y2),2)>xavans)  then     //полностью суммы хватает закрыть аванс
                    begin
                      form1.sdoxod.edit;
                      form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                      form1.sdoxod.fieldbyname('SDOXOD').asFloat:=xavans;
                      form1.sdoxod.post;
                     // ShowMessage('Полная выплата '+#13+form1.nacislName.Value+#13+floattostrf(xavans,ffNumber,12,2));
                      rtf:=true;
                      xavans:=0;
                     end
                      else
                     begin
                      form1.sdoxod.edit;
                      form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                      form1.sdoxod.fieldbyname('SDOXOD').asFloat:=x-y;
                      form1.sdoxod.fieldbyname('nalog').asFloat:=x2-y2;
                      form1.sdoxod.post;
                     // ShowMessage('Частичная выплата '+#13+form1.nacislName.Value+#13+floattostrf(x-y,ffNumber,12,2));
                      xavans:=xavans-((x-y)-(x2-y2));
                     end;

                     if Dround(xavans,2)=0 then rtf:=true;

               end;
            end;

          end;
        datam.qtmp2.next;
      end;

  end;
  datam.qtmp2.close;

  SmenaKodAvans:=rtf;

end;


function TForm_58.SmenaKodAvansFil(mes,god,nls:Real):Boolean;    //учет аванса по подразделениям
var xavans,x,y,z:Real;
           x2,y2:real;
  oldKod:Real;
  rtf:Boolean;
  xidnew,xIdvypl,xId:real;
  xDAt:TDate;
  ikod,j:integer;
  DAvans:array[0..9] of Real;
  oldActiv:Boolean;
begin

 //  ShowMEssage('avanas pod');

   //НЕ УЧИТЫВАЕТ GLNEW и ЧТО ДВА КОДА ОДИНАКОВЫХ НАЧИСЛЕНИЯ В РАЗНЫХ ПОДРАЗДЕЛЕНИЯХ

         if form1.sdoxod.locate('nls;mes;god;kodnac',VarArrayOf([nls,mes,god,0]),[loCaseInsensitive]) then
           begin
            xId:=form1.sdoxodID.VAlue; //запоминаем для удаления в конце если все ок
            xIdvypl:=form1.sdoxodidvypl.Value;
            xDat:=form1.sdoxodDAT.Value ;
            xAvans:=form1.sdoxodSDOXOD.Value-form1.sdoxodNALOG.Value;
           end
           else
             EXIT ;

   for ikod:=0 to 9 do DAvans[ikod]:=0;


   oldActiv:=form40.tavans.Active;
   if not form40.tavans.Active then form40.tavans.Active:=true;
   if form40.tavans.locate('nls',nls,[locaseinsensitive]) then
    begin
     for ikod:=0 to 9 do DAvans[ikod]:=Dround(xAvans*form40.tavans.fieldbyname('p'+inttostr(ikod)).asfloat/100,2);
    end;
   if oldactiv=false then form40.tavans.Active:=false;



    //
    rtf:=true;
    FOR ikod:=0 to 9 DO
    BEGIN
     IF DAvans[ikod]<>0 then
     BEGIN

      xAvans:=DAvans[ikod];
      datam.qtmp2.close;
      datam.qtmp2.Databasename:=form1.dbdir;
      datam.qtmp2.sql.clear;
      datam.qtmp2.sql.add('select o.kod,sum(o.kr),sum(o.snalog) from obrt1new o where o.wm='+floattostr(mes));
      datam.qtmp2.sql.add('and o.wg='+floattostr(god));
      datam.qtmp2.sql.add('and o.nls='+floattostr(nls));
      datam.qtmp2.sql.add('and o.filial='+floattostr(ikod));
      datam.qtmp2.sql.add('group by o.kod');
      datam.qtmp2.prepare;
      datam.qtmp2.open;

      //сначала проверка что хватает начислений с кодом 2000
      z:=0;
      datam.qtmp2.First;
       while not datam.qtmp2.eof do
        begin
         if form1.NACISL.Locate('kod',datam.qtmp2.fields[0].asFloat,[loCaseInsensitive]) then
           begin
            if (form1.NACISLKODDOX.Value='2000') or (form1.NACISLKODDOX.Value='2012')then
             begin
              x:=datam.qtmp2.Fields[1].asFloat;
              if form1.nacislRK.Value then x:=DRound(x,2)+DRound(x*form1.configRK.Value/100,2);

              datam.qtmp3.close;
              datam.qtmp3.Databasename:=form1.DBDIR;
              datam.qtmp3.sql.clear;
              datam.qtmp3.sql.add('select sum(sdoxod) from sdoxod where nls='+floattostr(nls));
              datam.qtmp3.sql.add('and mes='+floattostr(mes));
              datam.qtmp3.sql.add('and god='+floattostr(god));
              datam.qtmp3.sql.add('and kodnac='+floattostr(form1.nacislKod.Value));
              datam.qtmp3.prepare;
              datam.qtmp3.open;
              y:=datam.qtmp3.fields[0].asfloat;
              datam.qtmp3.close;

              z:=z+x-y;
            end;
          end;
        datam.qtmp2.next;
      end;

      if z<xavans then rtf:=false;
     END;
     END;

    IF NOT RTF THEN
      BEGIN
       MessageDlg(form1.kartFam.VAlue+' '+form1.kartIM.VAlue+' '+form1.kartOT.VAlue+#13
           +'Для распределения Аванса по подразделениям не хватает  начислений',mtInformation,[mbOk],0);
       SmenaKodAvansFil:=false;
       EXIT;
      END;

    //////////////
          datam.qtmp3.close;
          datam.qtmp3.databasename:=form1.dbdir;
          datam.qtmp3.sql.clear;
          datam.qtmp3.sql.add('select max(id) from sdoxod');
          datam.qtmp3.prepare;
          datam.qtmp3.Open;
          xidnew:=datam.qtmp3.Fields[0].asfloat;    //max id sdoxod
          datam.qtmp3.close;

  



 rtf:=true;
 FOR ikod:=0 to 9 DO
 BEGIN
  IF DAvans[ikod]<>0 then
  BEGIN
    rtf:=false;
    xAvans:=DAvans[ikod];
    datam.qtmp2.close;
    datam.qtmp2.Databasename:=form1.dbdir;
    datam.qtmp2.sql.clear;
    datam.qtmp2.sql.add('select o.kod,sum(o.kr),sum(o.snalog) from obrt1new o where o.wm='+floattostr(mes));
    datam.qtmp2.sql.add('and o.wg='+floattostr(god));
    datam.qtmp2.sql.add('and o.nls='+floattostr(nls));
    datam.qtmp2.sql.add('and o.filial='+floattostr(ikod));
    datam.qtmp2.sql.add('group by o.kod');
    datam.qtmp2.prepare;
    datam.qtmp2.open;

  //сначала проверка что хватает начислений с кодом 2000
  z:=0;
  datam.qtmp2.First;
  while not datam.qtmp2.eof do
   begin

      if form1.NACISL.Locate('kod',datam.qtmp2.fields[0].asFloat,[loCaseInsensitive]) then
       begin
        if (form1.NACISLKODDOX.Value='2000') or (form1.NACISLKODDOX.Value='2012') then
         begin
          x:=datam.qtmp2.Fields[1].asFloat;
          if form1.nacislRK.Value then x:=DRound(x,2)+DRound(x*form1.configRK.Value/100,2);

          datam.qtmp3.close;
          datam.qtmp3.Databasename:=form1.DBDIR;
          datam.qtmp3.sql.clear;
          datam.qtmp3.sql.add('select sum(sdoxod) from sdoxod where nls='+floattostr(nls));
          datam.qtmp3.sql.add('and mes='+floattostr(mes));
          datam.qtmp3.sql.add('and god='+floattostr(god));
          datam.qtmp3.sql.add('and kodnac='+floattostr(form1.nacislKod.Value));
          datam.qtmp3.prepare;
          datam.qtmp3.open;
          y:=datam.qtmp3.fields[0].asfloat;
          datam.qtmp3.close;

          z:=z+x-y;

         end;

       end;
     datam.qtmp2.next;
   end;

  //теперь замена если хватает


  if z>=xavans then
   begin

     datam.qtmp2.First;
     while not datam.qtmp2.eof do
      begin
         if form1.NACISL.Locate('kod',datam.qtmp2.fields[0].asFloat,[loCaseInsensitive]) then
          begin
           if (form1.NACISLKODDOX.Value='2000') or (form1.NACISLKODDOX.Value='2012') then
            begin
             x:=datam.qtmp2.Fields[1].asFloat;
             if form1.nacislRK.Value then x:=DRound(x,2)+DRound(x*form1.configRK.Value/100,2);
             x2:=datam.qtmp2.Fields[2].asFloat;
             datam.qtmp3.close;
             datam.qtmp3.Databasename:=form1.DBDIR;
             datam.qtmp3.sql.clear;
             datam.qtmp3.sql.add('select sum(sdoxod),sum(nalog) from sdoxod where nls='+floattostr(nls));
             datam.qtmp3.sql.add('and mes='+floattostr(mes));
             datam.qtmp3.sql.add('and god='+floattostr(god));
             datam.qtmp3.sql.add('and kodnac='+floattostr(form1.nacislKod.Value));
             datam.qtmp3.prepare;
             datam.qtmp3.open;
             y:=datam.qtmp3.fields[0].asfloat;
             y2:=datam.qtmp3.fields[1].asfloat;
             datam.qtmp3.close;

             if not rtf then   //
               begin

                       xidnew:=xidnew+1;
                       form1.sdoxod.Append;
                       form1.sdoxod.fieldbyname('id').asFloat:=xidnew;
                       form1.sdoxod.fieldbyname('tavans').asFloat:=1;
                       form1.sdoxod.fieldbyname('idvypl').asFloat:=xIdVypl;            //xIdvypl=id obrt2 !!!
                       form1.sdoxod.fieldbyname('nls').asFloat:=nls;
                       form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                       form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
                       form1.sdoxod.fieldbyname('mes').asFloat:=mes;
                       form1.sdoxod.fieldbyname('god').asFloat:=god;
                       form1.sdoxod.fieldbyname('sdoxod').asFloat:=0;
                       form1.sdoxod.fieldbyname('nalog').asFloat:=0;
                       form1.sdoxod.post;
                      // ShowMessage('Создание sdoxod');
                  

                  if (DRound((x-y)-(x2-y2),2)>xavans)  then     //полностью суммы хватает закрыть аванс
                    begin
                      form1.sdoxod.edit;
                      form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                      form1.sdoxod.fieldbyname('SDOXOD').asFloat:=xavans;
                      form1.sdoxod.post;
                     // ShowMessage('Полная выплата '+#13+form1.nacislName.Value+#13+floattostrf(xavans,ffNumber,12,2));
                      rtf:=true;
                      xavans:=0;
                     end
                      else
                     begin
                      form1.sdoxod.edit;
                      form1.sdoxod.fieldbyname('kodnac').asFloat:=form1.NACISLKOD.Value;
                      form1.sdoxod.fieldbyname('SDOXOD').asFloat:=x-y;
                      form1.sdoxod.fieldbyname('nalog').asFloat:=x2-y2;
                      form1.sdoxod.post;
                     // ShowMessage('Частичная выплата '+#13+form1.nacislName.Value+#13+floattostrf(x-y,ffNumber,12,2));
                      xavans:=xavans-((x-y)-(x2-y2));
                     end;

                     if Dround(xavans,2)=0 then rtf:=true;

               end;
            end;

          end;
        datam.qtmp2.next;
      end;

  end;
  datam.qtmp2.close;

 END;
 END;

 // все распределено удаляем
  if form1.sdoxod.locate('id',xId,[loCaseInsensitive]) then form1.sdoxod.delete;

  SmenaKodAvansFil:=true;

end;




procedure TForm_58.DelIdVyplNull(tNls:Real);     //удаляет sdoxod на недействующие ссылки в obrt2
var s:String;
begin
 datam.qTmp.Close;
 datam.qtmp.databasename:=form1.DBDIR;
 datam.qTmp.SQL.Clear;
 datam.qTmp.SQL.Add('select s.* from sdoxod s where s.god='+floattostr(RGod)+' and s.god>=2017');
 if tNls<>0 then datam.qTmp.SQL.Add('and s.nls='+floattostr(tNls));
 datam.qTmp.SQL.Add(' and s.idvypl>0 and (s.idvypl not in (select id from obrt2new))');
 datam.qTmp.Prepare;
 datam.qtmp.Open;
 datam.qTmp.first;
 while not datam.qtmp.eof do
  begin
     if form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asFloat,[loCaseinsensitive]) then s:=form1.nacislname.value else
        s:='<код начисления не найден>'+#13+floattostr(datam.qTmp.fieldbyname('kodnac').asFloat);

    form1.kart.locate('nls',datam.qtmp.fieldbyname('nls').asFloat,[loCaseInsensitive]) ;
    if MessageDlg(form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+
         'Обнаружена выплата дохода со ссылкой на недействуюущую выплату'+#13+
         'Доход='+floattostrf(datam.qtmp.fieldbyname('sdoxod').asFloat,ffNumber,12,2)+#13+
         'НДФЛ='+floattostrf(datam.qtmp.fieldbyname('nalog').asFloat,ffNumber,12,2)+#13+
         'Месяц='+floattostr(datam.qtmp.fieldbyname('mes').asFloat)+#13+
         'Код начисления='+s+#13+
         'Удалить данную выплату ?',mtError,[mbYes,mbNo],0) = mrYes then
           begin
            if form1.sdoxod.Locate('id',datam.qtmp.fieldbyname('id').asFloat, [loCaseInsensitive]) then
                  form1.sdoxod.delete;
           end;
   datam.qtmp.Next;
  end;;
 datam.qTmp.Close;


  datam.qTmp.close;
  datam.qTmp.sql.clear;
  datam.qTmp.sql.add('select s.id, s.sdoxod, s.nalog, s.god, s.mes, s.nls, s.kodnac from sdoxod s where  s.kodnac<>0');
  datam.qTmp.sql.add('and s.mes='+floattostr(RMes));
  datam.qTmp.sql.add('and s.god='+floattostr(RGod));
  if tNls<>0 then datam.qTmp.SQL.Add('and s.nls='+floattostr(tNls));
  datam.qTmp.sql.add('and (s.kodnac not in (select o.kod from obrt1new o where o.kod=s.kodnac and o.kr<>0 and o.wm=s.mes and o.wg=s.god and o.nls=s.nls))');
  datam.qTmp.prepare;
  datam.qTmp.open;

  datam.qTmp.first;
  while not datam.qTmp.Eof do
   begin
     form1.kart.locate('nls',datam.qTmp.fieldbyname('nls').asFloat,[loCaseinsensitive]);
     if form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asFloat,[loCaseinsensitive]) then s:=form1.nacislname.value else
        s:='<код начисления не найден>';

    if MessageDlg('Удалить несуществующую ссылку по выплате дохода'+#13+
       form1.kart.fieldbyname('fam').asString+' '+form1.kart.fieldbyname('im').asString+' '+form1.kart.fieldbyname('ot').asString+' '+#13+
         s+#13+
         'Доход='+floattostr(datam.qTmp.fieldbyname('sdoxod').asFloat)+#13+
                 'ндфл='+floattostr(datam.qTmp.fieldbyname('nalog').asFloat)+#13+
                 'год='+floattostr(datam.qTmp.fieldbyname('god').asFloat)+#13+
                 'месяц='+floattostr(datam.qTmp.fieldbyname('mes').asFloat)
                  ,mtWarning,[mbYes,mbNo],0) = mrYes then
              begin
               if form1.sdoxod.Locate('id',datam.qTmp.fieldbyname('id').asFloat, [loCaseInsensitive]) then
                  form1.sdoxod.delete;

              end;
    datam.qTmp.next;
   end;

  datam.qTmp.close;




end;



procedure Tform_58.FDatNacisl(tkoddox:String;wm,wg:Integer;datvyp:TDate; var datnacisl,datuderj,datperecisl:TDate);
var dd,mm,yy:Word;
begin
  datnacisl:=datvyp;
  datuderj:=datvyp;

 // ShowMessage(tkoddox+' '+floattostr(wm)+' '+floattostr(wg)+' '+datetostr(datvyp));

 if WG<=2022 THEN
  BEGIN
   if (tkoddox<>'2300')  and (tkoddox<>'2012')  then // мат. выгода
    begin
     datperecisl:=FPrazdnik(datvyp); //срок
    end
     else
    begin
     //отпускные и больничные срок, мат выгода последнее число месяца
     DecodeDAte(datvyp,yy,mm,dd);
     if mm<>12 then datperecisl:=EncodeDAte(yy,mm+1,1)-1 else datperecisl:=EncodeDate(yy,12,31);
     datperecisl:=FPrazdnik(datperecisl-1);
    end;

   if (tKoddox='2000') then       //оклад
    begin
     yy:=wg;
     mm:=wm;
     if mm<12 then datnacisl:=EncodeDate(yy,mm+1,1)-1 else datnacisl:=EncodeDate(yy,12,31);    //посл.число месяца
     if (datvyp<datnacisl) and (datvyp>EncodeDate(2000,1,1)) then
       begin
        datnacisl:=Datvyp; //если в течение месяца выплата то = дате выплаты
       end;
    end;

    if (tKoddox='2610' ) then  //мат.выгода
     begin
       yy:=wg;
       mm:=wm;
       if mm<12 then datnacisl:=EncodeDate(yy,mm+1,1)-1 else datnacisl:=EncodeDate(yy,12,31);    //посл.число месяца
       datuderj:=datnacisl;
       DecodeDAte(datvyp,yy,mm,dd);
       if mm<>12 then datperecisl:=EncodeDAte(yy,mm+1,1)-1 else datperecisl:=EncodeDate(yy,12,31);
       datperecisl:=FPrazdnik(datperecisl);
     end;
   END;

   IF datvyp>=EncodeDate(2023,1,1) THEN
    BEGIN
     datperecisl:=FGetDatPerecisl(datvyp);
    END;

end;


function Tform_58.FGetDatPerecisl(datvyp:TDate):TDate;
var dd,mm,yy:Word;
    datperecisl:TDate;
begin
     DecodeDAte(datvyp,yy,mm,dd);
     if (datvyp>=EncodeDate(yy,1,1)) and (datvyp<=EncodeDate(yy,1,22)) then datperecisl:=EncodeDate(yy,1,28);
     for mm:=1 to 11 do
      begin
       if (datvyp>=EncodeDate(yy,mm,23)) and (datvyp<=EncodeDate(yy,mm+1,22)) then datperecisl:=EncodeDate(yy,mm+1,28);
      end;
     datperecisl:=FPrazdnik2(datperecisl);
     if (datvyp>=EncodeDate(yy,12,23)) and (datvyp<=EncodeDate(yy,12,31)) then datperecisl:=FPrazdnikEndGod(datvyp);
     FGetDatPerecisl:=datperecisl;
end;



function Tform_58.FSetKPP(xoktmo:string):String;
var s:String;
begin
 s:='';
 if oktmo.Locate('oktmo',trim(xoktmo),[loCaseInsensitive]) then s:=oktmo.fieldbyname('kpp').asString;
 FSetKPP:=trim(s);

end;

function Tform_58.FSetGni(xoktmo:string):String;
var s:String;
begin
 s:='????';
 if oktmo.Locate('oktmo',trim(xoktmo),[loCaseInsensitive]) then s:=oktmo.fieldbyname('ifns').asString;
 FSetGni:=trim(s);

end;


procedure TForm_58.FormActivate(Sender: TObject);
begin
 if form55.NDFL then
  begin
   form_58.Visible:=False;
   JvXPButton2Click(Sender);
   HALT;
  end;
end;

procedure TForm_58.JvXPButton1Click(Sender: TObject);
begin
 form_58.Close;
end;

procedure TForm_58.JvXPButton2Click(Sender: TObject);
begin
 if rxCalcEdit1.Value=0 then
  begin
   MessageDlg('Введите номер справки',mtWarning,[mbOk],0);
   exit;
  end;
 form1.FSpr2006(1,Trunc(rxCalcEdit1.Value));

end;

procedure TForm_58.JvXPButton3Click(Sender: TObject);
var sGet,sGet2:String;
    i:integer;
    nSprMax:Real;
    oldDAte:TDAte;
begin

 RPriznak:=trim(Edit9.TExt);

       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select max(num) from spr2006 where GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       nSprMax:=datam.Query1.Fields[0].asFloat+1;
       datam.Query1.Close;

       if rxCalcEdit1.Value<nSprMax then
        begin
         MessageDlg('Начальный номер справок установлен '+floattostr(trunc(nSprMax))+#13+'т.к. обнаружены выгрузки справок',mtWarning,[mbOk],0);
         rxCalcEdit1.Value:=nSprMax;
        end;

if Length(Fac(Edit4.Text))<>4 then
   begin
    Edit4.SetFocus;
    exit;
   end;



 form2.CheckBox1.Checked:=true;
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 form2.DBGrid1.Columns[0].Visible:=False;
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;


{ sGet:=Trim(form2.ReadFS(GetCurrentDir()+form2.CodeStr('K^DOECQKGrerg~A9ecq'),100,150));
 sGet2:=Trim(form1.MDS(trim((form30.GetCRC())))) ;

 if (sGet<> sGet2) then
   begin
    MessageDlg('Данная операция доступна только зарегистрированным пользователям, '+#13+
     'См. <Помощь>  ->  <О программе>',mtWarning,[mbOk],0) ;
    exit;
   end;

}

 i:=0 ;
 form1.kart.first;
 while not form1.kart.Eof do
  begin
    if form1.kartG.Value='*' then i:=i+1;
   form1.kart.next;
  end;

 if i=0 then
  begin
   MessageDlg('Не выбраны сотрудники для формирования файла',mtWarning,[mbOk],0);
   exit;
  end;


 if MessageDlg('Сформировать файл со справками 2-НДФЛ в формате XML ?'+#13+
     'ОКТМО='+ComboBox1.Text+#13+
     'Получатель='+trim(edit6.text)+#13+
     'КПП='+trim(edit7.text)+#13+
     'Конечный получатель='+trim(Edit4.Text),mtInformation,[mbYes,mbNo],0) = mrNo then exit;
 if rxCalcEdit1.Value=0 then
  begin
   MessageDlg('Введите номер справки',mtWarning,[mbOk],0);
   exit;
  end;

 oldDate:=form_58.DateEdit1.Date;
 if Date()<=EncodeDate(2019,1,11) then form_58.DateEdit1.Date:=EncodeDate(2018,12,31);
 form1.FSpr2006(2,Trunc(rxCalcEdit1.Value));

 form_58.DateEdit1.Date:=oldDate;


end;

procedure TForm_58.V(Sender: TObject);
var i:integer;
    rtf:Boolean;
    _oldRMes:integer;
    k,k1,k2,npp:integer;
    st:Real;
begin

 if RGod>=2023 then Label2.Visible:=true else Label2.Visible:=false;

 _oldRMes:=RMes;
 if RMes<=3 then begin k1:=1; k2:=3; end;
 if (RMes>3) and (RMes<=6) then begin k1:=4 ; k2:=6; end;
 if (RMes>=7) and (RMes<=9) then begin k1:=7; k2:=9; end ;
 if (RMes>=10) and (RMes<=12) then begin k1:=10; k2:=12; end;
 for k:=k1 to k2 do
  begin
   RMes:=k;
   form1.kart.first;
     form102.RxLabel1.Caption:='Предварительная проверка '+ansilowercase(namemes[RMEs])   ;
     form102.ProgressBar1.Position:=0;
     form102.Show;
     form102.Refresh;
     npp:=0;
   while not form1.kart.eof do
    begin
       npp:=npp+1;
       form102.ProgressBar1.Position:=TRUNC(100*npp/form1.kart.RecordCount);
       form102.ProgressBar1.Refresh;
       form102.ProgressBar1.Repaint;
       if form1.kartSTATUS.Value='2' then st:=30 else st:=13;
       if not form1.Proverka2023(form1.kartnls.value,RMes,st) then
        begin
         //ShowMessage(form1.kartfam.value+#13+floattostr(RMes));
         form1.RaspredVicet;
        end;
     form1.kart.next;
    end;
  end;
 RMEs:=_oldRMEs;
 form102.close;

 if (RGOD=2022) then
  begin
  JvXpButton21.Visible:=false;
 //  if RGod=2022 then JvXpButton21.Visible:=true else JvXpButton21.Visible:=false;
  { MessageDlg('ВНИМАНИЕ ! ПЕРЕХОДНЫЙ ПЕРИОД ПО НДФЛ 2022->2023'+#13+
     'В случае, если выплата заработной платы за декабрь 2022 года'+#13+
      'частично или полностью была проведена в январе 2023 года,'+#13+'необходимо выполнить дополнительную обработку данных'+
         #13+'(Кнопка <Декабрь 2022>, при этом отчетный период должен быть установлен Декабрь 2022)',
       mtWarning,[mbOk],0);
   }
  end
   else
    begin
     JvXpButton21.Visible:=false;
    end;

 SKBK:='18210102010011000110'; 

  NUMLIST:=4;

  if RGod>=2018 then jvXPButton4.Visible:=false else jvXPButton4.Visible:=True;

  RKod:='';
  RInn:='';
  RKpp:='';
  RPRIZNAK:='1';
  { sprnum.DatabaseName:=form1.DBDIR;
   if not sprnum.Exists then sprnum.CreateTable;
   sprnum.Active:=True;
  }


   k6ndfl.DatabaseName:=form1.DBDIR;
   if not k6ndfl.Exists then k6ndfl.CreateTable;
   k6ndfl.Active:=True;


   Edit6.Text:=form1.config2GNI.Value;

     oktmo.DatabaseName:=form1.DBDIR;
     if oktmo.Active then oktmo.Active:=False;
     if not oktmo.Exists then oktmo.CreateTable;
     oktmo.Active:=true;

           if not oktmo.Locate('oktmo',form1.config2OKTMO.VAlue,[loCaseInsensitive]) then
             begin
              oktmo.append;
              oktmo.fieldbyname('oktmo').asString:=form1.config2OKTMO.Value;
              oktmo.fieldbyname('ifns').asString:=form1.config2GNI.Value;
              oktmo.fieldbyname('kpp').asString:=form1.config2KPP.Value;
              oktmo.post;
             end
              else
             begin
              if (trim(form1.config2KPP.Value)<>'')  and (trim(oktmo.fieldbyname('kpp').asString)='') then
                begin
                 oktmo.edit;
                 oktmo.fieldbyname('kpp').asString:=form1.config2KPP.Value;
                 oktmo.post;
                end;
             end ;





 if RMes<=3 then ComboBox2.ItemIndex:=0;
 if (RMes>3) and (RMes<=6) then ComboBox2.ItemIndex:=1;
 if (RMes>=7) and (RMes<=9) then ComboBox2.ItemIndex:=2;
 if (RMes>=10) and (RMes<=12) then ComboBox2.ItemIndex:=3;
 ComboBox3.ItemIndex:=0;
 if Length(Trim(form1.config2INN.Value))=12 then ComboBox3.ItemIndex:=0;
 if Length(Trim(form1.config2INN.Value))=10 then ComboBox3.ItemIndex:=5;  //214


 podpndfl.Active:=false;
 podpndfl.DatabaseName:=form1.DBDIR;
 if not podpndfl.Exists then podpndfl.CreateTable;
 podpndfl.Active:=true;
 if podpndfl.RecordCount=0 then
  begin
   podpndfl.append;
   podpndfl.fieldbyname('fio').asString:=form1.configRUKOVOD.Value;
   podpndfl.fieldbyname('agent').asFloat:=1;
   podpndfl.post;
   podpndfl.append;
   podpndfl.fieldbyname('fio').asString:=form1.configGLBUH.Value;
   podpndfl.fieldbyname('agent').asFloat:=1;
   podpndfl.post;
  end;


 if Trim(form1.config2OKTMO.Value)='' then
  begin
   MessageDlg('В настройках организации не заполнено поле КОД ОКТМО',mtWarning,[mbOk],0);
   exit;
  end;

  ComboBox1.Items.Clear;
  ComboBox1.Items.Add(Trim(form1.config2OKTMO.Value));
  form1.kart.first;
  while not form1.kart.eof do
   begin
     datam.kart2.Locate('nls',form1.kartNls.Value,[loCaseInsensitive]);
     if trim(datam.kart2OKTMO.Value)<>'' then
      begin
       rtf:=false;
       for i:=0 to ComboBox1.Items.Count do
        begin
         if ComboBox1.Items[i]=trim(datam.kart2OKTMO.Value) then rtf:=true;
        end;
       if not rtf then CombObox1.Items.Add(trim(datam.kart2OKTMO.Value));
      end;

    form1.kart.next;
   end;

   datam.qtmpstaj.close;
   datam.qtmpstaj.DatabaseName:=form1.dbdir;
   datam.qtmpstaj.sql.Clear;
   datam.qtmpstaj.sql.add('select * from oktmonls');
   datam.qtmpstaj.Prepare;
   datam.qtmpstaj.Open;
   datam.qtmpstaj.first;
   while not datam.qtmpstaj.eof do
    begin
     if trim(datam.qtmpstaj.fieldbyname('OKTMO').asString)<>'' then
      begin
       rtf:=false;
       for i:=0 to ComboBox1.Items.Count do
        begin
         if ComboBox1.Items[i]=trim(datam.qtmpstaj.fieldbyname('OKTMO').asString) then rtf:=true;
        end;
       if not rtf then CombObox1.Items.Add(trim(datam.qtmpstaj.fieldbyname('OKTMO').asString));
      end;
    datam.qtmpstaj.next;
   end;

   datam.qtmpstaj.close;


  form1.kart.first;
  ComboBox1.ItemIndex:=0;
  if ComboBox1.Items.Count>1 then
   begin
    MessageDlg('Внимание ! Обнаружены сотрудники с ОКТМО отличным от ОКТМО организации',mtInformation,[mbOk],0);
   end;

  ComboBox4.Items:=ComboBox1.Items;
  ComboBox4.ItemIndex:=ComboBox1.ItemIndex;

  Edit4.Text:=FSetGni(ComboBox1.Text);
  Edit5.Text:=FSetGni(ComboBox4.Text);
  Edit7.Text:=FSetKPP(ComboBox1.Text);
  Edit8.Text:=FSetKPP(ComboBox4.Text);

  if RGod>=2023 then
   begin
    jvXpButton20.Visible:=false;
    jvXpButton13.Visible:=false;
    
   end;


end;

procedure TForm_58.FormCreate(Sender: TObject);
begin
 DateEdit1.Date:=Date();
 DateEdit2.Date:=Date();
 ndflr5.DatabaseName:=form1.DBDIR;
 if not ndflr5.Exists then ndflr5.CreateTable;
 ndflr5.Active:=true;

 ndfl6.DatabaseName:=form1.DBDIR;
 if not ndfl6.Exists then ndfl6.CreateTable;
 ndfl6.Active:=true;
 ndfl6ob.DatabaseName:=form1.DBDIR;
 if not ndfl6ob.Exists then ndfl6ob.CreateTable;
 ndfl6ob.Active:=true;

 reorg.DatabaseName:=form1.DBDIR;
 if not reorg.Exists then reorg.CreateTable;
 reorg.Active:=true;


end;

procedure TForm_58.JvXPButton5Click(Sender: TObject);
begin
 if MessageDlg('Создать запись о подписанте',mtInformation,[mbYes,mbNo],0) = mrNo then exit;
 podpndfl.Append;
 podpndfl.FieldByName('agent').asFloat:=1;
 podpndfl.post;
 form617:=TForm617.Create(nil);
 form617.ShowModal;
 podpndfl.edit;
 podpndfl.post;
 form617.free;

end;

procedure TForm_58.JvXPButton6Click(Sender: TObject);
begin
 if podpndfl.RecordCount<=0 then exit;
 if MessageDlg('Удалить подписанта',mtWarning,[mbYes,mbNo],0) = mrYes then podpndfl.delete;
end;

procedure TForm_58.JvXPButton7Click(Sender: TObject);
begin
 if podpndfl.RecordCount<=0 then exit;
 form617:=TForm617.Create(nil);
 form617.ShowModal;
 podpndfl.edit;
 podpndfl.post;
 form617.free;
end;


function TForm_58.FPrazdnikEndGod(xDat:TDate):TDate ;  //последний рабочий день года, для срока уплаты с 23.12 по 31.12
var i:Integer;
    dd,mm,yy:Word;
    rtf:Boolean;
    xDat3:TDate;
begin
 rtf:=false;
 DecodeDate(xDat,yy,mm,dd);
 xDat3:=EncodeDate(yy,12,31);
 for i:=31 downto 15 do   //15подряд праздники не может быть
     begin
       if not rtf then
        begin
         if form1.kalend1.Locate('TYPE;MES;GOD',VarArrayOf([5,12,yy]),[loCaseInsensitive]) then
          begin
            if form1.kalend1.FieldByName('N'+IntToStr(i)).asFloat<>0 then
             begin
              xDat3:=EncodeDate(yy,12,i);
              rtf:=true;
             end;
           end;
       end;
     end;
 FPrazdnikEndGod:=xDat3;
end;

function TForm_58.FPrazdnik2(xDat:TDate):TDate ; //рабочий день даты xDat, если выходной - то первый рабочий следующий
var i:Integer;
    dd,mm,yy:Word;
    rtf:Boolean;
    xDat2,xDat3:TDate;
begin
 rtf:=false;
 xDat2:=xDat;
 xDat3:=xDat;
 for i:=1 to 15 do   //15подряд праздники не может быть
     begin
       if not rtf then
        begin
         DecodeDAte(xDat2,yy,mm,dd);
         if form1.kalend1.Locate('TYPE;MES;GOD',VarArrayOf([5,mm,yy]),[loCaseInsensitive]) then
          begin
            if form1.kalend1.FieldByName('N'+IntToStr(dd)).asFloat<>0 then
             begin
              xDat3:=xDat2;
              rtf:=true;
             end;
           end;
       end;
      xDat2:=xDat2+1;
     end;

 if not rtf then xDat3:=xDat+1;
 FPrazdnik2:=xDat3;

end;


function TForm_58.FPrazdnik(xDat:TDate):TDate ; //первый рабочий день после даты xDat
var i:Integer;
    dd,mm,yy:Word;
    rtf:Boolean;
    xDat2,xDat3:TDate;
begin
 rtf:=false;
 xDat2:=xDat;
 xDat3:=xDat;
 for i:=1 to 15 do   //15подряд праздники не может быть
     begin
       xDat2:=xDat2+1;
       if not rtf then
        begin
         DecodeDAte(xDat2,yy,mm,dd);
         if form1.kalend1.Locate('TYPE;MES;GOD',VarArrayOf([5,mm,yy]),[loCaseInsensitive]) then
          begin
            if form1.kalend1.FieldByName('N'+IntToStr(dd)).asFloat<>0 then
             begin
              xDat3:=xDat2;
         //     ShowMessage(datetostr(xdat3));
              rtf:=true;
             end;
           end;
         end;
     end;

 if not rtf then xDat3:=xDat+1;
 FPrazdnik:=xDat3;

end;

procedure TForm_58.JvXPButton8Click(Sender: TObject);
var FNameXLS:String;
    E:OleVAriant;
    s,sm,xFam,xIm,xOt:String;
    sDox,sDox9,sVycet,sfiks,sgpx,sKolvo,s115,s121,s142,s155,sarenda:array[1..3] of Real;
    sNdfl,sNdfl9:array[1..3] of Real;
    x,sSumma0,sNdfl0:Real;
    i,j:integer;
    rtf3:Boolean;
    PDEC2022:Real;
    ds:String;
    x2,xNdfl:Real;
    dd,mm,yy:Word;
    tDAt,tDat2,tDat3:TDate;
    st,nList:integer;
    Nz0,Nz:Integer;
    WOK,WOK5:Integer;
    idGuid:String;
    XMLFileName:String;
    rtfFL:Boolean;
    sneuderj,suderj,svozvr,slisnee:array[1..3] of Real;
    TLoad:Boolean;
    datstart,datend:TDate;
    sKpp:String;
    tNdflUderj:array[1..3] of Real;
    bRTF:Boolean;
    np1:Integer;
    wdat1,wdat2:TDate;
    TOK, npp:Integer;
    bRtf2:Boolean;
    xItog:Real;
    rtfgpx:boolean;
    RTFI :Boolean;
begin

   if RGod=2022 then form_58.IsprDat31122022;

   form85:=tform85.create(nil);
   form85.DelErrOklad;//удаляет ошибочные записи sdoxod для которых doxod=0 and nalog<>0
   form85.free;



   ComboBox1.Text:=ComboBox4.Text;
   Edit6.Text:=Edit5.Text;
   Edit7.Text:=Edit8.Text;



                      


    xMes1:=1;
    xMes2:=12; //период !!!!!!
    if ComboBox2.ItemIndex=0 then xMes2:=3;
    if ComboBox2.ItemIndex=1 then xMes2:=6;
    if ComboBox2.ItemIndex=2 then xMes2:=9;
    if ComboBox2.ItemIndex=3 then xMes2:=12;


   

  if Length(Fac(Edit5.Text))<>4 then
   begin
    Edit5.SetFocus;
    exit;
   end;
   TLoad:=false;
   if ndfl6.locate('god;mes;oktmo',VarArrayOf([RGod,xMes2,trim(ComboBox4.Text)]),[loCaseInsensitive]) then
    begin
        form650:=TForm650.Create(nil);
        form650.TGod:=RGod;
        form650.TMes:=xMes2;
        form650.TOktmo:=trim(ComboBox4.Text);

        form650.ShowModal;
        TLoad:=form650.TOk; //найдено и грузим
       
        if form650.TID<0 then
         begin
           form650.free;
           EXIT;
         end;
        ndfl6.Locate('id',form650.TId,[loCaseInsensitive]);
        form650.Free;



    end;

 skpp:=trim(Edit8.Text);

 if not TLoad then
  begin

                         bRTF:=True;

                         datam.qTmp.Close;
                         datam.qTmp.DatabaseName:=form1.DBDIR;
                         datam.qTmp.SQL.Clear;
                         datam.qTmp.SQL.Add('select g.*, k.fam,k.im,k.ot from glnew g, kart k, sdoxod s where s.nls=g.nls and k.nls=g.nls ');
                         datam.qTmp.SQL.Add(' and s.sdoxod<>0 and g.wm<='+floattostr(Rmes));
                         datam.qTmp.SQL.Add(' and g.wm=s.mes and s.kodnac=0');
                         datam.qTmp.SQL.Add(' and g.wg='+floattostr(rGod));
                         datam.qTmp.SQL.Add(' and s.god='+floattostr(rGod));
                //         datam.qTmp.SQL.Add(' and g.nls='+floattostr(vyplzpNls.Value));
                         datam.qTmp.SQL.Add(' and g.oklad*g.dayotr=0');
                         datam.qTmp.prepare;
                         datam.qTmp.open;
                         datam.qTmp.First;
                         while not datam.qTmp.eof do
                          begin
                             MessageDlg('Внутренняя ошибка таблицы glnew'+#13+
                                     namemes[datam.qTmp.fieldbyname('wm').asInteger]+#13+
                                       datam.qTmp.fieldByname('FAM').asString+' '+Copy(datam.qTmp.fieldByname('IM').asString,1,1)+'.'
                                         +Copy(datam.qTmp.fieldByname('OT').asString,1,1)+#13+
                                          'Выполните СЕРВИС - ВНУТРЕННИЙ ТЕСТ СОСТОЯНИЯ ТАБЛИЦЫ GLNEW в главном окне программы'+#13+
                                   'и после этого проверьте начисление оклада/тарифа в данном периоде у данного сотрудника'
                                          ,mtError,[mbOk],0);
                                      exit;
                            bRTF:=False;
                           datam.qTmp.next;
                          end;
                         datam.qTmp.close;
                         if not bRTF then
                           begin
                           // exit;
                           end;


   if MessageDlg('Раздел 2: '+RadioGroup1.Items[RadioGroup1.ItemIndex]+#13+'ОКТМО='+ComboBox4.Text+#13+'Код ИФНС получатель='+Trim(Edit5.Text)+
     #13+'КПП='+sKpp+
        #13+'Продолжить формирование ?',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

    Button1Click(nil);


  end;



    xMes21:=xMes2-2;

    datstart:=EncodeDate(RGod,xMes21,1);
    if xMes2<>12 then datend:=EncodeDate(RGod,xMes2+1,1)-1 else datend:=EncodeDate(RGod,12,31);



           ds:=form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER)+'.dbf';
           if tmprepdox.Active then tmprepdox.Active:=False;
           if tmprepdox.Exists then tmprepdox.DeleteTable;
           tmprepdox.TableName:=ds;
           tmprepdox.DatabaseName:=form1.DBDIR;
           tmprepdox.Exclusive:=True;
           tmprepdox.TableType:=ttDBase;
           tmprepdox.FieldDefs.Clear;
           tmprepdox.FieldDefs.Add('summa',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('id2',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('ndfl',ftFloat,0,false); 
           tmprepdox.FieldDefs.Add('dat',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat2',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat3',ftDate,0,false);
           tmprepdox.CreateTable;
           Reindex.ReindexTab(tmprepdox,form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER),'dat','');
           tmprepdox.IndexName:='dat';


    // распределение НДФЛ

 IF not TLoad then
  BEGIN
   form102.RxLabel1.Caption:='Заполнение регистров НДФЛ по сотрудникам'   ;
   form102.ProgressBar1.Position:=0;
   form102.Show;
   form102.Refresh;
   Nz0:=form1.kart.RecordCount;
   Nz:=0;

    form1.kart.first;
    for i:=1 to 3 do
     begin
      sDox[i]:=0; sVycet[i]:=0;sDox9[i]:=0;sKolvo[i]:=0; sNdfl[i]:=0; sNdfl9[i]:=0; sfiks[i]:=0; sgpx[i]:=0; sarenda[i]:=0;
      sneuderj[i]:=0;svozvr[i]:=0; suderj[i]:=0; slisnee[i]:=0;
      s115[i]:=0;s121[i]:=0;s142[i]:=0;s155[i]:=0;
     end;



    for i:=1 to 3 do tNdflUderj[i]:=0;
    while not form1.kart.Eof do
     begin

       nZ:=nZ+1;
       form102.ProgressBar1.Position:=Trunc(100*Nz/Nz0);
       form102.ProgressBar1.Refresh;

         TOK:=ZapolnDOK(form1.kartNLS.Value);

         rtf3:=true;
         datam.kart2.locate('nls',form1.kartNls.Value,[loCaseInsensitive]);

         IF TOK=0 THEN   //нет записей в базе oktmo по сотруднику
          BEGIN
           if (trim(datam.kart2OKTMO.Value)='') and (trim(form1.config2OKTMO.Value)<>form_58.ComboBox4.Text) then rtf3:=false;
           if (trim(datam.kart2OKTMO.Value)<>'') and (trim(datam.kart2OKTMO.Value)<>form_58.ComboBox4.Text) then rtf3:=false;
          END;

      IF (rtf3) and (RGod>=2023) THEN
       BEGIN

          ObrabObrtNalKart;
          NewSpr2023;
          st:=1;
          rtfFL:=false;
          if form1.kartSTATUS.Value='2' then st:=2  ;

          for j:=1 to 10 do for i:=1 to 12 do    //вычеты с мат.помощи
           begin
            if (qqD[j]<>'') and (qqDV[j,i]<>0) then sVycet[st]:=sVycet[st]+qqDV[j,i];
           end;



           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.DatabaseName:=form1.DBDIR;
           datam.Query1.SQL.Add('select * from sdoxod where');
           datam.Query1.SQl.add('nls='+floattostr(form1.kartNls.Value));
           datam.Query1.SQl.add('and dat>='+#39+formatdatetime('dd.mm.yyyy',Encodedate(RGod,1,1))+#39);
           datam.Query1.SQl.add('and dat<='+#39+formatdatetime('dd.mm.yyyy',EncodeDate(RGod,12,31))+#39);
           datam.query1.prepare;
           datam.query1.open;
           datam.Query1.first;
           while not datam.Query1.Eof do
            begin
              bRTF:=true;
              tdat:=datam.Query1.Fieldbyname('dat').asDatetime;
              tDat2:=tDAt;
              tdat3:=form_58.FGetDatPerecisl(tdat);
              if datam.Query1.fieldbyname('kodnac').asfloat<>0 then
                begin
                 form1.NACISL.Locate('kod',datam.Query1.fieldbyname('kodnac').asInteger,[loCaseInsensitive]);
                 if form1.NACISLPN.Value=1 then bRTF:=False;
                end;

              if TOK=0 then
               begin
                x2:=datam.Query1.Fieldbyname('sdoxod').asFloat;
                xNdfl:=datam.Query1.Fieldbyname('nalog').asFloat;
               end
                else
               begin
                if FOktmo(datam.Query1.Fieldbyname('mes').asinteger,datam.Query1.Fieldbyname('god').asinteger,form_58.ComboBox4.Text) then
                  begin
                   x2:=datam.Query1.Fieldbyname('sdoxod').asFloat;
                   xNdfl:=datam.Query1.Fieldbyname('nalog').asFloat;
                  end
                   else
                  begin
                   x2:=0;
                   xNdfl:=0;
                   bRTF:=false;
                  end;
               end;


              if datam.Query1.fieldbyname('god').asfloat<=2022 then  //аванс декабрь 2022 и если з/п выплачена в 2023 за декабрь
                begin
                  if (bRTF) and (datam.query1.FieldByName('tavans').asfloat=1) then
                    begin
                     tNdflUderj[st]:=tNdflUderj[st]+datam.Query1.Fieldbyname('nalog').asFloat;

                     if (tDat3>=datstart) and (tDat3<=datend) then
                       begin
                         if not tmprepdox.Locate('dat;dat2;dat3',VarArrayOf([tDAt,tDAt2,tDAt3]),[loCaseInsensitive]) then
                           begin
                             // Showmessage('Добавлено Раздел1 '+datetostr(tDat)+#13+datetostr(tDat2)+#13+datetostr(tDat3));
                            tmprepdox.Append;
                            tmprepdox.FieldByName('dat').asDateTime:=tDAt;
                            tmprepdox.FieldByName('dat2').asDateTime:=tDAt2;
                            tmprepdox.FieldByName('dat3').asDateTime:=tDAt3;
                            tmprepdox.post;
                           end;
                           tmprepdox.edit;
                           tmprepdox.FieldByName('summa').asFloat:=tmprepdox.FieldByName('summa').asFloat+x2;
                           tmprepdox.FieldByName('ndfl').asFloat:=tmprepdox.FieldByName('ndfl').asFloat+xndfl;
                           tmprepdox.post;
                       end;
                    end;
                  if datam.query1.FieldByName('tavans').asfloat=1 then bRTF:=false;
                end;


              IF (bRTF) and (tdat>=EncodeDate(RGod,1,1)) and (tdat<=datend) then
               BEGIN
                 rtfFL:=true;
                 sVycet[st]:=sVycet[st]+datam.Query1.Fieldbyname('rvicet').asFloat;
                 if (datam.kart2STATUS2.value=7) or (datam.kart2STATUS2.value=3) then
                  begin
                    s115[st]:=s115[st]+datam.Query1.Fieldbyname('sdoxod').asFloat;  //ВКС
                    s142[st]:=s142[st]+datam.Query1.Fieldbyname('nalog').asFloat;  //ВКС
                  end;
                 sDox[st]:=sDox[st]+datam.Query1.Fieldbyname('sdoxod').asFloat;
                 sNdfl[st]:=sNdfl[st]+datam.Query1.Fieldbyname('nalog').asFloat;
                 tNdflUderj[st]:=tNdflUderj[st]+datam.Query1.Fieldbyname('nalog').asFloat;

                 if datam.query1.fieldbyname('kodnac').asfloat<>0 then
                  begin
                   if Trunc(form1.NACISLSTPNALOG.Value)=9 then //дивиденды
                    begin
                     sDox9[st]:=sDox9[st]+datam.Query1.Fieldbyname('sdoxod').asFloat;
                     sNdfl9[st]:=sNdfl9[st]+datam.Query1.Fieldbyname('nalog').asFloat;
                     sDox[st]:=sDox[st]-datam.Query1.Fieldbyname('sdoxod').asFloat;     //дважды учитывается
                     sNdfl[st]:=sNdfl[st]-datam.Query1.Fieldbyname('nalog').asFloat;    //дважды учитывается
                    end;
                   if form1.NACISLKODDOX.Value='2010' then sGpx[st]:=sGpx[st]+datam.Query1.Fieldbyname('sdoxod').asFloat;

                   //аренда входит только в общий доход, в труд., гпх не входит
                   if (form1.NACISLKODDOX.Value='1400') or (form1.NACISLKODDOX.Value='1401') or (form1.NACISLKODDOX.Value='1402')
                        or (form1.NACISLKODDOX.Value='2400') then sArenda[st]:=sArenda[st]+datam.Query1.Fieldbyname('sdoxod').asFloat;


                  end;

                 // Showmessage('Сумма '+floattostr(tmprepdox.FieldByName('summa').asFloat)+#13+floattostr(tmprepdox.FieldByName('ndfl').asFloat));
                END;


               IF  (bRtf) and (tDat3>=datstart) and (tDat3<=datend) then
                  BEGIN
                   if not tmprepdox.Locate('dat;dat2;dat3',VarArrayOf([tDAt,tDAt2,tDAt3]),[loCaseInsensitive]) then
                     begin
                    // Showmessage('Добавлено '+datetostr(tDat)+#13+datetostr(tDat2)+#13+datetostr(tDat3));
                      tmprepdox.Append;
                      tmprepdox.FieldByName('dat').asDateTime:=tDAt;
                      tmprepdox.FieldByName('dat2').asDateTime:=tDAt2;
                      tmprepdox.FieldByName('dat3').asDateTime:=tDAt3;
                      tmprepdox.post;
                     end;
                     tmprepdox.edit;
                     tmprepdox.FieldByName('summa').asFloat:=tmprepdox.FieldByName('summa').asFloat+x2;
                     tmprepdox.FieldByName('ndfl').asFloat:=tmprepdox.FieldByName('ndfl').asFloat+xndfl;
                     tmprepdox.post;
                   END;

             datam.Query1.Next;
            end;

          if rtfFL then
            begin
             Skolvo[st]:=sKolvo[st]+1;
             if (datam.kart2STATUS2.value=7) or (datam.kart2STATUS2.value=3) then s121[st]:=s121[st]+1;
            end;
       END;


     IF (rtf3) and (RGod<=2022) THEN
      BEGIN

           ObrabObrtNalKart;

           st:=1;
           if form1.kartSTATUS.Value='2' then st:=2  ;

            x:=form_58.FGetUderjNdfl(form1.kartNls.Value,datend);

             if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
               begin
                if ndflr5PR2.Value=1 then
                 begin
                  if FOktmo(1,RGod,form_58.ComboBox4.Text) then   //конечно неправильно январь смотрим в дальнейшем в ndfl5 сделать oktmo Поле
                              x:=DRound(ndflr5.FieldByName('sud').asFloat,2);
                 end;
               end  ;

            tNdflUderj[st]:=tNdflUderj[st]+x ;

         
           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.DatabaseName:=form1.DBDIR;
           datam.Query1.SQL.Add('select * from sdoxod where');
           datam.Query1.SQl.add('nls='+floattostr(form1.kartNls.Value));
           datam.Query1.SQl.add('and dat>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
           datam.Query1.SQl.add('and dat<='+#39+formatdatetime('dd.mm.yyyy',datend)+#39);
           datam.query1.prepare;
           datam.query1.open;
           datam.Query1.first;
           while not datam.query1.eof do
            begin
              if datam.Query1.fieldbyname('kodnac').asfloat=0 then
                   FDatnacisl('2000',datam.Query1.Fieldbyname('mes').asInteger,datam.Query1.Fieldbyname('god').asInteger,
                                                       datam.Query1.FieldByNAme('dat').asDateTime,tdat,tDat2,tDat3)
                              else
                          begin
                           form1.NACISL.Locate('kod',datam.Query1.fieldbyname('kodnac').asfloat,[loCaseInSensitive]);
                           FDatnacisl(form1.NACISLKODDOX.Value,datam.Query1.Fieldbyname('mes').asInteger,datam.Query1.Fieldbyname('god').asInteger,
                                                       datam.Query1.FieldByNAme('dat').asDateTime,tdat,tDat2,tDat3)
                          end;
              bRTF:=true;
              if datam.Query1.fieldbyname('kodnac').asfloat<>0 then
                begin
                 form1.NACISL.Locate('kod',datam.Query1.fieldbyname('kodnac').asInteger,[loCaseInsensitive]);
                 if form1.NACISLPN.Value=1 then bRTF:=False;
                end;


              wdat1:=EncodeDate(datam.query1.fieldbyname('god').asInteger,datam.query1.fieldbyname('mes').asInteger,1);
              if datam.query1.fieldbyname('mes').asInteger<>12 then
                 wdat2:=EncodeDate(datam.query1.fieldbyname('god').asInteger,datam.query1.fieldbyname('mes').asInteger+1,1)-1
                   else
                    wdat2:=EncodeDate(datam.query1.fieldbyname('god').asInteger,12,31) ;

              if (RGod<=2022) and (datam.query1.FieldByName('tavans').asInteger=1) and (datam.query1.fieldbyname('dat').asDatetime<wdat2)
                 and ( datam.query1.fieldbyname('dat').asDatetime>=wdat1)
                    and (datam.query1.fieldbyname('dat').asDatetime<>EncodeDate(2022,12,30))  then
                   begin
                    //аванс не выплачен доход
                    if MessageDlg('Обнаружен не выплаченный доход по которому проведен аванс'+#13+form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+
                     #13+'Дата аванса: '+Formatdatetime('dd.mm.yyyy',datam.query1.fieldbyname('dat').asDatetime)+#13+
                       Floattostr(datam.query1.fieldbyname('sdoxod').asFloat)+' руб.'+#13+
                       'Месяц: '+datam.query1.fieldbyname('mes').asString+#13+
                      'Невозможно определить дату получения дохода на момент формирования отчета'+#13+
                      'Включать в Раздел 2 данную выплату ?',mtWarning,[mbYes,mbNo],0) = mrNo then bRtf:=False;

                   end;


              if TOK=0 then
               begin
                x2:=datam.Query1.Fieldbyname('sdoxod').asFloat;
                xNdfl:=datam.Query1.Fieldbyname('nalog').asFloat;
               end
                else
               begin
                if FOktmo(datam.Query1.Fieldbyname('mes').asinteger,datam.Query1.Fieldbyname('god').asinteger,form_58.ComboBox4.Text) then
                  begin
                   x2:=datam.Query1.Fieldbyname('sdoxod').asFloat;
                   xNdfl:=datam.Query1.Fieldbyname('nalog').asFloat;
                  end
                   else
                  begin
                   x2:=0;
                   xNdfl:=0;
                   bRTF:=false;
                  end;
               end;
               

              IF bRTF then
               BEGIN
                if not tmprepdox.Locate('dat;dat2;dat3',VarArrayOf([tDAt,tDAt2,tDAt3]),[loCaseInsensitive]) then
                  begin
                 // Showmessage('Добавлено '+datetostr(tDat)+#13+datetostr(tDat2)+#13+datetostr(tDat3));
                   tmprepdox.Append;
                   tmprepdox.FieldByName('dat').asDateTime:=tDAt;
                   tmprepdox.FieldByName('dat2').asDateTime:=tDAt2;
                   tmprepdox.FieldByName('dat3').asDateTime:=tDAt3;
                   tmprepdox.post;
                  end;
                  tmprepdox.edit;
                  tmprepdox.FieldByName('summa').asFloat:=tmprepdox.FieldByName('summa').asFloat+x2;
                  tmprepdox.FieldByName('ndfl').asFloat:=tmprepdox.FieldByName('ndfl').asFloat+xndfl;
                  tmprepdox.post;
                 // Showmessage('Сумма '+floattostr(tmprepdox.FieldByName('summa').asFloat)+#13+floattostr(tmprepdox.FieldByName('ndfl').asFloat));
                END;
             datam.query1.next;
            end;

           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.DatabaseName:=form1.DBDIR;
           datam.Query1.SQL.Add('select dayrab,datoklad,dayotr,oklad,daycas,snalog,wm,wg,nls from glnew ' );
           datam.Query1.SQl.add('where dayotr<>0');
           if RadioGroup1.ItemIndex=0 then       //по дате выплаты
            begin
             datam.Query1.SQl.add('and datoklad>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
             datam.Query1.SQl.add('and datoklad<='+#39+formatdatetime('dd.mm.yyyy',datend)+#39);
            end;
           if RadioGroup1.ItemIndex=1 then       //по дате начисления
            begin
             datam.Query1.SQl.add('and wm>='+floattostr(xMes21));
             datam.Query1.SQl.add('and wm<='+floattostr(xMes2));
             datam.Query1.SQL.Add('and wg='+floattostr(RGod)) ;
            end;
           datam.Query1.SQl.add('and nls='+floattostr(form1.kartNls.Value));
           datam.Query1.Prepare;
           datam.Query1.Open;
           datam.Query1.First;
           while not datam.Query1.Eof do
            begin
             x2:=0;
               if datam.Query1.FieldByName('dayrab').asFloat<>0 then
                x2:=mainlib.DRound(datam.Query1.FieldByName('oklad').asFloat*
                  DelenieCas(datam.Query1.FieldByName('dayotr').asFloat,
                    datam.Query1.FieldByName('dayrab').asFloat,datam.Query1.FieldByName('DAYCAS').asFloat),2)  ;
                x2:=x2+DRound(x2*form1.configRK.Value/100,2);
                xNdfl:=datam.Query1.Fieldbyname('snalog').asFloat;

              FDatnacisl('2000',datam.Query1.Fieldbyname('wm').asInteger,datam.Query1.Fieldbyname('wg').asInteger,datam.Query1.FieldByNAme('datoklad').asDateTime,tdat,tDat2,tDat3) ; //исч


             if RadioGroup1.ItemIndex=0 then
              begin
               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac=0 and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;
               end;


              //
              bRTf2:=true;
              if not FOktmo(datam.Query1.Fieldbyname('wm').asinteger,datam.Query1.Fieldbyname('wg').asinteger,form_58.ComboBox4.Text) then
                  begin
                   x2:=0;
                   xNdfl:=0;
                   bRTf2:=false;
                  end;
              //

             if bRtf2 then
              begin
               if not tmprepdox.Locate('dat;dat2;dat3',VarArrayOf([tDAt,tDAt2,tDAt3]),[loCaseInsensitive]) then
                begin
                 tmprepdox.Append;
                 tmprepdox.FieldByName('dat').asDateTime:=tDAt;
                 tmprepdox.FieldByName('dat2').asDateTime:=tDAt2;
                 tmprepdox.FieldByName('dat3').asDateTime:=tDAt3;
                 tmprepdox.post;
                end;
                tmprepdox.edit;
                tmprepdox.FieldByName('summa').asFloat:=tmprepdox.FieldByName('summa').asFloat+x2;
                tmprepdox.FieldByName('ndfl').asFloat:=tmprepdox.FieldByName('ndfl').asFloat+xndfl;
                tmprepdox.post;
              end;
              
             datam.Query1.Next;
            end;

           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.SQL.Add('select n.koddox,n.rk, o.* from nacisl n, obrt1new o where o.kod=n.kod');
            if RadioGroup1.ItemIndex=0 then       //по дате выплаты
            begin
             datam.Query1.SQl.add('and o.datprov>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
             datam.Query1.SQl.add('and o.datprov<='+#39+formatdatetime('dd.mm.yyyy',datend)+#39);
            end;
           if RadioGroup1.ItemIndex=1 then       //по дате начисления
            begin
             datam.Query1.SQl.add('and o.wm>='+floattostr(xMes21));
             datam.Query1.SQl.add('and o.wm<='+floattostr(xMes2));
             datam.Query1.SQL.Add('and o.wg='+floattostr(RGod)) ;
            end;
           datam.Query1.SQl.add('and o.nls='+floattostr(form1.kartNls.Value));
           datam.Query1.SQl.add('and n.pn<>1');
           datam.Query1.Prepare;
           datam.Query1.Open;
           datam.Query1.First;
           while not datam.Query1.Eof do
            begin
             x2:=datam.Query1.FieldByName('KR').asFloat;
             if datam.Query1.FieldByName('RK').asBoolean then x2:=x2+Dround(datam.Query1.FieldByName('KR').asFloat*form1.configRK.VAlue/100,2);
             x2:=DRound(x2,2);
              Fdatnacisl(datam.Query1.FieldByNAme('koddox').asString,datam.Query1.FieldByNAme('wm').asInteger,datam.Query1.FieldByNAme('wg').asInteger,
                  datam.Query1.FieldByNAme('datprov').asDateTime,tDat,tDat2,tDat3);

              xNdfl:=datam.Query1.Fieldbyname('snalog').asFloat;

             if RadioGroup1.ItemIndex=0 then
              begin
               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac='+floattostr(datam.Query1.FieldByName('kod').asFloat)+' and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;
              end;

             //
              bRTf2:=true;
              if not FOktmo(datam.Query1.Fieldbyname('wm').asinteger,datam.Query1.Fieldbyname('wg').asinteger,form_58.ComboBox4.Text) then
                  begin
                   x2:=0;
                   xNdfl:=0;
                   bRTf2:=false;
                  end;
              //

             if bRTf2 then
              begin
                if not tmprepdox.Locate('dat;dat2;dat3',VarArrayOf([tDAt,tDAt2,tDAt3]),[loCaseInsensitive]) then
                 begin
                  tmprepdox.Append;
                  tmprepdox.FieldByName('dat').asDateTime:=tDAt;
                  tmprepdox.FieldByName('dat2').asDateTime:=tDAt2;
                  tmprepdox.FieldByName('dat3').asDateTime:=tDAt3;
                  tmprepdox.post;
                 end;
                 tmprepdox.edit;
                 tmprepdox.FieldByName('summa').asFloat:=tmprepdox.FieldByName('summa').asFloat+x2;
                 tmprepdox.FieldByName('ndfl').asFloat:=tmprepdox.FieldByName('ndfl').asFloat+xndfl;
                 tmprepdox.post;
              end;

             datam.Query1.Next;
            end;





       st:=1;
       if form1.kartSTATUS.Value='2' then st:=2  ;

           datam.qtmp.Close;       //Начисл ГПХ
           datam.qtmp.SQL.Clear;
           datam.qtmp.DatabaseName:=form1.DBDIR;
           datam.qtmp.SQL.Add('select o.*, n.koddox, n.rk from obrt1new o, nacisl n where');
           datam.qtmp.sql.add('o.wg='+floattostr(RGOD));
           datam.qtmp.sql.add('and o.wm<='+floattostr(RMES));
           datam.qtmp.SQl.add('and o.kod=n.kod and o.nls='+floattostr(form1.kartNls.Value));
           datam.qtmp.prepare;
           datam.qtmp.open;
           datam.qtmp.first;
           while not datam.qtmp.eof do
            begin
              if (trim(datam.qtmp.fieldbyname('koddox').asString)='2010') then
                begin
                 x:=datam.qtmp.fieldbyname('kr').asFLoat;
                 if datam.qtmp.fieldbyname('rk').asboolean then x:=x+DRound(x*form1.configRK.Value/100,2);
                 sgpx[st]:=sgpx[st]+x;
                end;
             datam.qtmp.next;
            end;

       // !!!!!!!  обнулить Ddoxod9, qqMT[j,i] и т.д. для периода вне ОКТМО
       for i:=1 to 12 do
        begin
         if not FOktmo(i,RGod,form_58.ComboBox4.Text) then
           begin
            DDoxod35[i]:=0; DPn35[i]:=0;
            DDoxod9[i]:=0;
            for j:=1 to 10 do qqMT[j,i]:=0;
            for j:=1 to 10 do qqDV[j,i]:=0;
            for j:=1 to 6 do DST20[j,i]:=0;
            DImVyc[i]:=0; DPn1[i]:=0; DPn9[i]:=0;
            for j:=1 to 3 do sgpx[j]:=0;
           end;
        end;

       for i:=xMes1 to xMes2 do sDox[3]:=sDox[3]+DDoxod35[i];
       for i:=xMes1 to xMes2 do sNdfl[3]:=sNdfl[3]+DPn35[i];



        x:=0;
        for j:=1 to 10 do for i:=xMes1 to xMes2 do x:=x+qqMT[j,i]+DDoxod9[i];
        if x<>0 then sKolvo[st]:=sKolvo[st]+1;

        x:=0;
        for j:=1 to 10 do for i:=xMes1 to xMes2 do x:=x+DDoxod35[i];
        if x<>0 then sKolvo[3]:=sKolvo[3]+1;


        //доход
         x:=0;
        for j:=1 to 10 do for i:=xMes1 to xMes2 do x:=x+qqMT[j,i];
        sDox[st]:=sDox[st]+x;
        if (datam.kart2STATUS2.value=7) or (datam.kart2STATUS2.value=3) then
          begin
           s115[st]:=s115[st]+x;  //ВКС доход, кол-во
           if x<>0 then s121[st]:=s121[st]+1;
          end;

        //доход дивиденты
        x:=0;
        for i:=xMes1 to xMes2 do x:=x+DDoxod9[i];
        sDox9[st]:=sDox9[st]+x;
        //вычеты
        x:=0;
        for j:=1 to 10 do for i:=xMes1 to xMes2 do x:=x+qqDV[j,i];
        sVycet[st]:=sVycet[st]+x;
        x:=0;
        for j:=1 to 6 do for i:=xMes1 to xMes2 do x:=x+DST20[j,i];
        sVycet[st]:=sVycet[st]+x;
        x:=0;
        for i:=xMes1 to xMes2 do x:=x+DImVyc[i];
        sVycet[st]:=sVycet[st]+x;
        //ндфл


        x:=0;
        for i:=xMes1 to xMes2 do x:=x+DPn1[i];
        if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
         begin
          if ndflr5pr1.Value=1 then x:=ndflr5SISC.Value;  //из допсведений вместо расчетного налога
         end;
         sNdfl[st]:=sNdfl[st]+x;

        if (datam.kart2STATUS2.value=7) or (datam.kart2STATUS2.value=3) then  s142[st]:=s142[st]+x;  //ВКС исчислено

        //ндфл дивиж
        x:=0;
        for i:=xMes1 to xMes2 do x:=x+DPn9[i];
        sNdfl9[st]:=sNdfl9[st]+x;


       if FOktmo(1,RGod,form_58.ComboBox4.Text) then   //конечно неправильно январь смотрим в дальнейшем в ndfl5 сделать oktmo Поле
        begin
        if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
          begin
            sneuderj[st]:=sneuderj[st]+DRound(ndflr5.FieldByName('snuderj').asFloat,2);
            sfiks[st]:=sfiks[st]+DRound(ndflr5.FieldByName('sfix').asFloat,2);
            slisnee[st]:=slisnee[st]+DRound(ndflr5.FieldByName('suderj').asFloat,2);
          end;
        end;


        x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'');

      END;

      form1.kart.next;
     end;

     form102.close;
     PDEC2022:=0;
     if (RGod=2022) and (xMes2=12) then //з/п выплачена в январе 2023 за декабрь 2022 отнимаем из 2022
       begin
        if MessageDlg('Исключаем из расчета выплату з/п за вторую половину декабря 2022, которая была проведена в январе 2023 (если такая есть) ?',mtInformation,[mbYes,mbNo],0) = mrYes then
         begin
           PDEC2022:=1;
           form1.kart.first;
           form102.RxLabel1.Caption:='Обработка выплаты з/п зп декабрь 2022 в январе 2023'   ;
           form102.ProgressBar1.Position:=0;
           form102.Show;
           form102.Refresh;
           Nz0:=form1.kart.recordcount;
           Nz:=0;
           while not form1.kart.eof do
            begin
               nZ:=nZ+1;
               form102.ProgressBar1.Position:=Trunc(100*Nz/Nz0);
               form102.ProgressBar1.Refresh;
               datam.kart2.locate('nls',form1.kartnls.value,[locaseinsensitive]);
               ZapolnDOK(form1.kartnls.value);
               if FOktmo(12,RGod,form_58.ComboBox4.Text) then
                 begin
                  st:=1;
                  if form1.kartSTATUS.Value='2' then st:=2  ;
                  datam.qtmp.close;
                  datam.qtmp.sql.clear;
                  datam.qtmp.databasename:=form1.dbdir;
                  datam.qtmp.sql.add('select * from sdoxod where nls='+floattostr(form1.kartnls.value));
                  datam.qtmp.sql.add('and tavans<>1 and mes=12 and god=2022 and dat>='+#39+'01.01.2023'+#39);
                  datam.qtmp.prepare;
                  datam.qtmp.Open;
                  datam.qtmp.first;
                  while not datam.qtmp.eof do
                   begin
                     rtfgpx:=false;
                     if datam.qtmp.fieldbyname('kodnac').asfloat<>0 then
                      begin
                       if form1.nacisl.locate('kod',datam.qtmp.fieldbyname('kodnac').asfloat,[locaseinsensitive]) then
                         begin
                          if trim(form1.nacislkoddox.value)='2010' then rtfgpx:=true;
                         end;
                      end;
                     sDox[st]:=sDox[st]-datam.qtmp.fieldbyname('sdoxod').asfloat ;
                     if rtfgpx then sGpx[st]:=sGpx[st]-datam.qtmp.fieldbyname('sdoxod').asfloat;
                     sVycet[st]:=sVycet[st]-datam.qtmp.fieldbyname('rvicet').asfloat;
                     sNdfl[st]:=sNdfl[st]-datam.qtmp.fieldbyname('nalog').asfloat;
                     if (datam.kart2STATUS2.value=7) or (datam.kart2STATUS2.value=3) then  s142[st]:=s142[st]-datam.qtmp.fieldbyname('nalog').asfloat;  //ВКС исчислено

                    datam.qtmp.Next;
                   end;
                end; //октмо

           form1.kart.next;
          end;
         form102.close;
        end;
       end;

    END;


     form624:=Tform624.Create(nil);

    IF NOT TLOAD THEN
     BEGIN

      form624.PDEC2022_v:=PDEC2022;

      form624.RxCalcEdit1.Value:=DRound(sdox[1]+sdox9[1],2);
      form624.RxCalcEdit2.Value:=DRound(sdox9[1],2);
      form624.RxCalcEdit31.Value:=DRound(sdox[1]-sgpx[1]-sarenda[1],2);
      form624.RxCalcEdit34.Value:=DRound(sgpx[1],2);

      form624.RxCalcEdit3.Value:=DRound(svycet[1],2);
      form624.RxCalcEdit4.Value:=DRound(sndfl[1]+sndfl9[1],2);
      form624.RxCalcEdit5.Value:=DRound(sndfl9[1],2);
      form624.RxCalcEdit6.Value:=DRound(sfiks[1],2);

      form624.RxCalcEdit7.Value:=DRound(sdox[2]+sdox9[2],2);
      form624.RxCalcEdit8.Value:=DRound(sdox9[2],2);
      form624.RxCalcEdit32.Value:=DRound(sdox[2]-sgpx[2]-sarenda[2],2);
      form624.RxCalcEdit35.Value:=DRound(sgpx[2],2);

      form624.RxCalcEdit9.Value:=DRound(svycet[2],2);
      form624.RxCalcEdit10.Value:=DRound(sndfl[2]+sndfl9[2],2);
      form624.RxCalcEdit11.Value:=DRound(sndfl9[2],2);
      form624.RxCalcEdit12.Value:=DRound(sfiks[2],2);

      form624.RxCalcEdit17.Value:=DRound(sdox[3],2);
      form624.RxCalcEdit18.Value:=DRound(0,2);
      form624.RxCalcEdit33.Value:=0;
      form624.RxCalcEdit36.Value:=0;

      form624.RxCalcEdit19.Value:=DRound(0,2);
      form624.RxCalcEdit20.Value:=DRound(sndfl[3],2);
      form624.RxCalcEdit21.Value:=DRound(0,2);
      form624.RxCalcEdit22.Value:=DRound(0,2);


      form624.RxCalcEdit13.Value:=DRound(skolvo[1],2);
      form624.RxCalcEdit23.Value:=DRound(skolvo[2],2);
      form624.RxCalcEdit24.Value:=DRound(skolvo[3],2);


   //   form624.RxCalcEdit14.Value:=DRound(sndfl[1]+sndfl[2]+sndfl9[1]+sndfl9[2]+sndfl[3],2);
      form624.RxCalcEdit14.Value:=DRound(tNdflUderj[1],2);
      form624.RxCalcEdit25.Value:=DRound(tNdflUderj[2],2);
      form624.RxCalcEdit26.Value:=DRound(tNdflUderj[3],2);

      form624.RxCalcEdit15.Value:=DRound(sneuderj[1],2);
      form624.RxCalcEdit27.Value:=DRound(sneuderj[2],2);
      form624.RxCalcEdit28.Value:=DRound(sneuderj[3],2);

      form624.RxCalcEdit16.Value:=DRound(svozvr[1],2);
      form624.RxCalcEdit29.Value:=DRound(svozvr[2],2);
      form624.RxCalcEdit30.Value:=DRound(svozvr[3],2);

      form624.RxCalcEdit37.Value:=DRound(slisnee[1],2);
      form624.RxCalcEdit38.Value:=DRound(slisnee[2],2);
      form624.RxCalcEdit39.Value:=DRound(slisnee[3],2);

      form624.RxCalcEdit48.Value:=DRound(s115[1],2);
      form624.RxCalcEdit49.Value:=DRound(s115[2],2);
      form624.RxCalcEdit50.Value:=DRound(s115[3],2);

      form624.RxCalcEdit51.Value:=DRound(s121[1],2);
      form624.RxCalcEdit52.Value:=DRound(s121[2],2);
      form624.RxCalcEdit53.Value:=DRound(s121[3],2);

      form624.RxCalcEdit45.Value:=DRound(s142[1],2);
      form624.RxCalcEdit46.Value:=DRound(s142[2],2);
      form624.RxCalcEdit47.Value:=DRound(s142[3],2);

      form624.RxCalcEdit42.Value:=DRound(s155[1],2);
      form624.RxCalcEdit43.Value:=DRound(s155[2],2);
      form624.RxCalcEdit44.Value:=DRound(s155[3],2);

    
     END;

    IF  TLOAD THEN
     BEGIN
      form624.PDEC2022_v:=ndfl6DEC22.Value;
      form624.TID:=ndfl6ID.Value;
      form624.RxCalcEdit1.Value:=ndfl6p020.Value;
      form624.RxCalcEdit2.Value:=ndfl6p025.Value;
      form624.RxCalcEdit34.Value:=ndfl6p113.Value;
      form624.RxCalcEdit31.Value:=ndfl6p020.Value-ndfl6p113.Value-ndfl6p025.Value;


      form624.RxCalcEdit3.Value:=ndfl6p030.Value;
      form624.RxCalcEdit4.Value:=ndfl6p040.Value;
      form624.RxCalcEdit5.Value:=ndfl6p045.Value;
      form624.RxCalcEdit6.Value:=ndfl6p050.Value;

      form624.RxCalcEdit7.Value:=ndfl6p0202.Value;
      form624.RxCalcEdit8.Value:=ndfl6p0252.Value;
      form624.RxCalcEdit32.Value:=ndfl6p1132.Value;
      form624.RxCalcEdit35.Value:=ndfl6p0202.Value-ndfl6p1132.Value-ndfl6p0252.Value;


      form624.RxCalcEdit9.Value:=ndfl6p0302.Value;
      form624.RxCalcEdit10.Value:=ndfl6p0402.Value;
      form624.RxCalcEdit11.Value:=ndfl6p0452.Value;
      form624.RxCalcEdit12.Value:=ndfl6p0502.Value;
      form624.RxCalcEdit33.Value:=ndfl6p1133.Value;
      form624.RxCalcEdit36.Value:=ndfl6p0203.Value-ndfl6p1133.Value-ndfl6p0253.Value;


      form624.RxCalcEdit17.Value:=ndfl6p0203.Value;
      form624.RxCalcEdit18.Value:=ndfl6p0253.Value;
      form624.RxCalcEdit19.Value:=ndfl6p0303.Value;
      form624.RxCalcEdit20.Value:=ndfl6p0403.Value;
      form624.RxCalcEdit21.Value:=ndfl6p0453.Value;
      form624.RxCalcEdit22.Value:=ndfl6p0503.Value;


      form624.RxCalcEdit13.Value:=ndfl6p060.Value;
        form624.RxCalcEdit23.Value:=ndfl6p0602.Value;
        form624.RxCalcEdit24.Value:=ndfl6p0603.Value;

      form624.RxCalcEdit14.Value:=ndfl6p070.Value;
      form624.RxCalcEdit25.Value:=ndfl6p0702.Value;
      form624.RxCalcEdit26.Value:=ndfl6p0703.Value;

      form624.RxCalcEdit15.Value:=ndfl6p080.Value;
      form624.RxCalcEdit27.Value:=ndfl6p0802.Value;
      form624.RxCalcEdit28.Value:=ndfl6p0803.Value;

      form624.RxCalcEdit16.Value:=ndfl6p090.Value;
      form624.RxCalcEdit29.Value:=ndfl6p0902.Value;
      form624.RxCalcEdit30.Value:=ndfl6p0903.Value;

      form624.RxCalcEdit37.Value:=ndfl6p180.Value;
      form624.RxCalcEdit38.Value:=ndfl6p1802.Value;
      form624.RxCalcEdit39.Value:=ndfl6p1803.Value;

      form624.RxCalcEdit34.Value:=ndfl6p113.Value;
      form624.RxCalcEdit35.Value:=ndfl6p1132.Value;
      form624.RxCalcEdit36.Value:=ndfl6p1133.Value;

      form624.RxCalcEdit48.Value:=ndfl6p1151.Value;
      form624.RxCalcEdit49.Value:=ndfl6p1152.Value;
      form624.RxCalcEdit50.Value:=ndfl6p1153.Value;

      form624.RxCalcEdit51.Value:=ndfl6p1211.Value;
      form624.RxCalcEdit52.Value:=ndfl6p1212.Value;
      form624.RxCalcEdit53.Value:=ndfl6p1213.Value;

      form624.RxCalcEdit45.Value:=ndfl6p1421.Value;
      form624.RxCalcEdit46.Value:=ndfl6p1422.Value;
      form624.RxCalcEdit47.Value:=ndfl6p1423.Value;

      form624.RxCalcEdit42.Value:=ndfl6p1551.Value;
      form624.RxCalcEdit43.Value:=ndfl6p1552.Value;
      form624.RxCalcEdit44.Value:=ndfl6p1553.Value;


      ndfl6ob.first;
      while not ndfl6ob.eof do
       begin
        if ndfl6obID.Value=ndfl6ID.Value then
         begin
          tmprepdox.append;
          tmprepdox.fieldbyname('summa').asFloat:=ndfl6obSumma.Value;
          tmprepdox.fieldbyname('id2').asFloat:=ndfl6obId2.Value;
          tmprepdox.fieldbyname('ndfl').asFloat:=ndfl6obNdfl.Value;
          tmprepdox.fieldbyname('dat').asdateTime:=ndfl6obdat1.Value;
          tmprepdox.fieldbyname('dat2').asdateTime:=ndfl6obdat2.Value;
          tmprepdox.fieldbyname('dat3').asdateTime:=ndfl6obdat3.Value;
          tmprepdox.post;
         end;
        ndfl6ob.next;
       end;

     END;


     if ((not CheckBox1.Checked) and (xMes2=12)) or (CheckBox1.Checked) then form624.CheckBox3.Checked:=true else form624.CheckBox3.Checked:=false;

     form624.ShowModal;
     form624.Free;


end;

procedure TForm_58.JvXPButton10Click(Sender: TObject);
begin
 Form1.N68Click(nil);
end;


procedure TForm_58.FNdfl6raspred(tV:Boolean);     //tV=True - всегда пересчитывать, иначе если нет изменений по начислениям не пересчит
var
 i,k:integer;
    xpn,xpn9,xpn35:Real;
    x,y:Real;
    st:Real;
    xtmp,sumisx,sumprov,sumnalog:real;
    oldfiltered,oldfiltered2:Boolean;
    oldfilter,oldfilter2:String;
    oldDatasource,oldDatasource2:TDatasource;
    xSumma:Real;
    FStorno:array[1..12] of boolean;
    unalog,xid:real;
begin

//   showmessage(form1.kartfam.value);


  // if (form1.kartUVOLN.Value=1) and (not tV) then EXIT;


   datam.qtmp.databasename:=form1.dbdir;
   datam.qtmp.close;
   datam.qtmp.sql.clear;
   datam.qtmp.sql.add('SELECT sum(kr) from obrt1new where nls='+floattostr(form1.kartnls.value));
   datam.qtmp.sql.add('and wg='+floattostr(RGod));
   datam.qtmp.prepare;
   datam.qtmp.open;
   xSumma:=datam.qtmp.Fields[0].asFloat;
   datam.qtmp.close;

//   showmessage(floattostr(xSumma));


   datam.qtmp.databasename:=form1.dbdir;
   datam.qtmp.close;
   datam.qtmp.sql.clear;
   datam.qtmp.sql.add('SELECT sum(dayrab*dayotr*oklad) from glnew where nls='+floattostr(form1.kartnls.value));
   datam.qtmp.sql.add('and wg='+floattostr(RGod));
   datam.qtmp.prepare;
   datam.qtmp.open;
   xSumma:=xSumma+datam.qtmp.Fields[0].asFloat;
   xSumma:=DRound(xSumma,2);
   datam.qtmp.close;

//   showmessage(floattostr(xSumma));

   if not form1.foi.locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
    begin
     form1.foi.append;
     form1.foi.fieldbyname('nls').asFloat:=form1.kartnls.value;
     form1.foi.fieldbyname('god').asFloat:=RGod;
     form1.foi.fieldbyname('summa').asFloat:=xSumma;
     form1.foi.post;
    end
     else
    begin
     if (DRound(xSumma-form1.foiSumma.Value,2)=0) and (not tV) then EXIT; //выходим, не считаем т.к. нет  изменений
     form1.foi.edit;
     form1.foi.fieldbyname('summa').asFloat:=xSumma;
     form1.foi.post;
    end;

 //  showmessage('расчет регистров '+form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value);

   form850.rxLabel1.Caption:='Расчет регистров: '+form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value;
   form850.Show;
   form850.RxLabel1.Refresh;
   form850.RxLabel1.Repaint;

   datam.qtmp.databasename:=form1.dbdir;
   datam.qtmp.close;
   datam.qtmp.sql.clear;
   datam.qtmp.sql.add('UPDATE obrt1new SET datprov='+#39+'31.12.1899'+#39+' where nls='+floattostr(form1.kartnls.value));
   datam.qtmp.sql.add('and datprov='+#39+'30.12.1899'+#39);
   datam.qtmp.prepare;
   datam.qtmp.ExecSQL;
   datam.qtmp.close;

    oldfiltered:=form1.obrt1.filtered;
    oldfilter:=form1.obrt1.filter;
    oldfiltered2:=form1.obrt2.filtered;
    oldfilter2:=form1.obrt2.filter;
    olddatasource:=form11.DBGrid1.DataSource;
    olddatasource2:=form11.DBGrid2.DataSource;
    form11.DBGrid1.DataSource:=nil;
    form11.DBGrid2.DataSource:=nil;

    ZapolnMas;
    datam.qTmp.close;        //обнуляем поле svicet для мат помощи 4000
    datam.qTmp.Databasename:=form1.DBDIR;
    datam.qTmp.SQl.Clear;
    datam.qTmp.SQl.add('update obrt1new set svicet=0 where wg='+floattostr(Rgod)+' and nls='+floattostr(form1.kartnls.value));
    datam.qTmp.Prepare;
    datam.qTmp.ExecSQL;

    for i:=1 to 10 do for k:=1 to 12 do
     begin
      if qq2WDS[i,k]<>0 then
       begin
        {ShowMessage('код вычета '+q2WDK[i]+#13+
        'месяц='+floattostr(k)+#13+
         'суммы='+floattostr(q2WDS[i,k]));
          }
        if form1.obrt1.Locate('nls;wm;kod',VarArrayof([form1.kartnls.Value,k,qq2WD[i]]),[loCaseInsensitive]) then
         begin
          form1.obrt1.edit;
          form1.obrt1.fieldbyname('svicet').asFloat:=qq2WDS[i,k];
          form1.obrt1.post;
         // ShowMessage('заполнено');
         end;
       end;
     end;

    form_133.ZapolnStatus;
    for i:=1 to 12 do
     begin
      FStorno[i]:=false;
      if DStatus[i]=2 then st:=30 else st:=13;
      xpn:=DPn1[i];
      xpn9:=DPn9[i];
      xpn35:=DPn35[i];


           datam.qTmp.Close;
           datam.qTmp.SQL.Clear;
           datam.qTmp.SQL.Add('select n.koddox,n.rk,n.pvycet, n.stpnalog ,o.*,n.name from nacisl n, obrt1new o where o.kod=n.kod and o.wg='+floattostr(RGod));
           datam.qTmp.SQl.add('and o.wm='+floattostr(i));
           datam.qTmp.SQl.add('and o.nls='+floattostr(form1.kartNls.Value));
           datam.qTmp.SQl.add('and n.pn<>1');
           datam.qTmp.SQl.add('and (not o.proi='+#39+'*'+#39+')');   //остаток на прошлый год
           datam.qTmp.SQl.add('order by n.pvycet,o.datprov DESC');      //сортировка по убыванию      pvycet чтобы вычеты в конце были с кодом 2000 доход
           datam.qTmp.Prepare;
           datam.qTmp.Open;
           datam.qTmp.First;
         //  showmessage(floattostr(datam.qtmp.recordcount));
           FStorno[i]:=false;
           while not datam.qTmp.Eof do
            begin
           //  ShowMessage(datam.qTmp.fieldbyname('koddox').asString);

             if datam.qtmp.fieldbyname('kr').asFloat<0 then
               begin
                FStorno[i]:=True;

               end;


             if (datam.qTmp.fieldbyname('koddox').asString<>'1010') and ((datam.qTmp.fieldbyname('STPNALOG').asFloat=35)) then    //35%
              begin
                 y:=datam.qTmp.Fieldbyname('KR').asFloat;
                 if datam.qTmp.Fieldbyname('rk').asBoolean then y:=y+DRound(y*form1.configRK.Value/100,2);
                 if form1.obrt1.Locate('id',datam.qTmp.Fieldbyname('id').asFloat,[loCaseInsensitive]) then
                  begin
                    xtmp:=DRound(y*datam.qTmp.fieldbyname('STPNALOG').asFloat/100,0);
                    if xtmp<=xpn35 then
                     begin
                      xpn35:=xpn35-xtmp;
                      if xpn35=1 then   //остаток погрешность 1 рубль округление в последнюю запись
                        begin
                         xtmp:=xtmp+1;
                         xpn:=0;
                        end;
                      form1.obrt1.edit;
                     { if DRound(form1.obrt1.Fieldbyname('snalog').asFloat-xtmp,2)<>0 then
                          MessageDlg(form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.'+#13+
                           datam.qTmp.fieldbyname('name').asstring+' '+#13+'Изменено распределение ндфл '+#13+
                          'Старое значение: '+floattostrf(form1.obrt1.Fieldbyname('snalog').asFloat,ffNumber,12,2)+#13+
                          'Новое значение: '+floattostrf(xtmp,ffnumber,12,2),mtInformation,[mbOk],0);
                      }
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xtmp;
                      form1.obrt1.Post;
                   end
                    else
                     begin
                      form1.obrt1.edit;
                    {  if DRound(form1.obrt1.Fieldbyname('snalog').asFloat-xpn35,2)<>0 then
                          MessageDlg(form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.'+#13+
                           datam.qTmp.fieldbyname('name').asstring+' '+#13+'Изменено распределение ндфл '+#13+
                          'Старое значение: '+floattostrf(form1.obrt1.Fieldbyname('snalog').asFloat,ffNumber,12,2)+#13+
                          'Новое значение: '+floattostrf(xpn35,ffnumber,12,2),mtInformation,[mbOk],0);
                     }
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xpn35;
                      form1.obrt1.Post;
                      xpn:=0;
                     end;
                   end;
              end;

             if (datam.qTmp.fieldbyname('koddox').asString<>'1010') and ((datam.qTmp.fieldbyname('STPNALOG').asFloat=13)) then    //13% не дивиденды
              begin
                 y:=datam.qTmp.Fieldbyname('KR').asFloat;
                 if datam.qTmp.Fieldbyname('rk').asBoolean then y:=y+DRound(y*form1.configRK.Value/100,2);
                 sumisx:=y;
                 y:=y-datam.qTmp.Fieldbyname('svicet').asFloat;   //вычет мат.помощь 4000 в поле svicet
                 if form1.obrt1.Locate('id',datam.qTmp.Fieldbyname('id').asFloat,[loCaseInsensitive]) then
                  begin
                    xtmp:=DRound(y*st/100,0);

                    //showmessage(floattostr(xtmp));

                    {запрос уже проведено налог не трогаем}
                   if RGod>=2017 then
                    begin
                     datam.qtmp2.close;
                     datam.qtmp2.SQL.clear;
                     datam.qtmp2.databasename:=form1.DBDIR;
                     datam.qtmp2.SQL.add('select sum(sdoxod),sum(nalog) from sdoxod where nls='+floattostr(form1.kartnls.value));
                     datam.qtmp2.SQL.add('and mes='+Floattostr(form1.obrt1WM.value));
                     datam.qtmp2.SQL.add('and god='+Floattostr(form1.obrt1WG.value));
                     datam.qtmp2.SQL.add('and kodnac='+Floattostr(form1.obrt1KOD.value));
                     datam.qtmp2.prepare;
                     datam.qtmp2.open;
                     sumprov:=datam.qtmp2.fields[0].asFloat;
                     sumnalog:=datam.qtmp2.fields[1].asFloat;
                     datam.qtmp2.close;
                     if DRound(sumprov-sumisx,2)=0 then
                      begin
                       xtmp:=sumnalog;
                      { ShowMessage('Проведено полностью'+#13+form1.kartfam.value+#13+floattostr(form1.obrt1kod.value)+#13+floattostr(sumprov)+#13+
                        'месяц='+floattostr(form1.obrt1wm.value)+#13+
                        'Налог в sdoxod='+floattostr(sumnalog)+#13+
                         'Налог в obrt1='+floattostr(form1.obrt1snalog.value)+#13+
                         'не распределено еще='+floattostr(xpn));
                         }
                      end;
                    end;

                    {}

                   if xtmp<=xpn then
                    begin
                      xpn:=xpn-xtmp;
                      if xpn=1 then   //остаток погрешность 1 рубль округление в последнюю запись
                        begin
                         xtmp:=xtmp+1;
                         xpn:=0;
                        end;
                      form1.obrt1.edit;

                     { if DRound(form1.obrt1.Fieldbyname('snalog').asFloat-xtmp,2)<>0 then
                          MessageDlg(form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.'+#13+
                           datam.qTmp.fieldbyname('name').asstring+' '+#13+'Изменено распределение ндфл '+#13+
                          'Старое значение: '+floattostrf(form1.obrt1.Fieldbyname('snalog').asFloat,ffNumber,12,2)+#13+
                          'Новое значение: '+floattostrf(xtmp,ffnumber,12,2),mtInformation,[mbOk],0);
                     }
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xtmp;
                      form1.obrt1.Post;
                    end
                     else
                    begin
                      form1.obrt1.edit;
                    {  if DRound(form1.obrt1.Fieldbyname('snalog').asFloat-xpn35,2)<>0 then
                          MessageDlg(form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.'+#13+
                           datam.qTmp.fieldbyname('name').asstring+' '+#13+'Изменено распределение ндфл '+#13+
                          'Старое значение: '+floattostrf(form1.obrt1.Fieldbyname('snalog').asFloat,ffNumber,12,2)+#13+
                          'Новое значение: '+floattostrf(xpn35,ffnumber,12,2),mtInformation,[mbOk],0);
                     }
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xpn;
                      form1.obrt1.Post;
                      xpn:=0;
                     end;
                   end;
              end;
              if datam.qTmp.fieldbyname('koddox').asString='1010'  then    //дивиденды
              begin
                 y:=datam.qTmp.Fieldbyname('KR').asFloat;
                 if datam.qTmp.Fieldbyname('rk').asBoolean then y:=y+DRound(y*form1.configRK.Value/100,2);
                 if form1.obrt1.Locate('id',datam.qTmp.Fieldbyname('id').asFloat,[loCaseInsensitive]) then
                  begin
                    xtmp:=DRound(y*st/100,0);
                    if xtmp<=xpn9 then
                     begin
                      form1.obrt1.edit;
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xtmp;
                      form1.obrt1.Post;
                      xpn9:=xpn9-xtmp;
                   end
                    else
                     begin
                      form1.obrt1.edit;
                      form1.obrt1.Fieldbyname('snalog').asFloat:=xpn9;
                      form1.obrt1.Post;
                      xpn9:=0;
                     end;
                   end;
              end;



             datam.qTmp.Next;
            end;


           form2.LocateGlDb(form1.kartnls.value,i,RGod);  //оклад последний - в итоге будет первым чтобы вычеты к нему применить
           x:=DOklad[i];
           x:=DRound(x+x*form1.configRK.Value/100,2);
           xtmp:=DRound(x*st/100,0);
           if xtmp<>0 then xtmp:=xpn;  //остаток

             form1.GlDb.edit;
             form1.GlDb.Fieldbyname('snalog').asFloat:=xtmp;
             form1.GlDb.Post;


         if FStorno[i] then
          begin
           //перерасчет снова ндфл если есть сторно - в сторно записать разницу ндфл
           unalog:=form1.gldbSnalog.value;
           form1.obrt1.first;
           xid:=0;
           while not form1.obrt1.eof do
             begin
               if (form1.obrt1WM.value=i) and (form1.obrt1WG.Value=RGod) then
                begin
                 unalog:=unalog+form1.obrt1snalog.value;
                 if form1.obrt1KR.Value<0 then xid:=form1.obrt1ID.Value;
                end;
              form1.obrt1.Next;
             end;
          // showmessage('unalog'+#13+floattostr(unalog)+#13+floattostr(DPn1[i]));
            if xid<>0 then
             begin
              if form1.obrt1.Locate('id',xid,[locaseinsensitive]) then
               begin
                form1.obrt1.edit;
                form1.obrt1.fieldbyname('snalog').asfloat:=form1.obrt1snalog.value+DPn1[i]-unalog;
                form1.obrt1.post;
               // MessageDlg('Исправлен НДФЛ по операции сторно = '+floattostr(trunc(form1.obrt1snalog.value))+#13+
                //   namemes[i]+#13+form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value,mtInformation,[mbOk],0);
               end;
             end;
          end;

       end;



    form1.obrt1.filtered:=oldfiltered;
    form1.obrt1.filter:=oldfilter;
    form1.obrt2.filtered:=oldfiltered2;
    form1.obrt2.filter:=oldfilter2;
    form11.DBGrid1.DataSource:=olddatasource;
    form11.DBGrid2.DataSource:=olddatasource2;




    form2.LocateGlDb(form1.kartnls.value,RMes,RGod);
    form850.close;



end;



procedure TForm_58.Button1Click(Sender: TObject);
var nZ0,nZ:integer;
begin

  form1.NACISL.First;
  while not form1.NACISL.Eof do
   begin
     form1.nacisl.edit;
     if form1.nacislkoddox.value='2000' then form1.nacisl.fieldbyname('pvycet').asFloat:=1 else form1.nacisl.fieldbyname('pvycet').asFloat:=0;
     form1.nacisl.post;
    form1.NACISL.Next;
   end;
   form1.nacisl.first;

   form102.RxLabel1.Caption:='Расчет регистров 6-НДФЛ'   ;
   form102.ProgressBar1.Position:=0;
   form102.Show;
   form102.Refresh;
   Nz0:=form1.kart.RecordCount;




 form1.kart.First;
 while not form1.kart.eof do
  begin
     nZ:=nZ+1;
     form102.ProgressBar1.Position:=Trunc(100*Nz/Nz0);
     form102.ProgressBar1.Refresh;
     form102.ProgressBar1.Repaint;


     form_58.FNdfl6raspred(False);

   form1.kart.next;
  end;

   form102.Close;


end;

procedure TForm_58.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 k6ndfl.close;
 oktmo.close;
 ndflr5.Close;
 reorg.close;
end;

procedure TForm_58.ComboBox1Change(Sender: TObject);
begin
 if not oktmo.Locate('oktmo',trim(ComboBox1.Text),[loCaseInsensitive]) then
  begin
   oktmo.append;
   oktmo.fieldbyname('oktmo').asString:=trim(ComboBox1.Text);
   oktmo.post;
  end;

 Edit4.Text:=FSetGni(ComboBox1.Text);
 Edit7.Text:=FSetKPP(ComboBox1.Text);
 if Edit4.Text<>'????' then
   begin
    // MessageDlg('Изменен код получателя ИФНС '+Edit4.Text,mtInformation,[mbOk],0)
   end
     else
       MessageDlg('Не определен код получателя ИФНС '+Edit4.Text+#13+'Введите значение в соответствующее поле и нажмите Сохранить',mtInformation,[mbOk],0)  ;
 Edit4.SetFocus;
end;

procedure TForm_58.ComboBox4Change(Sender: TObject);
begin
 if not oktmo.Locate('oktmo',trim(ComboBox4.Text),[loCaseInsensitive]) then
  begin
   oktmo.append;
   oktmo.fieldbyname('oktmo').asString:=trim(ComboBox4.Text);
   oktmo.post;
  end;

 Edit5.Text:=FSetGni(ComboBox4.Text);
 Edit8.Text:=FSetKPP(ComboBox4.Text);
 if Edit5.Text<>'????' then
    begin
            // MessageDlg('Изменен код получателя ИФНС '+Edit5.Text,mtInformation,[mbOk],0)
    end
       else
               MessageDlg('Не определен код получателя ИФНС '+Edit5.Text+#13+'Введите значение в соответствующее поле и нажмите Сохранить',mtInformation,[mbOk],0)  ;
 Edit5.SetFocus;

end;

procedure TForm_58.SpeedButton1Click(Sender: TObject);
begin
 if not oktmo.Locate('oktmo',trim(ComboBox1.Text),[loCaseInsensitive]) then
  begin
   oktmo.append;
   oktmo.fieldbyname('oktmo').asString:=trim(ComboBox1.Text);
   oktmo.post;
  end;

 
 s:=trim(edit4.Text);
 s:=mainlib.FAc(s);
 if length(s)<>4 then
   begin
    Showmessage('Код ИФНС состоит из 4 цифр');
    exit;
   end;

   if MessageDlg('Сохранить код ИФНС='+Edit4.Text+#13+'КПП='+trim(Edit7.Text)+#13+'для ОКТМО='+ComboBox1.Text,mtWarning,[mbYes,mbNo],0) = mrYes then
     begin
      oktmo.edit;
      oktmo.fieldbyname('ifns').asString:=Trim(Edit4.Text);
      oktmo.fieldbyname('kpp').asString:=Trim(Edit7.Text);
      oktmo.post;
     end;

end;

procedure TForm_58.SpeedButton2Click(Sender: TObject);
var s:String;
begin
 if not oktmo.Locate('oktmo',trim(ComboBox4.Text),[loCaseInsensitive]) then
  begin
   oktmo.append;
   oktmo.fieldbyname('oktmo').asString:=trim(ComboBox4.Text);
   oktmo.post;
  end;

 s:=trim(edit5.Text);
 s:=mainlib.FAc(s);
 if length(s)<>4 then
   begin
    Showmessage('Код ИФНС состоит из 4 цифр');
    exit;
   end;


   if MessageDlg('Сохранить код ИФНС='+Edit5.Text+#13+#13+'КПП='+Edit8.Text+#13+'для ОКТМО='+ComboBox4.Text,mtWarning,[mbYes,mbNo],0) = mrYes then
     begin
      oktmo.edit;
      oktmo.fieldbyname('ifns').asString:=Trim(Edit5.Text);
      oktmo.fieldbyname('kpp').asString:=Trim(Edit8.Text);
      oktmo.post;
     end;

end;


procedure TForm_58.PSpr2016(wnls:Real);
var fNameXLS:String;
    E:OleVAriant;
    xRegion:Real;
    dd,mm,yy:Word;
    s1:String;
    nd,nc,i,j,k:Integer;
    x:Real;
    soktmo:String;
    sKpp:String;
    nSpr:Real;
    nczap:integer;
    nDatSpr:TDate;
begin

    fNameXls:=GetNameXlsn('2НДФЛ','nd')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\2NDFL2017.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      exit;
     end;
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;

     E.ActiveWorkBook.Sheets.Item[1].Range['AH5']:=RGod;
     E.ActiveWorkBook.Sheets.Item[1].Range['AH7']:='1';

     nSpr:=INT(rxCalcEdit1.Value);
     nDatSpr:=DateEdit1.DAte;

     //****
     if PRNUMDAT THEN
      BEGIN
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select num,dat from spr2006 where nls='+FloatToStr(form1.kartNls.Value));
       datam.Query1.SQL.Add('and GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       if datam.Query1.RecordCount>0 then
         begin
          nSpr:=datam.Query1.Fields[0].asFloat;
          nDatSpr:=datam.Query1.Fields[1].asDateTime;
          rxCalcEdit1.Value:=rxCAlcEdit1.Value-1;
         end;
       datam.Query1.Close;
      END;
     //**

     E.ActiveWorkBook.Sheets.Item[1].Range['AX5']:=nSpr;
     E.ActiveWorkBook.Sheets.Item[1].Range['BK7']:=trim(edit3.text);
     E.ActiveWorkBook.Sheets.Item[1].Range['BO10']:=form1.config2INN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['P11']:=form1.config2NAME.Value;
     datam.kart2.Locate('nls',form1.kartnls.value,[locaseinsensitive]);


     if Trim(datam.kart2OKTMO.Value)='' then soktmo:=form1.config2OKTMO.Value else
                                                                soktmo:=datam.kart2OKTMO.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['N10']:=soktmo;

     if oktmo.Locate('oktmo',soktmo,[locaseinsensitive]) then
      begin
       sKpp:=oktmoKpp.Value;
      end
       else
        skpp:=form1.config2KPP.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['CN10']:=sKPP;


     E.ActiveWorkBook.Sheets.Item[1].Range['CE7']:=form_58.FSetGni(soktmo);


     E.ActiveWorkBook.Sheets.Item[1].Range['AP10']:=form1.config2TEL.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['AB13']:=form1.kartINN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['J14']:=form1.kartFAM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['AV14']:=form1.kartIM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CP14']:=form1.kartOT.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CJ15']:=form1.kartSTRANA.Value;

     datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
     if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=6) then
                   E.ActiveWorkBook.Sheets.Item[1].Range['Y15']:=datam.kart2STATUS2.AsString
                         else E.ActiveWorkBook.Sheets.Item[1].Range['Y15']:='1';



        DecodeDAte(nDatSpr,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BS5']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BI5']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['BN5']:=s1;



     if form1.kartBIRTHDAY.Value>=Encodedate(1920,1,1) then
       begin
        DecodeDAte(form1.kartBirthday.Value,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BE15']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['AU15']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ15']:=s1;
       end;

       E.ActiveWorkBook.Sheets.Item[1].Range['AO16']:=form1.kartKODDOC.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BU16']:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['BI17']:=form1.kartINDEX.Value;
       if xRegion>=10 then E.ActiveWorkBook.Sheets.Item[1].Range['CJ17']:=xRegion else E.ActiveWorkBook.Sheets.Item[1].Range['CJ17']:='0'+FloatToStr(xRegion);

       E.ActiveWorkBook.Sheets.Item[1].Range['G18']:=form1.kartRAYON.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AP18']:=form1.kartGOROD.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['CM18']:=form1.kartNASPUNKT.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['G19']:=form1.kartULICA.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BW19']:=form1.kartDOM.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['CM19']:=form1.kartKORPUS.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['DC19']:=form1.kartKVART.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['A63']:=podpndflFIO.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BP61']:=podpndflAgent.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BF67']:=podpndflDOKUM.VAlue;


       E.ActiveWorkBook.Sheets.Item[1].Range['W20']:=form1.kartSTRANA2.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AI20']:=form1.kartADRES2.Value;
       if form1.kartSTATUS.Value='2' then E.ActiveWorkBook.Sheets.Item[1].Range['AI22']:=30 else E.ActiveWorkBook.Sheets.Item[1].Range['AI22']:=13;

       ObrabObrtNalKart;

     {   for i:=1 to 12 do begin for j:=1 to 6 do
          begin
           Showmessage('j='+floattostr(j)+', qV='+qV[j]+#13+
           'i='+floattostr(i)+', qDV='+floattostr(qDV[j,i]));
          end;
         end;
    }

  k:=0;
  for j:=1 to 10 do
    begin
     if qqD[j]<>'' then k:=j;  //кол-во кодов для дивидендов 13% добавить чтобы
    end;
  if RGod>=2015 then
   begin
    if DKOD9[1]<>'' then
     begin
      qqD[k+1]:=DKOD9[1];
      for i:=1 to 12 do qqMT[k+1,i]:=DDOX9[1,i];
      for i:=1 to 12 do qqMT[k+1,13]:=qqMT[k+1,13]+DDOX9[1,i];
     end;
   end;



     nc:=1;nd:=0;


     nczap:=0;
     for i:=1 to 12 do for j:=1 to 10 do if (qqD[j]<>'') and (qqMT[j,i]<>0) then  nczap:=nczap+1;
     nczap:=TRUNC(nczap/2)+2;

     for i:=1 to 12 do
      begin
       for j:=1 to 10 do
        begin
          if (qqD[j]<>'') and (qqMT[j,i]<>0) then
           begin
            nd:=nd+1;
            if nd=nczap then
             begin
              nd:=1; nc:=nc+1;
             end;
            if nc=1 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(nd+24)]:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['Q'+IntToStr(nd+24)]:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AF'+IntToStr(nd+24)]:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(nd+24)]:=qqDV[j,i];
             end;
            if nc=2 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['BG'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['BO'+IntToStr(nd+24)]:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['BW'+IntToStr(nd+24)]:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CL'+IntToStr(nd+24)]:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CT'+IntToStr(nd+24)]:=qqDV[j,i];
             end;
           end;
        end;
       end;

      for j:=1 to 6 do
      begin
        if (DS20[j]<>'') and (DST20[j,13]<>0) then
         begin
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['A47']:=DS20[j];
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['J47']:=DST20[j,13];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AB47']:=DS20[j];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AK47']:=DST20[j,13];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BD47']:=DS20[j];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BM47']:=DST20[j,13];
         end;
      end;

    if form1.kartIMVYC_SUMM.Value<>0 then
     begin
      E.ActiveWorkBook.Sheets.Item[1].Range['BR51']:=form1.kartImVyc_Num.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['CG47']:=form1.kartIMVYC_KOD.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['DE51']:=form1.kartIMVYC_gni.Value;

      DecodeDAte(form1.kartImVyc_Dat.Value,yy,mm,dd);
      if dd<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CD51']:='0'+floattostr(dd) else E.ActiveWorkBook.Sheets.Item[1].Range['CD51']:=floattostr(dd);
      if mm<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CI51']:='0'+floattostr(mm) else E.ActiveWorkBook.Sheets.Item[1].Range['CI51']:=floattostr(mm);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN51']:=yy;
      x:=0;
      for j:=1 to 12 do x:=x+DImVyc[j];
      E.ActiveWorkBook.Sheets.Item[1].Range['CP47']:=x;


     end;

    x:=0;
    for j:=1 to 10 do x:=x+qqMT[j,13];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF53']:=DRound(x,2);
    x:=DNal[13];
    for j:=1 to 12 do x:=x+DDoxod9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF54']:=DRound(x,2);

    x:=DIsc[13];
    for j:=1 to 13 do x:=x+DPn9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF55']:=DRound(x,2);
    E.ActiveWorkBook.Sheets.Item[1].Range['CN53']:=DRound(x,2);

    x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'')+FUplataNdfl(form1.kartNLS.Value,RGod,9,'');
    E.ActiveWorkBook.Sheets.Item[1].Range['CN54']:=DRound(x,2);

    if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
     begin

      E.ActiveWorkBook.Sheets.Item[1].Range['CN53']:=DRound(ndflr5.FieldByName('sud').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['AF55']:=DRound(ndflr5.FieldByName('sisc').asFloat,2);

      E.ActiveWorkBook.Sheets.Item[1].Range['CN55']:=DRound(ndflr5.FieldByName('suderj').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN56']:=DRound(ndflr5.FieldByName('snuderj').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['AF56']:=DRound(ndflr5.FieldByName('sfix').asFloat,2);
      if ndflr5.FieldByName('sfix').asFloat<>0 then
       begin
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59']:=ndflr5.FieldByName('num').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59']:=ndflr5.FieldByName('ifns').asString;
        DecodeDate(ndflr5.FieldByName('dat').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59']:=s1;


       end;
     end;

   try
    nczap:=nczap+24;
    if nczap<42 then E.ActiveWorkBook.Sheets.Item[1].Range['A'+Inttostr(nczap),'A'+IntToStr(43)].EntireRow.Delete(EmptyParam);
   except
   end;


   try
    E.DisplayAlerts:=false;
    E.WorkBooks[1].Save;
   except
   end;

   x:=0;                         //35% отдельная справка
   for i:=1 to 12 do x:=x+DDoxod35[i];
   if x<>0 then
     begin
      rxCalcEdit1.Value:=rxCalcEdit1.Value+1;
      PSpr2016_st35(wnls);
     end;

end;



procedure TForm_58.PSpr2017(wnls:Real);
var fNameXLS:String;
    E:OleVAriant;
    xRegion:Real;
    dd,mm,yy:Word;
    s1:String;
    nd,nc,i,j,k:Integer;
    x:Real;
    soktmo:String;
    sKpp:String;
    nSpr:Real;
    nczap:integer;
    nDatSpr:TDate;
begin

    fNameXls:=GetNameXlsn('2НДФЛ','nd')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\2NDFL2018.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      exit;
     end;
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;

     E.ActiveWorkBook.Sheets.Item[1].Range['AH5']:=RGod;
     E.ActiveWorkBook.Sheets.Item[1].Range['AH7']:=RPriznak;

     nSpr:=INT(rxCalcEdit1.Value);
     nDatSpr:=DateEdit1.DAte;

     //****
     if PRNUMDAT THEN
      BEGIN
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select num,dat from spr2006 where nls='+FloatToStr(form1.kartNls.Value));
       datam.Query1.SQL.Add('and GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       if datam.Query1.RecordCount>0 then
         begin
          nSpr:=datam.Query1.Fields[0].asFloat;
          nDatSpr:=datam.Query1.Fields[1].asDateTime;
          rxCalcEdit1.Value:=rxCAlcEdit1.Value-1;
         end;
       datam.Query1.Close;
      END;
     //**

     E.ActiveWorkBook.Sheets.Item[1].Range['AX5']:=nSpr;
     E.ActiveWorkBook.Sheets.Item[1].Range['BK7']:=trim(edit3.text);
     E.ActiveWorkBook.Sheets.Item[1].Range['BO10']:=form1.config2INN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['P11']:=form1.config2NAME.Value;
     datam.kart2.Locate('nls',form1.kartnls.value,[locaseinsensitive]);

     E.ActiveWorkBook.Sheets.Item[1].Range['AK12']:=form_58.RKOD;
     E.ActiveWorkBook.Sheets.Item[1].Range['AK13']:=form_58.RINN;
     E.ActiveWorkBook.Sheets.Item[1].Range['BF13']:=form_58.RKPP;


     if Trim(datam.kart2OKTMO.Value)='' then soktmo:=form1.config2OKTMO.Value else
                                                                soktmo:=datam.kart2OKTMO.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['N10']:=soktmo;

     if oktmo.Locate('oktmo',soktmo,[locaseinsensitive]) then
      begin
       sKpp:=oktmoKpp.Value;
      end
       else
        skpp:=form1.config2KPP.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['CN10']:=sKPP;


     E.ActiveWorkBook.Sheets.Item[1].Range['CE7']:=form_58.FSetGni(soktmo);


     E.ActiveWorkBook.Sheets.Item[1].Range['AP10']:=form1.config2TEL.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['AB17']:=form1.kartINN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['J18']:=form1.kartFAM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['AV18']:=form1.kartIM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CP18']:=form1.kartOT.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CJ19']:=form1.kartSTRANA.Value;

     datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
     if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=6) then
                   E.ActiveWorkBook.Sheets.Item[1].Range['Y19']:=datam.kart2STATUS2.AsString
                         else E.ActiveWorkBook.Sheets.Item[1].Range['Y19']:='1';



        DecodeDAte(nDatSpr,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BS5']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BI5']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['BN5']:=s1;



     if form1.kartBIRTHDAY.Value>=Encodedate(1920,1,1) then
       begin
        DecodeDAte(form1.kartBirthday.Value,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BE19']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['AU19']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ19']:=s1;
       end;

       E.ActiveWorkBook.Sheets.Item[1].Range['AO20']:=form1.kartKODDOC.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BU20']:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;


       E.ActiveWorkBook.Sheets.Item[1].Range['A63']:=podpndflFIO.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['DB61']:=podpndflAgent.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BF67']:=podpndflDOKUM.VAlue;



       if form1.kartSTATUS.Value='2' then E.ActiveWorkBook.Sheets.Item[1].Range['AI22']:=30 else E.ActiveWorkBook.Sheets.Item[1].Range['AI22']:=13;

       ObrabObrtNalKart;



  k:=0;
  for j:=1 to 10 do
    begin
     if qqD[j]<>'' then k:=j;  //кол-во кодов для дивидендов 13% добавить чтобы
    end;
  if RGod>=2015 then
   begin
    if DKOD9[1]<>'' then
     begin
      qqD[k+1]:=DKOD9[1];
      for i:=1 to 12 do qqMT[k+1,i]:=DDOX9[1,i];
      for i:=1 to 12 do qqMT[k+1,13]:=qqMT[k+1,13]+DDOX9[1,i];
     end;
   end;



     nc:=1;nd:=0;


     nczap:=0;
     for i:=1 to 12 do for j:=1 to 10 do if (qqD[j]<>'') and (qqMT[j,i]<>0) then  nczap:=nczap+1;
     nczap:=TRUNC(nczap/2)+2;

     for i:=1 to 12 do
      begin
       for j:=1 to 10 do
        begin
          if (qqD[j]<>'') and (qqMT[j,i]<>0) then
           begin
            nd:=nd+1;
            if nd=nczap then
             begin
              nd:=1; nc:=nc+1;
             end;
            if nc=1 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(nd+24)]:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['Q'+IntToStr(nd+24)]:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AF'+IntToStr(nd+24)]:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(nd+24)]:=qqDV[j,i];
             end;
            if nc=2 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['BG'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['BO'+IntToStr(nd+24)]:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['BW'+IntToStr(nd+24)]:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CL'+IntToStr(nd+24)]:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CT'+IntToStr(nd+24)]:=qqDV[j,i];
             end;
           end;
        end;
       end;

      for j:=1 to 6 do
      begin
        if (DS20[j]<>'') and (DST20[j,13]<>0) then
         begin
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['A47']:=DS20[j];
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['J47']:=DST20[j,13];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AB47']:=DS20[j];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AK47']:=DST20[j,13];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BD47']:=DS20[j];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BM47']:=DST20[j,13];
         end;
      end;

    if form1.kartIMVYC_SUMM.Value<>0 then
     begin
      E.ActiveWorkBook.Sheets.Item[1].Range['BR51']:=form1.kartImVyc_Num.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['CG47']:=form1.kartIMVYC_KOD.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['DE51']:=form1.kartIMVYC_gni.Value;

      DecodeDAte(form1.kartImVyc_Dat.Value,yy,mm,dd);
      if dd<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CD51']:='0'+floattostr(dd) else E.ActiveWorkBook.Sheets.Item[1].Range['CD51']:=floattostr(dd);
      if mm<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CI51']:='0'+floattostr(mm) else E.ActiveWorkBook.Sheets.Item[1].Range['CI51']:=floattostr(mm);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN51']:=yy;
      x:=0;
      for j:=1 to 12 do x:=x+DImVyc[j];
      E.ActiveWorkBook.Sheets.Item[1].Range['CP47']:=x;


     end;

    x:=0;
    for j:=1 to 10 do x:=x+qqMT[j,13];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF53']:=DRound(x,2);
    x:=DNal[13];
    for j:=1 to 12 do x:=x+DDoxod9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF54']:=DRound(x,2);

    x:=DIsc[13];
    for j:=1 to 13 do x:=x+DPn9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF55']:=DRound(x,2);
    E.ActiveWorkBook.Sheets.Item[1].Range['CN53']:=DRound(x,2);

    x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'')+FUplataNdfl(form1.kartNLS.Value,RGod,9,'');
    E.ActiveWorkBook.Sheets.Item[1].Range['CN54']:=DRound(x,2);


    datam.query1.close;
    datam.query1.sql.clear;
    datam.query1.sql.add('select * from uplatandfl where nls='+floattostr(form1.kartnls.value));
    datam.query1.sql.add('and type=2 and summa>0 and god='+floattostr(RGod));
    datam.query1.Prepare;
    datam.Query1.open;
    datam.query1.first;
     if datam.query1.RecordCount=1 then //аванс
      begin
        E.ActiveWorkBook.Sheets.Item[1].Range['AF56']:=DRound(datam.query1.FieldByName('summa').asFloat,2);
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59']:=datam.query1.FieldByName('numuved').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59']:=datam.query1.FieldByName('ifns').asString;
        DecodeDate(datam.query1.FieldByName('datuved').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59']:=s1;
      end;
    datam.query1.close;

    if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
     begin
      E.ActiveWorkBook.Sheets.Item[1].Range['CN53']:=DRound(ndflr5.FieldByName('sud').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['AF55']:=DRound(ndflr5.FieldByName('sisc').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN55']:=DRound(ndflr5.FieldByName('suderj').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN56']:=DRound(ndflr5.FieldByName('snuderj').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['AF56']:=DRound(ndflr5.FieldByName('sfix').asFloat,2);
      if ndflr5.FieldByName('sfix').asFloat<>0 then
       begin
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59']:=ndflr5.FieldByName('num').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59']:=ndflr5.FieldByName('ifns').asString;
        DecodeDate(ndflr5.FieldByName('dat').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59']:=s1;


       end;
     end;

   try
    nczap:=nczap+24;
    if nczap<42 then E.ActiveWorkBook.Sheets.Item[1].Range['A'+Inttostr(nczap),'A'+IntToStr(43)].EntireRow.Delete(EmptyParam);
   except
   end;

    E.WindowState:=-4137 ;

   try
    E.DisplayAlerts:=false;
    E.WorkBooks[1].Save;
   except
   end;


   x:=0;                         //35% отдельная справка
   for i:=1 to 12 do x:=x+DDoxod35[i];
   if x<>0 then
     begin
      rxCalcEdit1.Value:=rxCalcEdit1.Value+1;
      PSpr2016_st35(wnls);
     end;

end;



procedure TForm_58.PSpr2016_st35(wnls:Real);
var fNameXLS:String;
    E:OleVAriant;
    xRegion:Real;
    dd,mm,yy:Word;
    s1:String;
    nd,nc,i,j,k:Integer;
    x:Real;
    soktmo:String;
begin

    fNameXls:=GetNameXlsn('2НДФЛ','nd')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\2NDFL2016.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      exit;
     end;
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;

     E.ActiveWorkBook.Sheets.Item[1].Range['AH5']:=RGod;
     E.ActiveWorkBook.Sheets.Item[1].Range['AH7']:='1';
     E.ActiveWorkBook.Sheets.Item[1].Range['AX5']:=INT(rxCalcEdit1.Value);
     E.ActiveWorkBook.Sheets.Item[1].Range['BK7']:=trim(edit3.text);
     E.ActiveWorkBook.Sheets.Item[1].Range['BO10']:=form1.config2INN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['CN10']:=form1.config2KPP.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['P11']:=form1.config2NAME.Value;
     datam.kart2.Locate('nls',form1.kartnls.value,[locaseinsensitive]);


     if Trim(datam.kart2OKTMO.Value)='' then soktmo:=form1.config2OKTMO.Value else
                                                                soktmo:=datam.kart2OKTMO.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['N10']:=soktmo;

     E.ActiveWorkBook.Sheets.Item[1].Range['CE7']:=form_58.FSetGni(soktmo);


     E.ActiveWorkBook.Sheets.Item[1].Range['AP10']:=form1.config2TEL.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['AB13']:=form1.kartINN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['J14']:=form1.kartFAM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['AV14']:=form1.kartIM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CP14']:=form1.kartOT.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CJ15']:=form1.kartSTRANA.Value;

     datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
     if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=6) then
                   E.ActiveWorkBook.Sheets.Item[1].Range['Y15']:=datam.kart2STATUS2.AsString
                         else E.ActiveWorkBook.Sheets.Item[1].Range['Y15']:='1';

        DecodeDAte(DateEdit1.DAte,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BS5']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BI5']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['BN5']:=s1;



     if form1.kartBIRTHDAY.Value>=Encodedate(1920,1,1) then
       begin
        DecodeDAte(form1.kartBirthday.Value,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BE15']:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['AU15']:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ15']:=s1;
       end;

       E.ActiveWorkBook.Sheets.Item[1].Range['AO16']:=form1.kartKODDOC.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BU16']:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['BI17']:=form1.kartINDEX.Value;
       if xRegion>=10 then E.ActiveWorkBook.Sheets.Item[1].Range['CJ17']:=xRegion else E.ActiveWorkBook.Sheets.Item[1].Range['CJ17']:='0'+FloatToStr(xRegion);

       E.ActiveWorkBook.Sheets.Item[1].Range['G18']:=form1.kartRAYON.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AP18']:=form1.kartGOROD.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['CM18']:=form1.kartNASPUNKT.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['G19']:=form1.kartULICA.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BW19']:=form1.kartDOM.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['CM19']:=form1.kartKORPUS.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['DC19']:=form1.kartKVART.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['A59']:=podpndflFIO.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BP57']:=podpndflAgent.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BF63']:=podpndflDOKUM.VAlue;


       E.ActiveWorkBook.Sheets.Item[1].Range['W20']:=form1.kartSTRANA2.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AI20']:=form1.kartADRES2.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AI22']:=35;

       for i:=1 to 2 do
         begin
          DKOD35[i]:=''; DKODVIC35[i]:='';
          for k:=1 to 13 do
           begin
            DDOX35[i,k]:=0;
            DVIC35[i,k]:=0;
           end;
         end;

       ObrabObrtNalKart;

      nc:=1;nd:=0;
    {
      for i:=1 to 12 do
             for j:=1 to 2 do
              Showmessage(floattostr(i)+#13+floattostr(j)+#13+DKOD35[j]+#13+
               Floattostr(WDS35[j,i])+#13+floattostr(DDox35[j,i]));
     }

     nc:=1;nd:=0;
     for i:=1 to 12 do
      begin
       for j:=1 to 2 do
        begin
          if (DKOD35[j]<>'') and (DDOX35[j,i]<>0) then
           begin
            nd:=nd+1;
            if nd=16 then
             begin
              nd:=1; nc:=nc+1;
             end;
            if nc=1 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(nd+24)]:=DKOD35[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['Q'+IntToStr(nd+24)]:=DDOX35[j,i];
               if (DKODVIC35[j]<>'') and (WDS35[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AF'+IntToStr(nd+24)]:=DKODVIC35[j];
               if (DKODVIC35[j]<>'') and (WDS35[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(nd+24)]:=WDS35[j,i];
             end;
            if nc=2 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['BG'+IntToStr(nd+24)]:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['BO'+IntToStr(nd+24)]:=DKOD35[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['BW'+IntToStr(nd+24)]:=WDS35[j,i];
               if (DKODVIC35[j]<>'') and (WDS35[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CL'+IntToStr(nd+24)]:=DKODVIC35[j];
               if (DKODVIC35[j]<>'') and (WDS35[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CT'+IntToStr(nd+24)]:=WDS35[j,i];
             end;
           end;
        end;
       end;


    x:=0;
    for j:=1 to 12 do x:=x+DDoxod35[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF49']:=DRound(x,2);
    E.ActiveWorkBook.Sheets.Item[1].Range['AF50']:=DRound(x,2);

    x:=0;
    for j:=1 to 12 do x:=x+DPn35[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF51']:=DRound(x,2);
    E.ActiveWorkBook.Sheets.Item[1].Range['CN49']:=DRound(x,2);

    x:=FUplataNdfl(form1.kartNLS.Value,RGod,35,'');
    E.ActiveWorkBook.Sheets.Item[1].Range['CN50']:=DRound(x,2);



   try
    E.DisplayAlerts:=false;
    E.WorkBooks[1].Save;
   except
   end;


end;



procedure TForm_58.JvXPButton9Click(Sender: TObject);

begin
 form642:=Tform642.Create(nil);
 form642.ShowModal;
 form642.Free;

end;

procedure TForm_58.JvXPButton11Click(Sender: TObject);
begin
 form643:=TForm643.Create(nil);
 form643.ShowModal;
 form643.Free;
end;


procedure TForm_58.FormRepSrok(tTypeDat:Integer);    // 0 - по сроку уплаты , 1 - по дате исчисления
var ds:String;
    gm1,gm2:Integer;
    gok:Boolean;
    km,kg:integer;
    tdat1,tdat2,tdat3:TDate;
    x2:Real;
    fNameXLS:String;
    Nstroka:Integer;
    E:OleVariant;
    sname:string;
    xs0,xn0,xndfl:Real;
    datstart,datend, srokdat1, srokdat2:TDate;
    bRTF:Boolean;
begin

{
 form58:=TForm58.Create(Self);
 form58.ShowModal;
 gm1:=form58.fm1;
 gm2:=form58.fm2;
 gOk:=form58.TOk;
 form58.free;
 if not gOk then exit;
}

    form3327:=TForm3327.Create(Self);
    form3327.ShowModal;
    gOk:=form3327.GOk;
    if gOk then
     begin
      datstart:=form3327.jvDateEdit1.Date;
      datend:=form3327.jvDateEdit2.Date;
     end;
    form3327.free;
    if not gOk then exit;

 Button1Click(nil);

{
  datstart:=EncodeDate(RGod,gm1,1);
  if gm2<>12 then datend:=EncodeDate(RGod,gm2+1,1)-1 else datend:=EncodeDate(RGod,12,31);
}

  srokdat1:=datstart;  //запоминаем т.к. запрос по даты выплаты а нам по сроку нужно уплаты
  srokdat2:=datend;
  datstart:=datstart-60;  //+-60 Дн срок в любом случае от даты выплаты
  datend:=datend+60;


           ds:=form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER)+'.dbf';
           if tmprepdox.Active then tmprepdox.Active:=False;
           if tmprepdox.Exists then tmprepdox.DeleteTable;
           tmprepdox.TableName:=ds;
           tmprepdox.DatabaseName:=form1.DBDIR;
           tmprepdox.Exclusive:=True;
           tmprepdox.TableType:=ttDBase;
           tmprepdox.FieldDefs.Clear;

           tmprepdox.FieldDefs.Add('kod',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('koddox',ftString,10); {оклад}
           tmprepdox.FieldDefs.Add('summa',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('wm',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('wg',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('ndfl',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('dat',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat2',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat3',ftDate,0,false);
           tmprepdox.FieldDefs.Add('g',ftString,1); {для фильтра}
           tmprepdox.CreateTable;
           Reindex.ReindexTab(tmprepdox,form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER),'dat','');
           tmprepdox.IndexName:='dat';



               datam.qtmp.close;
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select s.* from sdoxod s where' );
               datam.qtmp.SQL.Add('s.dat>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
               datam.qtmp.SQL.Add(' and s.dat<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                    if datam.qTmp.fieldbyname('kodnac').asfloat=0 then
                       FDatnacisl('2000',datam.qTmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3)
                              else
                          begin
                           form1.NACISL.Locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInSensitive]);
                           FDatnacisl(form1.NACISLKODDOX.Value,datam.Qtmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3)
                          end;

                     x2:=datam.Qtmp.Fieldbyname('sdoxod').asFloat;
                     xNdfl:=datam.Qtmp.Fieldbyname('nalog').asFloat;



                  bRTF:=True;

                   if (datam.qTmp.Fieldbyname('god').asInteger<=2022) and (datam.qTmp.FieldByNAme('dat').asDateTime>=EncodeDate(2023,1,1)) then
                      begin
                         datam.qtmpstaj.close;
                         datam.qtmpstaj.databasename:=form1.dbdir;
                         datam.qtmpstaj.sql.clear;
                         datam.qtmpstaj.sql.add('select * from obrt2new where nls='+floattostr(datam.qtmp.fieldbyname('nls').asfloat));
                         datam.qtmpstaj.sql.add('and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
                         datam.qtmpstaj.prepare;
                         datam.qtmpstaj.open;
                         if datam.qtmpstaj.RecordCount>=1 then
                          begin
                            if datam.qtmpstaj.fieldbyname('datprov').asdatetime>=EncodeDate(2023,1,1) then tdat1:=datam.qtmpstaj.fieldbyname('datprov').asdatetime;
                          end;
                         datam.qtmpstaj.close;
                       end;

                    if datam.qTmp.fieldbyname('kodnac').asfloat<>0 then  //не облагается ндфл убрать
                     begin
                      form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInsensitive]);
                      if form1.nacislPN.Value=1 then bRTF:=False;
                     end;


                   if datam.qTmp.fieldbyname('kodnac').asfloat=0 then     //аванс пустой без начисления убрать !
                      begin
                         datam.Query1.Close;
                         datam.Query1.SQL.Clear;
                         datam.Query1.SQL.Add('select g.* from glnew g, kart k where k.nls=g.nls ');
                         datam.Query1.SQL.Add(' and g.wm='+floattostr(datam.qTmp.fieldbyname('mes').asFloat));
                         datam.Query1.SQL.Add(' and g.wg='+floattostr(datam.qTmp.fieldbyname('god').asFloat));
                         datam.Query1.SQL.Add(' and g.nls='+floattostr(datam.qTmp.fieldbyname('nls').asFloat));
                         datam.Query1.SQL.Add(' and g.oklad*g.dayotr>0 and g.dayrab<>0');
                         datam.Query1.prepare;
                         datam.Query1.open;
                         if datam.Query1.RecordCount>=1 then bRTF:=True else bRTF:=False;
                      end;


                 if bRTF then
                  begin
                    if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([datam.qTmp.fieldbyname('mes').asInteger,datam.qTmp.fieldbyname('kodnac').asfloat,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
                     begin
                      tmprepdox.append;
                      tmprepdox.fieldbyname('kod').asFloat:=datam.qTmp.fieldbyname('kodnac').asfloat;
                      tmprepdox.fieldbyname('wm').asInteger:=datam.qTmp.fieldbyname('mes').asInteger;
                      tmprepdox.fieldbyname('wg').asInteger:=datam.qTmp.fieldbyname('god').asInteger;
                      tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
                      tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
                      tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
                      tmprepdox.post;
                      end;
                     tmprepdox.edit;
                     tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
                     tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
                     tmprepdox.post;
                  end;

                  datam.qtmp.next;
                 end;
                datam.qtmp.close;


 datam.Query1.Close;
 datam.Query1.SQL.Clear;
 datam.Query1.SQL.Add('select g.* from glnew g, kart k where k.nls=g.nls');
 datam.Query1.SQL.Add(' and g.datoklad>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
 datam.Query1.SQL.Add(' and g.datoklad<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);


 datam.Query1.Prepare;
 datam.Query1.Open;
 datam.Query1.First;
 while not datam.Query1.Eof do
  begin
    km:=datam.Query1.fieldByName('wm').asInteger;
    kg:=datam.Query1.fieldByName('wg').asInteger;

     x2:=DRound(datam.Query1.FieldByName('OKLAD').asFloat*DelenieCas(datam.Query1.FieldByName('DAYOTR').asFloat,
                    datam.Query1.FieldByName('DAYRAB').asFloat,datam.Query1.FieldByName('DAYCAS').asFloat),form1.DRZn) ;
            x2:=x2+DRound(x2*form1.configRK.Value/100,2);
            xndfl:=datam.Query1.fieldByName('snalog').asFloat;



               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac=0 and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;



   if DRound(x2,2)>0 then
    begin
     FdatNacisl('2000',km,kg,datam.Query1.FieldByName('datoklad').asDatetime,tdat1,tdat2,tdat3);

     if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([km,0,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
      begin
       tmprepdox.append;
       tmprepdox.fieldbyname('kod').asFloat:=0;
       tmprepdox.fieldbyname('wm').asInteger:=km;
       tmprepdox.fieldbyname('wg').asInteger:=kg;
       tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
       tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
       tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
       tmprepdox.post;
      end;
      tmprepdox.edit;
      tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
      tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
      tmprepdox.post;
    end;

   datam.Query1.Next;
  end;

 datam.Query1.Close;
 datam.Query1.SQL.Clear;
 datam.Query1.SQL.Add('select g.* from obrt1new g, kart k where k.nls=g.nls');
 datam.Query1.SQL.Add(' and g.datprov>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
 datam.Query1.SQL.Add(' and g.datprov<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);
 datam.Query1.Prepare;
 datam.Query1.Open;
 datam.Query1.First;
 while not datam.Query1.Eof do
  begin
    km:=datam.Query1.fieldByName('wm').asInteger;
    kg:=datam.Query1.fieldByName('wg').asInteger;

     x2:=datam.Query1.FieldByName('KR').asFloat ;
     form1.NACISL.Locate('kod',datam.Query1.FieldByName('kod').asInteger,[locaseInsensitive]);
     if form1.NACISLRK.Value then x2:=x2+DRound(x2*form1.configRK.Value/100,2);
     xndfl:=datam.Query1.fieldByName('snalog').asFloat;

               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac='+floattostr(datam.Query1.FieldByName('kod').asInteger)+' and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;


   if (DRound(x2,2)>0) and
          (   ((form1.NACISLpn.Value<>1) ) OR   ((form1.NACISLpn.Value=1) ) ) then
    begin

     FdatNacisl(form1.NACISLkoddox.Value,km,kg,datam.Query1.FieldByName('datprov').asDatetime,tdat1,tdat2,tdat3);
     if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([km,datam.Query1.FieldByName('kod').asInteger,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
      begin
       tmprepdox.append;
       tmprepdox.fieldbyname('kod').asInteger:=datam.Query1.FieldByName('kod').asInteger;
       tmprepdox.fieldbyname('wm').asInteger:=km;
       tmprepdox.fieldbyname('wg').asInteger:=kg;
       tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
       tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
       tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
       tmprepdox.post;
      end;
      tmprepdox.edit;
      tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
      tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
      tmprepdox.post;

    end;
   datam.Query1.Next;
  end;


  
   fNameXls:=GetNameXlsn('ОтчНДФЛ','r')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\reportdat3.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      //form_58.free;
      exit;
     end;
 
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    if tTypeDat=0 then E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Сводный отчет по срокам перечисления НДФЛ в период '+formatdatetime('dd.mm.yyyy',srokdat1)+' - '+formatdatetime('dd.mm.yyyy',srokdat2);
    if tTypeDat=1 then E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Сводный отчет об исчисленном НДФЛ за период '+formatdatetime('dd.mm.yyyy',srokdat1)+' - '+formatdatetime('dd.mm.yyyy',srokdat2);
    if tTypeDat=2 then E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Сводный отчет об удержанном НДФЛ за период '+formatdatetime('dd.mm.yyyy',srokdat1)+' - '+formatdatetime('dd.mm.yyyy',srokdat2);

    tmprepdox.first;
    Nstroka:=5;
    xs0:=0; xn0:=0;
    while not tmprepdox.eof do
     begin

      bRtf:=False;

      if (tTypeDat=0) and (tmprepdox.fieldbyname('dat3').asDateTime>=srokdat1) and (tmprepdox.fieldbyname('dat3').asDateTime<=srokdat2) then bRtf:=True; //срок уплаты
      if (tTypeDat=1) and (tmprepdox.fieldbyname('dat').asDateTime>=srokdat1) and (tmprepdox.fieldbyname('dat').asDateTime<=srokdat2) then bRtf:=True; //срок уплаты
      if (tTypeDat=2) and (tmprepdox.fieldbyname('dat2').asDateTime>=srokdat1) and (tmprepdox.fieldbyname('dat2').asDateTime<=srokdat2) then bRtf:=True; //срок уплаты



     IF BRTF THEN
      BEGIN
       Nstroka:=Nstroka+1;
       form1.NACISL.Locate('kod',tmprepdox.FieldByName('kod').asInteger,[locaseInsensitive]);
       E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)].Value:=copy(ansilowercase(namemes[tmprepdox.fieldbyname('wm').asInteger]),1,3)+' '+floattostr(tmprepdox.fieldbyname('wg').asInteger);
       if tmprepdox.FieldByName('kod').asInteger<>0 then
        begin
         E.ActiveWorkBook.Sheets.Item[1].Range['B'+IntToStr(Nstroka)].Value:=form1.nacisl.fieldbyname('name').asString;
         E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:=form1.nacisl.fieldbyname('koddox').asString;
        end
         else
        begin
         E.ActiveWorkBook.Sheets.Item[1].Range['B'+IntToStr(Nstroka)].Value:='оклад/тариф';
         E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:='2000';
        end;
       E.ActiveWorkBook.Sheets.Item[1].Range['D'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('summa').asFloat;
       xs0:=xs0+tmprepdox.fieldbyname('summa').asFloat;
       xn0:=xn0+tmprepdox.fieldbyname('ndfl').asFloat;

       if tmprepdox.fieldbyname('dat').asDateTime>EncodeDate(2000,1,1) then
           E.ActiveWorkBook.Sheets.Item[1].Range['F'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat').asDateTime);
       if tmprepdox.fieldbyname('dat2').asDateTime>EncodeDate(2000,1,1) then
           E.ActiveWorkBook.Sheets.Item[1].Range['G'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat2').asDateTime);
       E.ActiveWorkBook.Sheets.Item[1].Range['H'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('ndfl').asFloat;
       if tmprepdox.fieldbyname('dat').asDateTime>EncodeDate(2000,1,1) then
         E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat').asDateTime);
       if tmprepdox.fieldbyname('dat2').asDateTime>EncodeDate(2000,1,1) then
            E.ActiveWorkBook.Sheets.Item[1].Range['J'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat2').asDateTime);
       if tmprepdox.fieldbyname('dat3').asDateTime>EncodeDate(2000,1,1) then
          E.ActiveWorkBook.Sheets.Item[1].Range['K'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat3').asDateTime);
      END;
      tmprepdox.next;
     end;

        try
         for km:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A6','K'+IntToStr(NSTROKA)].Borders[km].LineStyle:=1;
        except
        end;

        try
         if (tTypeDat=0) then
          begin
           E.ActiveWorkbook.Sheets.Item[1].Range['K4','K'+IntToStr(NSTROKA)].Font.Color:=clNavy;
           E.ActiveWorkbook.Sheets.Item[1].Range['K4','K'+IntToStr(NSTROKA)].Font.Italic:=True;
          end;
         if (tTypeDat=1) then
          begin
           E.ActiveWorkbook.Sheets.Item[1].Range['I4','I'+IntToStr(NSTROKA)].Font.Color:=clNavy;
           E.ActiveWorkbook.Sheets.Item[1].Range['I4','I'+IntToStr(NSTROKA)].Font.Italic:=True;
          end;
         if (tTypeDat=2) then
          begin
           E.ActiveWorkbook.Sheets.Item[1].Range['J4','J'+IntToStr(NSTROKA)].Font.Color:=clNavy;
           E.ActiveWorkbook.Sheets.Item[1].Range['J4','J'+IntToStr(NSTROKA)].Font.Italic:=True;
          end;
        except
        end;

      Nstroka:=Nstroka+1;
      E.ActiveWorkBook.Sheets.Item[1].Range['D'+IntToStr(Nstroka)].Value:=xs0;
      E.ActiveWorkBook.Sheets.Item[1].Range['H'+IntToStr(Nstroka)].Value:=xn0;

    E.Application.WindowState:=1;
    E.WindowState:=-4137 ;
    E:=UnAssigned;

end;

procedure TForm_58.FormRepSrok22(tExcel:Boolean;tTypeDat:Integer);  //с выводом Excel либо просто остановиться после формирования tmp таблицы
var ds,soktmo:String;
    gm1,gm2:Integer;
    gok:Boolean;
    km,kg:integer;
    tdat1,tdat2,tdat3:TDate;
    x2,xid:Real;
    fNameXLS:String;
    Nstroka:Integer;
    E:OleVariant;
    sname:string;
    xs0,xn0,xndfl:Real;
    datstart,datend, srokdat1, srokdat2:TDate;
    bRTF:Boolean;
    TOK:integer;
    vId:Integer;
begin



 if not tExcel then
  begin
   form3305:=TForm3305.Create(Self);
  // if not tExcel then form3305.tSrok2:=false;
   form3305.ShowModal;
   gOk:=form3305.GOk;
   if gOk then
    begin
     datstart:=form3305.tD3;
     datend:=form3305.tD3;
     vId:=form3305.TId;
    end;
   form3305.free;
   if not gOk then exit;
   if vId<=0 then exit; //остатки на начало по другому
  end;


  if (tExcel) then
   begin
    form3327:=TForm3327.Create(Self);
    form3327.ShowModal;
    gOk:=form3327.GOk;
    if gOk then
     begin
      datstart:=form3327.jvDateEdit1.Date;
      datend:=form3327.jvDateEdit2.Date;
     end;
    form3327.free;
    if not gOk then exit;
   end;

{ form58:=TForm58.Create(Self);
 form58.ShowModal;
 gm1:=form58.fm1;
 gm2:=form58.fm2;
 gOk:=form58.TOk;
 form58.free;
 if not gOk then exit;
}
 Button1Click(nil);


 // datstart:=EncodeDate(RGod,gm1,1);
 // if gm2<>12 then datend:=EncodeDate(RGod,gm2+1,1)-1 else datend:=EncodeDate(RGod,12,31);


  srokdat1:=datstart;  //запоминаем т.к. запрос по даты выплаты а нам по сроку нужно уплаты
  srokdat2:=datend;
  datstart:=datstart-60;  //+-60 Дн срок в любом случае от даты выплаты
  datend:=datend+60;


           ds:=form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER)+'.dbf';
           if tmprepdox.Active then tmprepdox.Active:=False;
           if tmprepdox.Exists then tmprepdox.DeleteTable;
           tmprepdox.TableName:=ds;
           tmprepdox.DatabaseName:=form1.DBDIR;
           tmprepdox.Exclusive:=True;
           tmprepdox.TableType:=ttDBase;
           tmprepdox.FieldDefs.Clear;
           tmprepdox.FieldDefs.Add('id',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('brtf',ftFloat,0,false);       //для фильтра
           tmprepdox.FieldDefs.Add('uroven',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('nls',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('summa',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('ndfl',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('mes',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('god',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('dat',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat3',ftDate,0,false);
           tmprepdox.FieldDefs.Add('oktmo',ftString,15); {для фильтра}
           tmprepdox.CreateTable;
           Reindex.ReindexTab(tmprepdox,form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER),'dat','');
           tmprepdox.IndexName:='dat';

           xId:=0;



               datam.qtmp.close;
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select k.*, s.* from kart k, sdoxod s where k.nls=s.nls and' );
               datam.qtmp.SQL.Add('s.dat>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
               datam.qtmp.SQL.Add(' and s.dat<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);
               datam.qtmp.SQL.Add('order by s.dat, s.mes, s.god, k.fam,k.im,k.ot');
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=datam.Qtmp.Fieldbyname('sdoxod').asFloat;
                   xNdfl:=datam.Qtmp.Fieldbyname('nalog').asFloat;
                   tdat1:=datam.qtmp.fieldbyname('dat').asdatetime;

                   //
                    if datam.qTmp.fieldbyname('kodnac').asfloat=0 then
                       FDatnacisl('2000',datam.qTmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3)
                              else
                          begin
                           form1.NACISL.Locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInSensitive]);
                           FDatnacisl(form1.NACISLKODDOX.Value,datam.Qtmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3);

                          end;

          //       ShowMessage(datetostr(datstart)+' '+datetostr(datend)+#13+datetostr(tdat1)+' '+datetostr(tdat2)+' '+datetostr(tdat3));
                  //  tdat1:=tdat2; //дата выплаты
                   // tdat3:=form_58.FGetDatPerecisl(tdat1);

                    bRTF:=True;

                     if (datam.qTmp.Fieldbyname('god').asInteger<=2022) and (datam.qTmp.FieldByNAme('dat').asDateTime>=EncodeDate(2023,1,1)) then
                      begin
                         datam.qtmpstaj.close;
                         datam.qtmpstaj.databasename:=form1.dbdir;
                         datam.qtmpstaj.sql.clear;
                         datam.qtmpstaj.sql.add('select * from obrt2new where nls='+floattostr(datam.qtmp.fieldbyname('nls').asfloat));
                         datam.qtmpstaj.sql.add('and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
                         datam.qtmpstaj.prepare;
                         datam.qtmpstaj.open;
                         if datam.qtmpstaj.RecordCount>=1 then
                          begin
                            if datam.qtmpstaj.fieldbyname('datprov').asdatetime>=EncodeDate(2023,1,1) then tdat1:=datam.qtmpstaj.fieldbyname('datprov').asdatetime;
                          end;
                         datam.qtmpstaj.close;
                       end;

                   { ShowMessage(floattostr(datam.qtmp.fieldbyname('sdoxod').asfloat)+' '+datetostr(datam.qtmp.fieldbyname('dat').asdatetime)+#13+
                       datetostr(tdat1)+' '+datetostr(tdat2)+' '+datetostr(tdat3));
                   }
                   
                    if datam.qTmp.fieldbyname('kodnac').asfloat<>0 then  //не облагается ндфл убрать
                     begin
                      form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInsensitive]);
                      if form1.nacislPN.Value=1 then bRTF:=False;
                     end;

                    if not tExcel then
                     begin
                      if (datam.qTmp.fieldbyname('god').asfloat<RGod) and (vId>=1) then  //выплата за 2022
                       begin
                        bRTF:=False;
                       end;
                     end;  

                   if tExcel then
                     begin
                      if (datam.qTmp.fieldbyname('god').asfloat<=2022) then  //выплата за 2022
                       begin
                       //  отсюда взять код для разделения отчет по дате выплаты дохода для разделения аванса
                         datam.qtmpstaj.close;
                         datam.qtmpstaj.databasename:=form1.dbdir;
                         datam.qtmpstaj.sql.clear;
                         datam.qtmpstaj.sql.add('select * from obrt2new where nls='+floattostr(datam.qtmp.fieldbyname('nls').asfloat));
                         datam.qtmpstaj.sql.add('and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
                         datam.qtmpstaj.prepare;
                         datam.qtmpstaj.open;
                         if datam.qtmpstaj.RecordCount>=1 then
                          begin
                           // if datam.qtmpstaj.fieldbyname('datprov').asdatetime<EncodeDate(RGod,1,1) then  bRTF:=False;
                          end;
                         datam.qtmpstaj.close;
                       end;
                     end;
                  
                      form1.kart.locate('nls',datam.qtmp.fieldbyname('nls').asFloat,[locaseinsensitive]);
                      TOK:=Form_58.ZapolnDOK(form1.kartnls.value);
                      soktmo:=form_58.FGetOktmo(Trunc(datam.qtmp.fieldbyname('mes').asFloat),Trunc(datam.qtmp.fieldbyname('god').asFloat));
                      soktmo:=trim(soktmo);

                 if bRTF then
                  begin
                   if not tmprepdox.Locate('mes;god;dat;dat3;oktmo',VarArrayOf([datam.qTmp.fieldbyname('mes').asFloat,
                      datam.qTmp.fieldbyname('god').asFloat,
                       tdat1,tdat3,soktmo]),[loCaseInsensitive]) then
                     begin
                      tmprepdox.append;
                      xid:=xid+1;
                      tmprepdox.fieldbyname('id').asFloat:=xid;
                      tmprepdox.fieldbyname('uroven').asFloat:=1;
                      tmprepdox.fieldbyname('mes').asFloat:=datam.qTmp.fieldbyname('mes').asfloat;
                      tmprepdox.fieldbyname('god').asFloat:=datam.qTmp.fieldbyname('god').asfloat;
                      tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
                      tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
                      tmprepdox.fieldbyname('oktmo').asString:=soktmo;
                      tmprepdox.post;
                     end;
                      tmprepdox.edit;
                      tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
                      tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
                      tmprepdox.post;

                   if tExcel then //для Excel выводим всю информацию без группировки
                    begin
                      if not tmprepdox.Locate('nls;mes;god;dat;dat3',VarArrayOf([datam.qTmp.fieldbyname('nls').asFloat,datam.qTmp.fieldbyname('mes').asFloat,
                         datam.qTmp.fieldbyname('god').asFloat,
                           tdat1,tdat3]),[loCaseInsensitive]) then
                        begin
                         tmprepdox.append;
                         xid:=xid+1;
                         tmprepdox.fieldbyname('id').asFloat:=xid;
                         tmprepdox.fieldbyname('uroven').asFloat:=2;
                         tmprepdox.fieldbyname('nls').asFloat:=datam.qTmp.fieldbyname('nls').asfloat;
                         tmprepdox.fieldbyname('mes').asFloat:=datam.qTmp.fieldbyname('mes').asfloat;
                         tmprepdox.fieldbyname('god').asFloat:=datam.qTmp.fieldbyname('god').asfloat;
                         tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
                         tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
                         tmprepdox.fieldbyname('oktmo').asString:=soktmo;
                         tmprepdox.post;
                        end;
                         tmprepdox.edit;
                         tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
                         tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
                         tmprepdox.post;
                     end;

                   if not tExcel then //для уплаты группируем
                    begin
                      if not tmprepdox.Locate('nls;mes;god;dat3',VarArrayOf([datam.qTmp.fieldbyname('nls').asFloat,datam.qTmp.fieldbyname('mes').asFloat,
                         datam.qTmp.fieldbyname('god').asFloat,
                           tdat3]),[loCaseInsensitive]) then
                        begin
                         tmprepdox.append;
                         xid:=xid+1;
                         tmprepdox.fieldbyname('id').asFloat:=xid;
                         tmprepdox.fieldbyname('uroven').asFloat:=2;
                         tmprepdox.fieldbyname('nls').asFloat:=datam.qTmp.fieldbyname('nls').asfloat;
                         tmprepdox.fieldbyname('mes').asFloat:=datam.qTmp.fieldbyname('mes').asfloat;
                         tmprepdox.fieldbyname('god').asFloat:=datam.qTmp.fieldbyname('god').asfloat;
                         tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
                         tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
                         tmprepdox.fieldbyname('oktmo').asString:=soktmo;
                         tmprepdox.post;
                        end;
                         tmprepdox.edit;
                         tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
                         tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
                         tmprepdox.post;
                     end;

                  end;

                  datam.qtmp.next;
                 end;
                datam.qtmp.close;


    tmprepdox.first;
    while not tmprepdox.eof do
     begin
       tmprepdox.edit;
       tmprepdox.fieldbyname('brtf').asfloat:=0;
       tmprepdox.post;
       brtf:=false;

       if (tTypeDat=0) and (tmprepdox.fieldbyname('dat3').asDateTime>=srokdat1) and (tmprepdox.fieldbyname('dat3').asDateTime<=srokdat2) then bRtf:=True;  //по сроку
       if (tTypeDat=1) and (tmprepdox.fieldbyname('dat').asDateTime>=srokdat1) and (tmprepdox.fieldbyname('dat').asDateTime<=srokdat2) then bRtf:=True;  //исчисл

        if (BRtf) then  //только ФИО оставляем
         begin
          tmprepdox.edit;
          tmprepdox.fieldbyname('brtf').asfloat:=1;
          tmprepdox.post;
        end;
      tmprepdox.next;
     end;

   if not tExcel then EXIT; //выходим если только tmprepdox нужна
     
   fNameXls:=GetNameXlsn('ОтчДоходНДФЛ','r')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\reportdat2.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      //form_58.free;
      exit;
     end;
 
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

//    E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Отчет по сроку уплаты НДФЛ: '+ansilowercase(namemes[gm1]+' - '+namemes[gm2])+' '+floattostr(RGod)+' г.';

    if tTypeDat=0 then E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Отчет о сроке перечисления НДФЛ в период '+formatdatetime('dd.mm.yyyy', srokdat1)+'-'+formatdatetime('dd.mm.yyyy',srokdat2);
    if tTypeDat=1 then E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:='Отчет об исчисленном НДФЛ за период '+formatdatetime('dd.mm.yyyy', srokdat1)+'-'+formatdatetime('dd.mm.yyyy',srokdat2);


    tmprepdox.first;
    Nstroka:=5;
    xs0:=0; xn0:=0;
    while not tmprepdox.eof do
     begin

      IF tmprepdox.fieldbyname('brtf').asFloat=1 THEN
      BEGIN


       Nstroka:=Nstroka+1;
       E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)].Value:=copy(ansilowercase(namemes[tmprepdox.fieldbyname('mes').asInteger]),1,3)
             +' '+floattostr(tmprepdox.fieldbyname('god').asInteger);

       if tmprepdox.fieldbyname('uroven').asfloat=2 then
        begin
         form1.kart.Locate('nls',tmprepdox.FieldByName('nls').asfloat,[locaseInsensitive]);
         E.ActiveWorkBook.Sheets.Item[1].Range['B'+IntToStr(Nstroka)].Value:=form1.kartfam.value+' '+copy(form1.kartim.value,1,1)+'.'+copy(form1.kartot.value,1,1)+'.';
        end;


       E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('summa').asFloat;
        if tmprepdox.fieldbyname('uroven').asfloat=1 then
          begin
           xs0:=xs0+tmprepdox.fieldbyname('summa').asFloat;
           xn0:=xn0+tmprepdox.fieldbyname('ndfl').asFloat;
          end;
          
       E.ActiveWorkBook.Sheets.Item[1].Range['F'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('ndfl').asFloat;

        E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('summa').asFloat-tmprepdox.fieldbyname('ndfl').asFloat;

       if tmprepdox.fieldbyname('dat').asDateTime>EncodeDate(2000,1,1) then
         E.ActiveWorkBook.Sheets.Item[1].Range['E'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat').asDateTime);
       if tmprepdox.fieldbyname('dat3').asDateTime>EncodeDate(2000,1,1) then
          E.ActiveWorkBook.Sheets.Item[1].Range['G'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat3').asDateTime);
       E.ActiveWorkBook.Sheets.Item[1].Range['H'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('oktmo').asstring;

       try
        if tmprepdox.fieldbyname('uroven').asfloat=2 then
         begin
          E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)+':'+'I'+IntToStr(Nstroka)].Font.Italic:=true;
         end;
         if tmprepdox.fieldbyname('uroven').asfloat=1 then
         begin
         //E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)+':'+'I'+IntToStr(Nstroka)].Font.Size:=10;
          E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)+':'+'I'+IntToStr(Nstroka)].Interior.Color:=WColorRep;
         end;
       except
       end;

      END;
      tmprepdox.next;
     end;

        try
         for km:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A6','I'+IntToStr(NSTROKA)].Borders[km].LineStyle:=1;
        except
        end;

        try
         if tTypeDat=0 then
          begin
           E.ActiveWorkbook.Sheets.Item[1].Range['G4','G'+IntToStr(NSTROKA)].Font.Color:=clNavy;
           E.ActiveWorkbook.Sheets.Item[1].Range['G4','G'+IntToStr(NSTROKA)].Font.Italic:=True;
          end;
         if tTypeDat=1 then
          begin
           E.ActiveWorkbook.Sheets.Item[1].Range['E4','E'+IntToStr(NSTROKA)].Font.Color:=clNavy;
           E.ActiveWorkbook.Sheets.Item[1].Range['E4','E'+IntToStr(NSTROKA)].Font.Italic:=True;
          end;

        except
        end;

      Nstroka:=Nstroka+1;
      E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:=xs0;
      E.ActiveWorkBook.Sheets.Item[1].Range['F'+IntToStr(Nstroka)].Value:=xn0;
      E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(Nstroka)].Value:=xs0-xn0;


    E.Application.WindowState:=1;
    E.WindowState:=-4137 ;
    E:=UnAssigned;

     form_58.tmprepdox.Active:=false;

end;



procedure TForm_58.FormRepDatNew(formtype,tnevyp:integer);
var ds:String;
    gm1,gm2:Integer;
    gok:Boolean;
    km,kg:integer;
    tdat1,tdat2,tdat3:TDate;
    x2:Real;
    fNameXLS:String;
    Nstroka:Integer;
    E:OleVariant;
    sname:string;
    xs0,xn0,xndfl:Real;
    datstart,datend:TDate;
    bRTF:Boolean;
begin


    form3327:=TForm3327.Create(Self);
    form3327.ShowModal;
    gOk:=form3327.GOk;
    if gOk then
     begin
      datstart:=form3327.jvDateEdit1.Date;
      datend:=form3327.jvDateEdit2.Date;
     end;
    form3327.free;
    if not gOk then exit;

{
 form58:=TForm58.Create(Self);
 form58.ShowModal;
 gm1:=form58.fm1;
 gm2:=form58.fm2;
 gOk:=form58.TOk;
 form58.free;
 if not gOk then exit;
}
 Button1Click(nil);


{  datstart:=EncodeDate(RGod,gm1,1);
  if gm2<>12 then datend:=EncodeDate(RGod,gm2+1,1)-1 else datend:=EncodeDate(RGod,12,31);
}
           ds:=form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER)+'.dbf';
           if tmprepdox.Active then tmprepdox.Active:=False;
           if tmprepdox.Exists then tmprepdox.DeleteTable;
           tmprepdox.TableName:=ds;
           tmprepdox.DatabaseName:=form1.DBDIR;
           tmprepdox.Exclusive:=True;
           tmprepdox.TableType:=ttDBase;
           tmprepdox.FieldDefs.Clear;

           tmprepdox.FieldDefs.Add('kod',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('koddox',ftString,10); {оклад}
           tmprepdox.FieldDefs.Add('summa',ftFloat,0,false); {оклад}
           tmprepdox.FieldDefs.Add('wm',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('wg',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('ndfl',ftFloat,0,false);
           tmprepdox.FieldDefs.Add('dat',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat2',ftDate,0,false);
           tmprepdox.FieldDefs.Add('dat3',ftDate,0,false);
           tmprepdox.FieldDefs.Add('g',ftString,1); {для фильтра}
           tmprepdox.CreateTable;
           Reindex.ReindexTab(tmprepdox,form1.DBDIR+'\tpdoxx'+floattostr(datam.TUSER),'dat','');
           tmprepdox.IndexName:='dat';



               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select s.* from sdoxod s where' );
               if (formtype=0) or (formtype=3)  then
                 begin
                  datam.qtmp.SQL.Add('s.dat>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
                  datam.qtmp.SQL.Add(' and s.dat<='+#39+FormatDateTime('dd.mm.yyyy',datend+100)+#39);   //учитывает переход 2022 - 2023 выплата в 2023 за 2022 
                 end;

               
               if (formtype=1) or (formtype=4) then
                 begin
                  datam.qtmp.SQL.Add(' s.mes>='+floattostr(gm1));
                  datam.qtmp.SQL.Add('and s.mes<='+floattostr(gm2)+' and s.god='+floattostr(rGod));
                 end;

               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                    if datam.qTmp.fieldbyname('kodnac').asfloat=0 then
                       FDatnacisl('2000',datam.qTmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3)
                              else
                          begin
                           form1.NACISL.Locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInSensitive]);
                           FDatnacisl(form1.NACISLKODDOX.Value,datam.Qtmp.Fieldbyname('mes').asInteger,datam.qTmp.Fieldbyname('god').asInteger,
                                                       datam.qTmp.FieldByNAme('dat').asDateTime,tdat1,tDat2,tDat3)
                          end;

                     x2:=datam.Qtmp.Fieldbyname('sdoxod').asFloat;
                     xNdfl:=datam.Qtmp.Fieldbyname('nalog').asFloat;

                   bRTF:=True;
                    
                     if (datam.qTmp.Fieldbyname('god').asInteger<=2022) and (datam.qTmp.FieldByNAme('dat').asDateTime>=EncodeDate(2023,1,1)) then
                      begin
                         datam.qtmpstaj.close;
                         datam.qtmpstaj.databasename:=form1.dbdir;
                         datam.qtmpstaj.sql.clear;
                         datam.qtmpstaj.sql.add('select * from obrt2new where nls='+floattostr(datam.qtmp.fieldbyname('nls').asfloat));
                         datam.qtmpstaj.sql.add('and id='+floattostr(datam.qtmp.fieldbyname('idvypl').asfloat));
                         datam.qtmpstaj.prepare;
                         datam.qtmpstaj.open;
                         if datam.qtmpstaj.RecordCount>=1 then
                          begin
                            if datam.qtmpstaj.fieldbyname('datprov').asdatetime>=EncodeDate(2023,1,1) then tdat1:=datam.qtmpstaj.fieldbyname('datprov').asdatetime;
                          end;
                         datam.qtmpstaj.close;
                       end;

               //  ShowMessage(floattostr(datam.qTmp.Fieldbyname('sdoxod').asInteger)+' '+datetostr(datam.qTmp.Fieldbyname('dat').asdatetime)+#13+' '+datetostr(tdat1)+' '+datetostr(tdat2)+' '+datetostr(tdat3));

                  if (formtype=0) or (formtype=1) then
                   begin
                    if datam.qTmp.fieldbyname('kodnac').asfloat<>0 then  //не облагается ндфл убрать
                     begin
                      form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInsensitive]);
                      if form1.nacislPN.Value=1 then bRTF:=False;
                     end;
                   end;
                  if (formtype=3) or (formtype=4) then
                   begin
                    if datam.qTmp.fieldbyname('kodnac').asfloat<>0 then  //не облагается ндфл убрать
                     begin
                      form1.nacisl.locate('kod',datam.qTmp.fieldbyname('kodnac').asfloat,[loCaseInsensitive]);
                      if form1.nacislPN.Value<>1 then bRTF:=False;
                     end;
                   end;


                   if datam.qTmp.fieldbyname('kodnac').asfloat=0 then     //аванс пустой без начисления убрать !
                      begin
                         datam.Query1.Close;
                         datam.Query1.SQL.Clear;
                         datam.Query1.SQL.Add('select g.* from glnew g, kart k where k.nls=g.nls ');
                         datam.Query1.SQL.Add(' and g.wm='+floattostr(datam.qTmp.fieldbyname('mes').asFloat));
                         datam.Query1.SQL.Add(' and g.wg='+floattostr(datam.qTmp.fieldbyname('god').asFloat));
                         datam.Query1.SQL.Add(' and g.nls='+floattostr(datam.qTmp.fieldbyname('nls').asFloat));
                         datam.Query1.SQL.Add(' and g.oklad*g.dayotr>0 and g.dayrab<>0');
                         datam.Query1.prepare;
                         datam.Query1.open;
                         if datam.Query1.RecordCount>=1 then bRTF:=True else bRTF:=False;
                      end;

                  if ((formtype=3) or (formtype=4)) and (datam.qTmp.fieldbyname('kodnac').asfloat=0) then bRTF:=False;

                 if bRTF then
                  begin
                    if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([datam.qTmp.fieldbyname('mes').asInteger,datam.qTmp.fieldbyname('kodnac').asfloat,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
                     begin
                      tmprepdox.append;
                      tmprepdox.fieldbyname('kod').asFloat:=datam.qTmp.fieldbyname('kodnac').asfloat;
                      tmprepdox.fieldbyname('wm').asInteger:=datam.qTmp.fieldbyname('mes').asInteger;
                      tmprepdox.fieldbyname('wg').asInteger:=datam.qTmp.fieldbyname('god').asInteger;
                      tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
                      tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
                      tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
                      tmprepdox.post;
                      end;
                     tmprepdox.edit;
                     tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
                     tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
                     tmprepdox.post;
                  end;

                  datam.qtmp.next;
                 end;
                datam.qtmp.close;


 datam.Query1.Close;
 datam.Query1.SQL.Clear;
 datam.Query1.SQL.Add('select g.* from glnew g, kart k where k.nls=g.nls');
 if (formtype=1) then
  begin
   datam.Query1.SQL.Add(' and g.wm>='+floattostr(gm1));
   datam.Query1.SQL.Add('and g.wm<='+floattostr(gm2)+' and g.wg='+floattostr(rGod));
  end;
 if (formtype=0) then  //!!!!!datoklad не заполяется, никогда не попадает лишний блок
  begin
   datam.Query1.SQL.Add(' and g.datoklad>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
   datam.Query1.SQL.Add(' and g.datoklad<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);
  end;
 if (formtype=3) or (formtype=4) then
   begin
    datam.Query1.SQL.Add(' and g.nls=987548');
   end;


 datam.Query1.Prepare;
 datam.Query1.Open;
 datam.Query1.First;
 while not datam.Query1.Eof do
  begin
    km:=datam.Query1.fieldByName('wm').asInteger;
    kg:=datam.Query1.fieldByName('wg').asInteger;

     x2:=DRound(datam.Query1.FieldByName('OKLAD').asFloat*DelenieCas(datam.Query1.FieldByName('DAYOTR').asFloat,
                    datam.Query1.FieldByName('DAYRAB').asFloat,datam.Query1.FieldByName('DAYCAS').asFloat),form1.DRZn) ;
            x2:=x2+DRound(x2*form1.configRK.Value/100,2);
            xndfl:=datam.Query1.fieldByName('snalog').asFloat;



               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac=0 and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;



   if DRound(x2,2)>0 then
    begin
     FdatNacisl('2000',km,kg,datam.Query1.FieldByName('datoklad').asDatetime,tdat1,tdat2,tdat3);

     if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([km,0,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
      begin
       tmprepdox.append;
       tmprepdox.fieldbyname('kod').asFloat:=0;
       tmprepdox.fieldbyname('wm').asInteger:=km;
       tmprepdox.fieldbyname('wg').asInteger:=kg;
       tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
       tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
       tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
       tmprepdox.post;
      end;
      tmprepdox.edit;
      tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
      tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
      tmprepdox.post;
    end;

   datam.Query1.Next;
  end;

 datam.Query1.Close;
 datam.Query1.SQL.Clear;
 datam.Query1.SQL.Add('select g.* from obrt1new g, kart k where k.nls=g.nls');

 if (formtype=1) or (formtype=4)  then
  begin
   datam.Query1.SQL.Add('and g.wm>='+floattostr(gm1)+' and g.wm<='+floattostr(gm2)+' and g.wg='+floattostr(rGod));
  end;
 if (formtype=0)  or (formtype=3)  then  //datprov не заполняется, никогда не попадет лишний блок
  begin
   datam.Query1.SQL.Add(' and g.datprov>='+#39+FormatDateTime('dd.mm.yyyy',datstart)+#39);
   datam.Query1.SQL.Add(' and g.datprov<='+#39+FormatDateTime('dd.mm.yyyy',datend)+#39);
  end;

 datam.Query1.Prepare;
 datam.Query1.Open;
 datam.Query1.First;
 while not datam.Query1.Eof do
  begin
    km:=datam.Query1.fieldByName('wm').asInteger;
    kg:=datam.Query1.fieldByName('wg').asInteger;

     x2:=datam.Query1.FieldByName('KR').asFloat ;
     form1.NACISL.Locate('kod',datam.Query1.FieldByName('kod').asInteger,[locaseInsensitive]);
     if form1.NACISLRK.Value then x2:=x2+DRound(x2*form1.configRK.Value/100,2);
     xndfl:=datam.Query1.fieldByName('snalog').asFloat;

               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac='+floattostr(datam.Query1.FieldByName('kod').asInteger)+' and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   x2:=x2-datam.qtmp.fieldbyname('sdoxod').asFloat;
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;


   if (DRound(x2,2)>0) and
          (   ((form1.NACISLpn.Value<>1) and (formtype<=1)) OR   ((form1.NACISLpn.Value=1) and (formtype>=3)) ) then
    begin

     FdatNacisl(form1.NACISLkoddox.Value,km,kg,datam.Query1.FieldByName('datprov').asDatetime,tdat1,tdat2,tdat3);
     if not tmprepdox.Locate('wm;kod;dat;dat2;dat3',VarArrayOf([km,datam.Query1.FieldByName('kod').asInteger,tdat1,tdat2,tdat3]),[loCaseInsensitive]) then
      begin
       tmprepdox.append;
       tmprepdox.fieldbyname('kod').asInteger:=datam.Query1.FieldByName('kod').asInteger;
       tmprepdox.fieldbyname('wm').asInteger:=km;
       tmprepdox.fieldbyname('wg').asInteger:=kg;
       tmprepdox.fieldbyname('dat').asDAteTime:=tdat1;
       tmprepdox.fieldbyname('dat2').asDAteTime:=tdat2;
       tmprepdox.fieldbyname('dat3').asDAteTime:=tdat3;
       tmprepdox.post;
      end;
      tmprepdox.edit;
      tmprepdox.fieldbyname('summa').asFloat:=tmprepdox.fieldbyname('summa').asFloat+x2;
      tmprepdox.fieldbyname('ndfl').asFloat:=tmprepdox.fieldbyname('ndfl').asFloat+xndfl;
      tmprepdox.post;

    end;
   datam.Query1.Next;
  end;


  
   fNameXls:=GetNameXlsn('ОтчНДФЛ','r')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\reportdat.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      //form_58.free;
      exit;
     end;
 
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    s:='';
    s:='Отчет о начисленном доходе за период ';
    if formtype=1 then s:='Отчет о начислении дохода при выплате заработной платы за период ';
 //   if formtype=0 then s:='Отчет о выплаченном доходе в период ';

    E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:=s+formatdatetime('dd.mm.yyyy',datstart)+' - '+formatdatetime('dd.mm.yyyy',datend);

    tmprepdox.first;
    Nstroka:=5;
    xs0:=0; xn0:=0;
    while not tmprepdox.eof do
     begin

      bRtf:=True;
      if tnevyp=0 then bRtf:=True;
      if (tnevyp=1) and (tmprepdox.fieldbyname('dat2').asDateTime>=EncodeDate(2010,1,1)) then bRtf:=False;  //не выплачен доход

      if not ((tmprepdox.fieldbyname('dat').asDateTime>=datstart) and (tmprepdox.fieldbyname('dat').asDateTime<=datend)) then bRtf:=false; //!!!!!Дата начисления дохода только 

      IF BRTF THEN
      BEGIN
       Nstroka:=Nstroka+1;
       form1.NACISL.Locate('kod',tmprepdox.FieldByName('kod').asInteger,[locaseInsensitive]);
       E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstroka)].Value:=copy(ansilowercase(namemes[tmprepdox.fieldbyname('wm').asInteger]),1,3)+' '+floattostr(tmprepdox.fieldbyname('wg').asInteger);
       if tmprepdox.FieldByName('kod').asInteger<>0 then
        begin
         E.ActiveWorkBook.Sheets.Item[1].Range['B'+IntToStr(Nstroka)].Value:=form1.nacisl.fieldbyname('name').asString;
         E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:=form1.nacisl.fieldbyname('koddox').asString;
        end
         else
        begin
         E.ActiveWorkBook.Sheets.Item[1].Range['B'+IntToStr(Nstroka)].Value:='оклад/тариф';
         E.ActiveWorkBook.Sheets.Item[1].Range['C'+IntToStr(Nstroka)].Value:='2000';
        end;
       E.ActiveWorkBook.Sheets.Item[1].Range['D'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('summa').asFloat;
       xs0:=xs0+tmprepdox.fieldbyname('summa').asFloat;
       xn0:=xn0+tmprepdox.fieldbyname('ndfl').asFloat;

       if tmprepdox.fieldbyname('dat').asDateTime>EncodeDate(2000,1,1) then
           E.ActiveWorkBook.Sheets.Item[1].Range['F'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat').asDateTime);
       if tmprepdox.fieldbyname('dat2').asDateTime>EncodeDate(2000,1,1) then
           E.ActiveWorkBook.Sheets.Item[1].Range['G'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat2').asDateTime);
       E.ActiveWorkBook.Sheets.Item[1].Range['H'+IntToStr(Nstroka)].Value:=tmprepdox.fieldbyname('ndfl').asFloat;
       if tmprepdox.fieldbyname('dat').asDateTime>EncodeDate(2000,1,1) then
         E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat').asDateTime);
       if tmprepdox.fieldbyname('dat2').asDateTime>EncodeDate(2000,1,1) then
            E.ActiveWorkBook.Sheets.Item[1].Range['J'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat2').asDateTime);
       if tmprepdox.fieldbyname('dat3').asDateTime>EncodeDate(2000,1,1) then
          E.ActiveWorkBook.Sheets.Item[1].Range['K'+IntToStr(Nstroka)].Value:=formatDateTime('dd.mm.yyyy',tmprepdox.fieldbyname('dat3').asDateTime);
      END;
      tmprepdox.next;
     end;

        try
         for km:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A6','K'+IntToStr(NSTROKA)].Borders[km].LineStyle:=1;
        except
        end;

        try
         E.ActiveWorkbook.Sheets.Item[1].Range['F4','F'+IntToStr(NSTROKA)].Font.Color:=clNavy;
         E.ActiveWorkbook.Sheets.Item[1].Range['F4','F'+IntToStr(NSTROKA)].Font.Italic:=True;
        except
        end;


      Nstroka:=Nstroka+1;
      E.ActiveWorkBook.Sheets.Item[1].Range['D'+IntToStr(Nstroka)].Value:=xs0;
      E.ActiveWorkBook.Sheets.Item[1].Range['H'+IntToStr(Nstroka)].Value:=xn0;

    E.Application.WindowState:=1;
    E.WindowState:=-4137 ;
    E:=UnAssigned;

end;

procedure TForm_58.JvXPButton12Click(Sender: TObject);
begin
 form3355:=tform3355.create(nil);
 form3355.ShowModal;
 form3355.free;
end;

procedure TForm_58.JvXPButton13Click(Sender: TObject);
var yy,oldRgod:Integer;
begin
 if MessageDlg('Выполнить предварительный расчет'+#13+
           'Для перехода остатков в случае отражения в Разделе 2 фактической даты выплаты необходимо выполнить и за предыдущий год',mtInformation,[mbYes,mbNo],0) = mrNo then exit;


 form206:=tform206.Create(nil);
   form206.Edit1.text:=Floattostr(RGod);
   form206.caption:='Предварительный расчет по году'    ;
   form206.ShowModal;
   form206.caption:='';
   try
    yy:=StrToInt(trim(form206.edit1.text));
   except
   end;
   form206.free;

   if yy=RGod then Button1.Click;

   if yy<>RGod then
    begin
     oldRgod:=Rgod;
     RGod:=yy;
      form1.ComboBox2.Text:=floattostr(RGod);
      form1.SmenaGod(false)  ;
     Button1.Click;
     Rgod:=oldrgod;
     form1.ComboBox2.Text:=floattostr(RGod);
     form1.SmenaGod(false)  ;
    end;


end;

procedure TForm_58.JvXPButton14Click(Sender: TObject);
begin
 form807:=Tform807.Create(nil);
 form807.ShowModal;
 form807.free;
 ComboBox1Change(nil);
end;

procedure TForm_58.JvXPButton15Click(Sender: TObject);
begin
 form807:=Tform807.Create(nil);
 form807.ShowModal;
 form807.free;
 ComboBox4Change(nil);
end;

function TForm_58.FGetUderjNdfl(tNls:Real;tDatEnd:TDAte):Real;      //налог удержан с начала года
var datstart:TDate;
    bRTF:Boolean;
    xndfl:Real;
    xNdfl0:Real;
begin
    xNdfl0:=0;
    datstart:=EncodeDAte(RGod,1,1) ;
           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.DatabaseName:=form1.DBDIR;
           datam.Query1.SQL.Add('select * from sdoxod where');
           datam.Query1.SQl.add('nls='+floattostr(tNls));
           datam.Query1.SQl.add('and dat>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
           datam.Query1.SQl.add('and dat<='+#39+formatdatetime('dd.mm.yyyy',tdatend)+#39);
           datam.query1.prepare;
           datam.query1.open;
           datam.Query1.first;
           while not datam.query1.eof do
            begin
              bRTF:=true;
              if datam.Query1.fieldbyname('kodnac').asfloat<>0 then
                begin
                 form1.NACISL.Locate('kod',datam.Query1.fieldbyname('kodnac').asInteger,[loCaseInsensitive]);
                 if form1.NACISLPN.Value=1 then bRTF:=False;
                end;

              if not FOktmo(datam.Query1.fieldbyname('mes').asInteger,datam.Query1.fieldbyname('god').asInteger,ComboBox4.Text) then bRtf:=false;

              if bRTF then xNdfl0:=xNdfl0+datam.Query1.Fieldbyname('nalog').asFloat;

             datam.query1.next;
            end;

           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.DatabaseName:=form1.DBDIR;
           datam.Query1.SQL.Add('select dayrab,datoklad,dayotr,oklad,daycas,snalog,wm,wg,nls from glnew ' );
           datam.Query1.SQl.add('where dayotr<>0');
           datam.Query1.SQl.add('and datoklad>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
           datam.Query1.SQl.add('and datoklad<='+#39+formatdatetime('dd.mm.yyyy',tdatend)+#39);
           datam.Query1.SQl.add('and nls='+floattostr(tNls));
           datam.Query1.Prepare;
           datam.Query1.Open;
           datam.Query1.First;
           while not datam.Query1.Eof do
            begin
               xNdfl:=datam.Query1.Fieldbyname('snalog').asFloat;
               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac=0 and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                   xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;

               if FOktmo(datam.Query1.fieldbyname('wm').asInteger,datam.Query1.fieldbyname('wg').asInteger,ComboBox4.Text) then xNdfl0:=xNdfl0+xndfl;

             datam.Query1.Next;
            end;

           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.SQL.Add('select n.koddox,n.rk, o.* from nacisl n, obrt1new o where o.kod=n.kod');
           datam.Query1.SQl.add('and o.datprov>='+#39+formatdatetime('dd.mm.yyyy',datstart)+#39);
           datam.Query1.SQl.add('and o.datprov<='+#39+formatdatetime('dd.mm.yyyy',tdatend)+#39);
           datam.Query1.SQl.add('and o.nls='+floattostr(tNls));
           datam.Query1.SQl.add('and n.pn<>1');
           datam.Query1.Prepare;
           datam.Query1.Open;
           datam.Query1.First;
           while not datam.Query1.Eof do
            begin
               xNdfl:=datam.Query1.Fieldbyname('snalog').asFloat;
               datam.qtmp.close;   //отнимаем частичный доход
               datam.qtmp.DatabaseName:=form1.DBDIR;
               datam.qtmp.sql.clear;
               datam.qtmp.sql.add('select * from sdoxod where kodnac='+floattostr(datam.Query1.FieldByName('kod').asFloat)+' and nls='+floattostr(datam.Query1.FieldByName('nls').asFloat));
               datam.qtmp.sql.add('and mes='+floattostr(datam.Query1.FieldByName('wm').asFloat));
               datam.qtmp.sql.add('and god='+floattostr(datam.Query1.FieldByName('wg').asFloat));
               datam.qtmp.prepare;
               datam.qtmp.open;
                while not datam.qtmp.eof do
                 begin
                  xndfl:=xndfl-datam.qtmp.fieldbyname('nalog').asFloat;
                  datam.qtmp.next;
                 end;
                datam.qtmp.close;

              if FOktmo(datam.Query1.fieldbyname('wm').asInteger,datam.Query1.fieldbyname('wg').asInteger,ComboBox4.Text) then xNdfl0:=xNdfl0+xndfl;

             datam.Query1.Next;
            end;

 FGetUderjNdfl:=xNdfl0;

end;


procedure TForm_58.JvXPButton16Click(Sender: TObject);
begin
 if MessageDlg('Сбросить все данные по выгрузке справок (дата, номер) за текущий год',mtWarning,[mbYes,mbNo],0) = mrNo then exit;

       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('DELETE from spr2006 where GOD='+FloatToStr(RGod));
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
       datam.Query1.Close;
       rxCalcEdit1.Value:=1;
       MessageDlg('Выполнено',mtInformation,[mbOk],0);
end;

procedure TForm_58.N20171Click(Sender: TObject);
var tNls:Real;
    i:integer;
    oldRx:Real;
    nSpr:Real;
begin
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select max(num) from spr2006 where GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       nSpr:=datam.Query1.Fields[0].asFloat;
       datam.Query1.Close;

 oldRx:=rxCalcEdit1.Value;
 if Length(Fac(Edit4.Text))<>4 then
   begin
    Edit4.SetFocus;
    exit;
   end;
 if rxCalcEdit1.Value=0 then
  begin
   MessageDlg('Введите номер справки',mtWarning,[mbOk],0);
   exit;
  end;

form1.kart.first;
 while not form1.kart.Eof do
  begin
   form1.kart.edit;
   form1.kart.fieldbyname('g').asString:='';
   form1.kart.post;
   form1.kart.next;
  end;
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 tNls:=form1.kartNLS.Value;
 form2.DBGrid1.Columns[0].Visible:=false;  //Check in DBGrid
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;

 if not form2.TOK then exit;

 PRNUMDAT:=False;

 if nSpr>0 then if MessageDlg('Обнаружена выгрузка справок в файл XML с присвоением номеров и дат'+#13+'Взять данные из выгрузки ?'+#13+
  '(иначе подставлять из параметров окна: Нач.№ '+floattostr(Trunc(rxCalcEdit1.Value))+', дата='+FormatDateTime('dd.mm.yyyy',DateEdit1.Date)+')',mtInformation,[mbYes,mbNo],0) = mrYes then PRNUMDAT:=True else PRNUMDAT:=False;

 form1.kart.first;
 i:=0;
 while not form1.kart.Eof do
  begin
   if form1.kart.FieldByName('G').asString='*' then
      begin
       i:=i+1;
       form_58.PSpr2016(form1.kartNls.Value);
       rxCalcEdit1.Value:=rxCAlcEdit1.Value+1;
      end;
   form1.kart.Next;
  end;

 if i=0 then
  begin
    form1.kart.locate('nls',tNls,[loCaseInsensitive]);
    form_58.PSpr2016(form1.kartnls.Value);
  end;

{
 Form2.ShowModal;
 if not form2.TOk then exit;
}

 // idNum:=idNum-1; {потому что внутри +1}

 if PRNUMDAT then rxCalcEdit1.Value:=oldRx;
end;

procedure TForm_58.N20171720181Click(Sender: TObject);
var tNls:Real;
    i:integer;
    oldRx:Real;
    nSpr:Real;
begin
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select max(num) from spr2006 where GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       nSpr:=datam.Query1.Fields[0].asFloat;
       datam.Query1.Close;

 oldRx:=rxCalcEdit1.Value;
 if Length(Fac(Edit4.Text))<>4 then
   begin
    Edit4.SetFocus;
    exit;
   end;
 if rxCalcEdit1.Value=0 then
  begin
   MessageDlg('Введите номер справки',mtWarning,[mbOk],0);
   exit;
  end;

form1.kart.first;
 while not form1.kart.Eof do
  begin
   form1.kart.edit;
   form1.kart.fieldbyname('g').asString:='';
   form1.kart.post;
   form1.kart.next;
  end;
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 tNls:=form1.kartNLS.Value;
 form2.DBGrid1.Columns[0].Visible:=false;  //Check in DBGrid
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;

 if not form2.TOK then exit;

 PRNUMDAT:=False;

 if nSpr>0 then if MessageDlg('Обнаружена выгрузка справок в файл XML с присвоением номеров и дат'+#13+'Взять данные из выгрузки ?'+#13+
  '(иначе подставлять из параметров окна: Нач.№ '+floattostr(Trunc(rxCalcEdit1.Value))+', дата='+FormatDateTime('dd.mm.yyyy',DateEdit1.Date)+')',mtInformation,[mbYes,mbNo],0) = mrYes then PRNUMDAT:=True else PRNUMDAT:=False;

 form1.kart.first;
 i:=0;
 while not form1.kart.Eof do
  begin
   if form1.kart.FieldByName('G').asString='*' then
      begin
       i:=i+1;
       form_58.PSpr2017(form1.kartNls.Value);
       rxCalcEdit1.Value:=rxCAlcEdit1.Value+1;
      end;
   form1.kart.Next;
  end;

 if i=0 then
  begin
    form1.kart.locate('nls',tNls,[loCaseInsensitive]);
    form_58.PSpr2017(form1.kartnls.Value);
  end;

{
 Form2.ShowModal;
 if not form2.TOk then exit;
}

 // idNum:=idNum-1; {потому что внутри +1}

 if PRNUMDAT then rxCalcEdit1.Value:=oldRx;
end;

procedure TForm_58.JvXPButton4Click(Sender: TObject);
begin

 if RGod>=2018 then
   if Messagedlg('Данная форма не действует с 2018 года'+#13+'Все равно продолжить ?',
          mtInformation,[mbYes,mbNo],0) = mrNo then EXIT;
 



    RPriznak:=trim(Edit9.TExt);


    N20171720181Click(nil);
{
   PopupMenu2.Popup(JvXPButton4.Left+JvXPButton4.Width+form_58.Left+Panel2.Left,
         JvXPButton4.Top+JvXPButton4.Height+form_58.Top+Panel2.Top);
}
end;

procedure TForm_58.JvXPButton17Click(Sender: TObject);
begin
 Form11133:=TForm11133.Create(nil);
 form11133.ShowModal;
 if form11133.TOk then
  begin
    RKod:=trim(form_58.reorgkod.Value);
    RInn:=trim(form_58.reorginn.Value);
    RKpp:=trim(form_58.reorgKpp.Value);
    Edit99.Text:='Код='+RKod;
    if RInn<>'' then Edit99.Text:=Edit99.Text+', ИНН='+RInn;
    if RKpp<>'' then Edit99.Text:=Edit99.Text+', КПП='+RKpp;
    Edit10.Text:=Edit99.TExt;
  end
   else
    begin
     Edit99.Text:='<нет>';
     RKod:='';  RInn:='';  RKpp:='';  RPriznak:='1';
    end;
 form11133.Free;

end;

procedure TForm_58.ComboBox3Change(Sender: TObject);
begin
 Showmessage('Выбрано:'+#13+Combobox3.Text+#13+Fkodndfl6());
end;

procedure TForm_58.JvXPButton18Click(Sender: TObject);
begin
 JvXPButton17Click(NIL);
end;

procedure TForm_58.JvXPButton19Click(Sender: TObject);
begin
   PopupMenu3.Popup(JvXPButton19.Left+JvXPButton19.Width+form_58.Left+Panel2.Left,JvXPButton19.Top+JvXPButton19.Height+form_58.Top+Panel2.Top);


end;

procedure TForm_58.N21711566021020181Click(Sender: TObject);
 var tNls:Real;
    i,j:integer;
    x:Real;
    oldRx:Real;
    nSpr:Real;
    F:TextFile;
    nErr:Integer;
    rtf:boolean;
begin

   form228:=Tform228.Create(nil);
           form228.Label1.Caption:='№ страницы первой справки (зависит от кол-ва стр. в 6-НДФЛ)';
           form228.RxCalcEdit1.Value:=NUMLIST;
           form228.ShowModal;
           if not form228.TOk then
             begin
              form228.free;
              exit;
             end;
           NUMLIST:=Trunc(INT(form228.rxCAlcEdit1.Value));
           form228.Free;

 if MessageDlg('Отметить все записи для формирования Справок ?'+#13+
      'Рекомендуется в случае, если необходимо сформировать по всем сотрудникам в отчетном году.'+#13+
       '(по cотрудникам, у которых не было дохода в течение отчетного года, формирование не произойдет)'+#13+
        'Иначе - выбор из списка вручную.',
         mtInformation,[mbYes,mbNo],0) = mrYes then rtf:=true else rtf:=false;

  RPriznak:=trim(Edit9.TExt);

       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('select max(num) from spr2006 where GOD='+FloatToStr(RGod));
       datam.Query1.SQL.Add('and stavka=13');
       datam.Query1.Prepare;
       datam.Query1.Open;
       nSpr:=datam.Query1.Fields[0].asFloat;
       datam.Query1.Close;

 oldRx:=rxCalcEdit1.Value;
 if Length(Fac(Edit4.Text))<>4 then
   begin
    Edit4.SetFocus;
    exit;
   end;
 if rxCalcEdit1.Value=0 then
  begin
   MessageDlg('Введите номер справки',mtWarning,[mbOk],0);
   exit;
  end;



form1.kart.first;
 while not form1.kart.Eof do
  begin
   form1.kart.edit;
   if rtf then form1.kart.fieldbyname('g').asString:='*' else form1.kart.fieldbyname('g').asString:='';
   form1.kart.post;
   form1.kart.next;
  end;

 form2.CheckBox1.Checked:=true; 
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 tNls:=form1.kartNLS.Value;
 form2.DBGrid1.Columns[0].Visible:=false;  //Check in DBGrid
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;

 if not form2.TOK then exit;


 tIsprDec2022:=false;
 if RGOD=2022 then if MessageDlg('Исключить из справок выплату з/п за вторую половину декабря 2022г, которая была проведена в январе 2023г (если такая есть)'+#13+
   'Предварительно должна быть проведена обработка (кнопка Декабрь 2022)',mtInformation,[mbYes,mbNo],0) = mrYes then tIsprDec2022:=true;
 

 PRNUMDAT:=False;

 if nSpr>0 then if MessageDlg('Обнаружена выгрузка справок в файл XML с присвоением номеров и дат'+#13+'Взять данные из выгрузки ?'+#13+
  '(иначе подставлять из параметров окна: Нач.№ '+floattostr(Trunc(rxCalcEdit1.Value))+', дата='+FormatDateTime('dd.mm.yyyy',DateEdit1.Date)+')',mtInformation,[mbYes,mbNo],0) = mrYes then PRNUMDAT:=True else PRNUMDAT:=False;

 form1.kart.first;


 AssignFile(F,form1.DBCurr+'\vygr.txt');
 Rewrite(F);
 Writeln(F,'Список сотрудников, по которым не сформированы справки');

 i:=0;  nErr:=0; //не выгружено
 while not form1.kart.Eof do
  begin
   if form1.kart.FieldByName('G').asString='*' then
      begin
       i:=i+1;
       ZapolnMas;
       x:=0;
       for j:=1 to 12 do x:=x+DPlus[j];
       if x>0 then
         begin
          form_58.PSpr2019(form1.kartNls.Value);
         // rxCalcEdit1.Value:=rxCAlcEdit1.Value+1;
         end
          else
         begin
          nerr:=nerr+1;
          WriteLn(F,form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value);
         end;
      end;
   form1.kart.Next;
  end;

 CloseFile(F);

 if nErr>0 then
  begin
   MessageDlg('По некоторым сотрудникам из выбранного списка не сформированы справки по причине отсутствия дохода в отчетном периоде'+#13+
     'Список сейчас будет открыт',mtInformation,[mbOk],0);
   ShellExecute(handle,'open',PChar(form1.DBCurr+'\vygr.txt'),nil,nil,SW_SHOW);
  end;

 if i=0 then
  begin
    form1.kart.locate('nls',tNls,[loCaseInsensitive]);
    form_58.PSpr2019(form1.kartnls.Value);
  end;

 NUMLIST:=4;

 if PRNUMDAT then rxCalcEdit1.Value:=oldRx;


end;

procedure TForm_58.N5711566021020181Click(Sender: TObject);
var tNls:Real;
    i:integer;
begin
  RPriznak:=trim(Edit9.TExt);


form1.kart.first;
 while not form1.kart.Eof do
  begin
   form1.kart.edit;
   form1.kart.fieldbyname('g').asString:='';
   form1.kart.post;
   form1.kart.next;
  end;
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 tNls:=form1.kartNLS.Value;
 form2.DBGrid1.Columns[0].Visible:=false;  //Check in DBGrid
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;

 if not form2.TOK then exit;

 tIsprDec2022:=false;
 if RGOD=2022 then if MessageDlg('Исключить из справок выплату з/п за вторую половину декабря 2022г, которая была проведена в январе 2023г (если такая есть)'+#13+
   'Предварительно должна быть проведена обработка (кнопка Декабрь 2022)',mtInformation,[mbYes,mbNo],0) = mrYes then tIsprDec2022:=true;


 form1.kart.first;
 i:=0;
 while not form1.kart.Eof do
  begin
   if form1.kart.FieldByName('G').asString='*' then
      begin
       i:=i+1;
       form_58.PSpr20192(form1.kartNls.Value,true);

      end;
   form1.kart.Next;
  end;

 if i=0 then
  begin
    form1.kart.locate('nls',tNls,[loCaseInsensitive]);
    form_58.PSpr20192(form1.kartnls.Value,true);
  end;

end;

procedure TForm_58.ProcIsprOKTMO;
var i,j:Integer;
    x:Real;
begin
 for i:=12 downto 2 do DNal[i]:=DNal[i]-DNal[i-1];   //база с начала года
          for i:=1 to 12 do
           begin
            if FGEtOktmo(i,RGod)<>trim(combobox1.text) then
              begin
               for j:=1 to 10 do qqMT[j,i]:=0;
               for j:=1 to 10 do qqDV[j,i]:=0;
               DNal[i]:=0;
               DIsc[i]:=0;
               DDoxod9[i]:=0;
               DPn9[i]:=0;
               DImVyc[i]:=0;
               for j:=1 to 6 do DST20[j,i]:=0;
              end;
            end;
           for j:=1 to 6 do
            begin
             x:=0;
             for i:=1 to 12 do  x:=x+DST20[j,i];
             DST20[j,13]:=x;
            end;
            for j:=1 to 10 do
            begin
             x:=0;                ;
             for i:=1 to 12 do  x:=x+qqMT[j,i];
             qqMT[j,13]:=x;
            end;
             x:=0;
             for i:=1 to 12 do  x:=x+DNal[i];
             DNal[13]:=x;
             x:=0;
             for i:=1 to 12 do x:=x+DImVyc[i];
             DNal[13]:=DNal[13]-x; // не учитывается им. вычет ???
             x:=0;
             for i:=1 to 12 do  x:=x+DISC[i];
             DISC[13]:=x;
end;


procedure TForm_58.PSpr20192(wnls:Real;tFakt:boolean);
var fNameXLS:String;
    E:OleVAriant;
    xRegion:Real;
    dd,mm,yy:Word;
    s1:String;
    nd,nc,i,j,k:Integer;
    x:Real;
    soktmo:String;
    sKpp:String;
    nSpr:Real;
    nczap:integer;
    nDatSpr:TDate;
    TOK:Real;
    RtfI:boolean;
begin

    fNameXls:=GetNameXlsn('2НДФЛ','r')+'.xls';
    if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\2NDFL20192.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
     begin
      MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
      exit;
     end;
    E:=CreateOleObject('Excel.Application');
    E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);
    E.Visible:=True;
    E.Application.WindowState:=2;

    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;

     E.ActiveWorkBook.Sheets.Item[1].Range['AH5'].Value:=RGod;




     E.ActiveWorkBook.Sheets.Item[1].Range['BO10'].Value:=form1.config2INN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['P11'].Value:=form1.config2NAME.Value;
     datam.kart2.Locate('nls',form1.kartnls.value,[locaseinsensitive]);

     E.ActiveWorkBook.Sheets.Item[1].Range['AK12'].Value:=form_58.RKOD;
     E.ActiveWorkBook.Sheets.Item[1].Range['AK13'].Value:=form_58.RINN;
     E.ActiveWorkBook.Sheets.Item[1].Range['BF13'].Value:=form_58.RKPP;


     if Trim(datam.kart2OKTMO.Value)='' then soktmo:=form1.config2OKTMO.Value else
                                                                soktmo:=datam.kart2OKTMO.Value;

     //****
     RtfI:=true;
     TOK:=form_58.ZapolnDOK(form1.kartNls.Value);   //изменение ОКТМО
       if TOK=1 then
         begin
          s:=FGetOktmo(1,RGod);
          for i:=2 to 12 do if FGEtOktmo(i,RGod)<>s then RtfI:=false;
          if RtfI then soktmo:=s else
              begin
               soktmo:=trim(combobox1.text);
               MessageDlg(form1.kartfam.value+' обнаружено изменение ОКТМО в течение года'+#13+
                  'Справка формируется по ОКТМО '+soktmo+#13+
                    'Для формирования по другому ОКТМО выберите значение из списка КОД ОКТМО и повторите',mtInformation,[mbOk],0);
              end;   //октмо на начало года, изменений в течение года не было
         end;





     E.ActiveWorkBook.Sheets.Item[1].Range['N10'].Value:=soktmo;

     if oktmo.Locate('oktmo',soktmo,[locaseinsensitive]) then
      begin
       sKpp:=oktmoKpp.Value;
      end
       else
        skpp:=form1.config2KPP.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['CN10'].Value:=sKPP;


        E.ActiveWorkBook.Sheets.Item[1].Range['AP10'].Value:=form1.config2TEL.Value;

     E.ActiveWorkBook.Sheets.Item[1].Range['AB17'].Value:=form1.kartINN.asString;
     E.ActiveWorkBook.Sheets.Item[1].Range['J18'].Value:=form1.kartFAM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['AV18'].Value:=form1.kartIM.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CP18'].Value:=form1.kartOT.Value;
     E.ActiveWorkBook.Sheets.Item[1].Range['CJ19'].Value:=form1.kartSTRANA.Value;

     datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
     if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=7) then
                   E.ActiveWorkBook.Sheets.Item[1].Range['Y19'].Value:=datam.kart2STATUS2.AsString
                         else E.ActiveWorkBook.Sheets.Item[1].Range['Y19'].Value:='1';


        nDatSpr:=DAteEdit1.DAte;
        DecodeDAte(nDatSpr,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BJ5'].Value:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ5'].Value:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['BE5'].Value:=s1;



     if form1.kartBIRTHDAY.Value>=Encodedate(1920,1,1) then
       begin
        DecodeDAte(form1.kartBirthday.Value,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['BE19'].Value:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['AU19'].Value:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm);
        E.ActiveWorkBook.Sheets.Item[1].Range['AZ19'].Value:=s1;
       end;

       E.ActiveWorkBook.Sheets.Item[1].Range['AO20'].Value:=form1.kartKODDOC.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['BU20'].Value:=form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value;


       E.ActiveWorkBook.Sheets.Item[1].Range['A63'].Value:=podpndflFIO.Value;



       if form1.kartSTATUS.Value='2' then E.ActiveWorkBook.Sheets.Item[1].Range['AI22'].Value:=30 else E.ActiveWorkBook.Sheets.Item[1].Range['AI22'].Value:=13;

       ObrabObrtNalKart;

       //*****обработка двойного октмо в течение года

       if not RtfI then //было изменение в течение года
         begin
          ProcIsprOKTMO;
         end;

       if tIsprDec2022 then Ispr2NDFLDec2022;


       //



  k:=0;
  for j:=1 to 10 do
    begin
     if qqD[j]<>'' then k:=j;  //кол-во кодов для дивидендов 13% добавить чтобы
    end;
  if RGod>=2015 then
   begin
    if DKOD9[1]<>'' then
     begin
      qqD[k+1]:=DKOD9[1];
      for i:=1 to 12 do qqMT[k+1,i]:=DDOX9[1,i];
      for i:=1 to 12 do qqMT[k+1,13]:=qqMT[k+1,13]+DDOX9[1,i];
     end;
   end;


   if (RGod>=2023) and (tFakt) then NewSpr2023; //перераспределяем коды доходов, вычетов с 2023г


     nc:=1;nd:=0;


     nczap:=0;
     for i:=1 to 12 do for j:=1 to 10 do if (qqD[j]<>'') and (qqMT[j,i]<>0) then  nczap:=nczap+1;
     nczap:=TRUNC(nczap/2)+2;

     for i:=1 to 12 do
      begin
       for j:=1 to 10 do
        begin
          if (qqD[j]<>'') and (qqMT[j,i]<>0) then
           begin
            nd:=nd+1;
            if nd=nczap then
             begin
              nd:=1; nc:=nc+1;
             end;
            if nc=1 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(nd+24)].value:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['I'+IntToStr(nd+24)].value:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['Q'+IntToStr(nd+24)].value:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AF'+IntToStr(nd+24)].value:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['AN'+IntToStr(nd+24)].value:=qqDV[j,i];
             end;
            if nc=2 then
             begin
               E.ActiveWorkBook.Sheets.Item[1].Range['BG'+IntToStr(nd+24)].value:=i;
               E.ActiveWorkBook.Sheets.Item[1].Range['BO'+IntToStr(nd+24)].value:=qqD[j];
               E.ActiveWorkBook.Sheets.Item[1].Range['BW'+IntToStr(nd+24)].value:=qqMT[j,i];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CL'+IntToStr(nd+24)].value:=qqV[j];
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then E.ActiveWorkBook.Sheets.Item[1].Range['CT'+IntToStr(nd+24)].value:=qqDV[j,i];
             end;
           end;
        end;
       end;

      for j:=1 to 6 do
      begin
        if (DS20[j]<>'') and (DST20[j,13]<>0) then
         begin
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['A47'].Value:=DS20[j];
              if j=1 then E.ActiveWorkBook.Sheets.Item[1].Range['J47'].Value:=DST20[j,13];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AB47'].Value:=DS20[j];
              if j=2 then E.ActiveWorkBook.Sheets.Item[1].Range['AK47'].Value:=DST20[j,13];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BD47'].Value:=DS20[j];
              if j=3 then E.ActiveWorkBook.Sheets.Item[1].Range['BM47'].Value:=DST20[j,13];

              if j=4 then E.ActiveWorkBook.Sheets.Item[1].Range['A48'].Value:=DS20[j];
              if j=4 then E.ActiveWorkBook.Sheets.Item[1].Range['J48'].Value:=DST20[j,13];
              if j=5 then E.ActiveWorkBook.Sheets.Item[1].Range['AB48'].Value:=DS20[j];
              if j=5 then E.ActiveWorkBook.Sheets.Item[1].Range['AK48'].Value:=DST20[j,13];
              if j=6 then E.ActiveWorkBook.Sheets.Item[1].Range['BD48'].Value:=DS20[j];
              if j=6 then E.ActiveWorkBook.Sheets.Item[1].Range['BM48'].Value:=DST20[j,13];

         end;
      end;

    if form1.kartIMVYC_SUMM.Value<>0 then
     begin
      // E.ActiveWorkBook.Sheets.Item[1].Range['BR51'].Value:=form1.kartImVyc_Num.Value;

     //  E.ActiveWorkBook.Sheets.Item[1].Range['DE51'].Value:=form1.kartIMVYC_gni.Value;

    {
      DecodeDAte(form1.kartImVyc_Dat.Value,yy,mm,dd);
      if dd<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CD51'].Value:='0'+floattostr(dd) else E.ActiveWorkBook.Sheets.Item[1].Range['CD51'].Value:=floattostr(dd);
      if mm<10 then E.ActiveWorkBook.Sheets.Item[1].Range['CI51'].Value:='0'+floattostr(mm) else E.ActiveWorkBook.Sheets.Item[1].Range['CI51'].Value:=floattostr(mm);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN51'].Value:=yy;
   }
      x:=0;
      for j:=1 to 12 do x:=x+DImVyc[j];
      if x<>0 then E.ActiveWorkBook.Sheets.Item[1].Range['CP47'].Value:=x;
      if x<>0 then E.ActiveWorkBook.Sheets.Item[1].Range['CG47'].Value:=form1.kartIMVYC_KOD.Value;

     end;

    x:=0;
    for j:=1 to 10 do x:=x+qqMT[j,13];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF53'].Value:=DRound(x,2);
    x:=DNal[13];
    for j:=1 to 12 do x:=x+DDoxod9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF54'].Value:=DRound(x,2);

    x:=DIsc[13];
    for j:=1 to 13 do x:=x+DPn9[j];
    E.ActiveWorkBook.Sheets.Item[1].Range['AF55'].Value:=DRound(x,2);
    E.ActiveWorkBook.Sheets.Item[1].Range['AF56'].Value:=DRound(x,2);

    x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'')+FUplataNdfl(form1.kartNLS.Value,RGod,9,'');
    if not RtfI then x:=FUplataNdfl(form1.kartNLS.Value,RGod,13,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,30,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,9,soktmo)   ;
    E.ActiveWorkBook.Sheets.Item[1].Range['CN55'].Value:=DRound(x,2);


    datam.query1.close;
    datam.query1.sql.clear;
    datam.query1.sql.add('select * from uplatandfl where nls='+floattostr(form1.kartnls.value));
    datam.query1.sql.add('and type=2 and summa>0 and god='+floattostr(RGod));
    datam.query1.Prepare;
    datam.Query1.open;
    datam.query1.first;
     if datam.query1.RecordCount=1 then //аванс
      begin
        E.ActiveWorkBook.Sheets.Item[1].Range['CN53'].Value:=DRound(datam.query1.FieldByName('summa').asFloat,2);
       {
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59'].Value:=datam.query1.FieldByName('numuved').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59'].Value:=datam.query1.FieldByName('ifns').asString;
        DecodeDate(datam.query1.FieldByName('datuved').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59'].Value:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59'].Value:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59'].Value:=s1;
       }
      end;
    datam.query1.close;

    if ndflr5.Locate('nls;god',VarArrayOf([form1.kartNls.Value,RGod]),[loCaseInsensitive]) then
     begin
      if ndflr5PR2.Value=1 then E.ActiveWorkBook.Sheets.Item[1].Range['AF56'].Value:=DRound(ndflr5.FieldByName('sud').asFloat,2);
      if ndflr5PR1.Value=1 then E.ActiveWorkBook.Sheets.Item[1].Range['AF55'].Value:=DRound(ndflr5.FieldByName('sisc').asFloat,2);


      E.ActiveWorkBook.Sheets.Item[1].Range['CN54'].Value:=DRound(ndflr5.FieldByName('sprib').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CN56'].Value:=DRound(ndflr5.FieldByName('suderj').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CC61'].Value:=DRound(ndflr5.FieldByName('DOXNEUD').asFloat,2);
      E.ActiveWorkBook.Sheets.Item[1].Range['CC62'].Value:=DRound(ndflr5.FieldByName('snuderj').asFloat,2);

      E.ActiveWorkBook.Sheets.Item[1].Range['CN53'].Value:=DRound(ndflr5.FieldByName('sfix').asFloat,2);
     {
      if ndflr5.FieldByName('sfix').asFloat<>0 then
       begin
        E.ActiveWorkBook.Sheets.Item[1].Range['BO59'].Value:=ndflr5.FieldByName('num').asString;
        E.ActiveWorkBook.Sheets.Item[1].Range['DE59'].Value:=ndflr5.FieldByName('ifns').asString;
        DecodeDate(ndflr5.FieldByName('dat').asdateTime,yy,mm,dd);
        E.ActiveWorkBook.Sheets.Item[1].Range['CN59'].Value:=yy;
        if dd<10 then s1:='0'+floattostr(dd) else s1:=floattostr(dd) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CD59'].Value:=s1;
        if mm<10 then s1:='0'+floattostr(mm) else s1:=floattostr(mm) ;
        E.ActiveWorkBook.Sheets.Item[1].Range['CI59'].Value:=s1;
       end;
      }
     end;

   try
    nczap:=nczap+24;
    if nczap<42 then
       E.ActiveWorkBook.Sheets.Item[1].Rows[Floattostr(nczap)+':43'].Hidden:=True;
    // E.ActiveWorkBook.Sheets.Item[1].Range['A'+Inttostr(nczap),'A'+IntToStr(43)].EntireRow.Delete(EmptyParam);
   except
   end;

   E.WindowState:=-4137 ;

   try
    E.DisplayAlerts:=false;
    E.WorkBooks[1].Save;
   except
   end;

   E:=UnAssigned;
   

   x:=0;                         //35% отдельная справка
   for i:=1 to 12 do x:=x+DDoxod35[i];
   if x<>0 then
     begin
      rxCalcEdit1.Value:=rxCalcEdit1.Value+1;
      PSpr2016_st35(wnls);
     end;

end;



procedure TForm_58.RxLabel10Click(Sender: TObject);
var dd,mm,yy:Word;
begin

 DecodeDAte(DAteEdit2.Date,yy,mm,dd);
 ZapolnDOK(1);

 if FOktmo(mm,yy,ComboBox4.Text) then
  Showmessage('Ok') else Showmessage('No');

end;


procedure Tform_58.ZapolnSpr2021(FType:Integer;sKpp:String);
var
    i,j,k,nd,nst:Integer;
  
    y,z,supl:Real;
    fileName:String;
    FST:array[1..4] of Integer;
    jtf1,jtf2,jtf4,jtf0:Boolean;
    jn,jn0:integer;
    jtf90,jtf91:boolean;
    xDoxod9,x:Real;
    xRegion:Real;
    xSumVic:Real;
    xFam,xIm,xOt:String;
    xKartStatus:String  ;
    xisc,xuderj:Real;
    wStavka:Real;
    fNameXLS:String;
    nzapSt13,nLevzapSt13:integer;
    rtfdn:Boolean;
    xtt,xtt2:Real;
    tisc,tud:Real;
    RtfI:Boolean;
    TOK:Real;
    soktmo:String;
begin

   FST[1]:=13;
   FST[2]:=9;
   FST[3]:=30;
   FST[4]:=35;




    xRegion:=0;

    //удаляем номера справок перед формированием
    if (FTYPE=2) and (form624.CheckBox4.Checked) then
     begin
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('delete from spr2006 where nls='+FloatToStr(form1.kartNLS.Value));
       datam.Query1.SQL.Add('and god='+FloatToStr(Rgod));
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
       datam.Query1.Close;
     end;

    datam.QKladr.Close;
    datam.Qkladr.DatabaseName:=form52.DBKLADR2;
    datam.Qkladr.SQL.Clear;
    datam.Qkladr.SQL.Add('select region from region where name LIKE '+#39+Trim(AnsiUpperCase(form1.kartREGION.Value))+'%'+#39) ;
    datam.Qkladr.Prepare;
    datam.Qkladr.Open;
    if datam.Qkladr.RecordCount<>1 then
     begin
      // MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.VAlue+' АДРЕС неверно указан Регион.',mtWarning,[mbOk],0);
      xRegion:=0;
     end
      else xRegion:=datam.QKladr.Fields[0].asFloat;


 for i:=1 to 2 do
 begin
  DKOD35[i]:=''; DKODVIC35[i]:='';
   for k:=1 to 13 do
      begin
       DDOX35[i,k]:=0;
       DVIC35[i,k]:=0;
      end;
 end;


 //****
     RtfI:=true;
     TOK:=form_58.ZapolnDOK(form1.kartNls.Value);   //изменение ОКТМО
       if TOK=1 then
         begin
          s:=form_58.FGetOktmo(1,RGod);
          for i:=2 to 12 do if form_58.FGEtOktmo(i,RGod)<>s then RtfI:=false;
          if RtfI then soktmo:=s else
              begin
               soktmo:=trim(form_58.combobox1.text);
               MessageDlg(form1.kartfam.value+' обнаружено изменение ОКТМО в течение года'+#13+
                  'Справка формируется по ОКТМО '+soktmo+#13+
                    'Для формирования по другому ОКТМО выберите значение из списка КОД ОКТМО и повторите',mtInformation,[mbOk],0);
              end;   //октмо на начало года, изменений в течение года не было
         end;


 ObrabObrtNalKart;

 if (not RtfI) or ((RtfI) and (TOK=1) and (soktmo<>trim(form_58.combobox1.text))) then
   begin
    form_58.ProcIsprOktmo;
   end;

 // Showmessage(DS[1]+' '+DS[2]+' '+DS[3]);


 if form624.PDEC2022_v=1 then Ispr2NDFLDec2022;
                                                                                               


  k:=0;
  for j:=1 to 10 do
    begin
     if qqD[j]<>'' then k:=j;  //кол-во кодов для дивидендов 13% добавить чтобы
    end;
  if RGod>=2015 then
   begin
    if DKOD9[1]<>'' then
     begin
      qqD[k+1]:=DKOD9[1];
      for i:=1 to 12 do qqMT[k+1,i]:=DDOX9[1,i];
      for i:=1 to 12 do qqMT[k+1,13]:=qqMT[k+1,13]+DDOX9[1,i];
     end;
   end;


  if RGod>=2023 then NewSpr2023; //перераспределяем коды доходов, вычетов с 2023г ;


  s:='';
  for j:=1 to 10 do
   begin
    s:=s+'j='+floattostr(j)+' '+qqD[j]+#13;
    for i:=1 to 12 do if qqMT[j,i]<>0 then s:=s+'i='+floattostr(i)+' сумма='+floattostr(qqMT[j,i])+#13;
   end;
//  ShowMessage(s);



    //jtf1 - есть 13% jtf2 - 9% jtf4 - 35%
    // jtf0 - вообще ечть любой из доходов
    jtf1:=False; jtf2:=False; jtf4:=false;
    for j:=1 to 10 do if qqMT[j,13]<>0 then jtf1:=TRUE;
    for i:=1 to 12 do if DDOXOD9[i]<>0 then jtf2:=True;
    for j:=1 to 2 do if DDOX35[j,13]<>0 then jtf4:=TRUE;

    if RGod>=2015 then  jtf2:=False;  //дивиденды 2015 13%




 jtf0:=false;
 if (jtf1) or (jtf2) or (jtf4) then jtf0:=true;


 jn0:=0;
 if (jtf1) then jn0:=jn0+1;  //всего сколько справок т.е. доходов
 if (jtf2) then jn0:=jn0+1;
 if (jtf4) then jn0:=jn0+1;


 jtf90:=true; //признак шапка
 jtf91:=true; //признак дно

 jn:=0; //признак что первый проход заголовок справки
 FOR nst:=1 to 1 do
 BEGIN
  if (jtf1) and (nst=1) then jn:=jn+1;    // jn - внутренний счетчик
  if (jtf2) and (nst=2) then jn:=jn+1;
  if (jtf4) and (nst=4) then jn:=jn+1;




  IF JTF0 THEN
  BEGIN


    //сохранем  номера справок
    if (FTYPE=2) and (form624.CheckBox4.Checked) then
     begin
       if nst=1 then wStavka:=13;
       if nst=2 then wStavka:=9;
       if nst=4 then wStavka:=35;
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('INSERT INTO spr2006(nls,num,dat,god,stavka) values('+FloattoStr(form1.kartNLS.Value)+','+
               FloatToStr(form1.idNum+1)+','+#39+FormatdateTime('dd.mm.yyyy',form_58.DateEdit2.Date)+#39+','+FloatToStr(RGod)+','+FloatToStr(wStavka)+')');
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
       datam.Query1.Close;
     end;



  if (FType=2) and (jtf0) and (jn=1) and (jtf90) then    //jn1=1 - певрый проход всего 1..4 выводим шапку при первом проходе далее доходы на каждый проход
   begin
    jtf90:=false; //сброс вывода шапка

    form1.idNum:=form1.idNum+1;   //внутренний номер справки


    form1.nnspr1:=form1.nnspr1+1;


    if (RGod>=2018) then
     begin

          form1.nnspr2:=form1.nnspr2+1;
          itkolvo:=itkolvo+1;
          EREESTR.ActiveWorkBook.Sheets.Item[2].Range['B'+IntToStr(form1.nnspr2+19)].Value:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;
          EREESTR.ActiveWorkBook.Sheets.Item[2].Range['A'+IntToStr(form1.nnspr2+19)].Value:=form1.idNum;
          if form1.kartSTATUS.Value='2' then EREESTR.ActiveWorkBook.Sheets.Item[2].Range['C'+IntToStr(form1.nnspr2+19)].Value:=30 else
                                         EREESTR.ActiveWorkBook.Sheets.Item[2].Range['C'+IntToStr(form1.nnspr2+19)].Value:=13;

          WriteLn(FXML,'<СправДох НомСпр="'+IntToStr(form1.idNum)+'" НомКорр="'+trim(form_58.edit3.text)+'">')   ;

          s:='<ПолучДох ';
          if form1.kartINN.Value<>'' then s:=s+' ИННФЛ="'+form1.kartINN.Value+'"';
          if form1.kartSTATUS.Value='2' then xKartStatus:='2' else xKartStatus:='1';

          datam.kart2.locate('nls',form1.kartnls.value,[loCaseInsensitive]);
          if (datam.kart2STATUS2.Value>=1) and (datam.kart2STATUS2.Value<=7) then
                                               xKartStatus:=datam.kart2STATUS2.AsString
                                                                          else xKartStatus:='1';

          s:=s+' Статус="'+xKartStatus+'" ДатаРожд="'+FormatDateTime('dd.mm.yyyy',form1.kartBirthday.Value)+'" Гражд="'+FGrajd(form1.kartSTRANA.Value)+'">';
          WriteLn(FXML,s);

          if trim(form1.kartOT.Value)<>'' then
               WriteLn(FXML,' <ФИО Фамилия="'+form1.kartFAM.Value+'" Имя="'+form1.kartIm.Value+'" Отчество="'+form1.kartOT.Value+'"/>')
                 else
                   WriteLn(FXML,' <ФИО Фамилия="'+form1.kartFAM.Value+'" Имя="'+form1.kartIm.Value+'"/>') ;

          WriteLn(FXML,' <УдЛичнФЛ КодУдЛичн="'+form1.kartKODDOC.Value+'" СерНомДок="'+form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartpass.Value+'"/>');
          WriteLn(FXML,'</ПолучДох> ');


     end;










    {начало 13%}
   IF (nst=1) and (jtf1) THEN
    BEGIN


    k:=0;
    nd:=0;

    if RGod>=2018 then
     begin

       if form1.kartSTATUS.Value='2' then Writeln(FXML,'          <СведДох Ставка="30">') else Writeln(FXML,'          <СведДох Ставка="13" КБК="'+Form_58.sKbk+'">');

              x:=0;
              for j:=1 to 10 do x:=x+qqMT[j,13];

            xtt:=DNal[13];
            for i:=1 to 12 do xtt:=xtt+DDoxod9[i];

            xtt2:=DIsc[13];
          // for i:=1 to 12 do xtt2:=xtt2+DPn9[i];

          supl:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'');
          if not RtfI then supl:=FUplataNdfl(form1.kartNLS.Value,RGod,13,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,30,soktmo);

           if not form_58.ndflr5.Locate('nls;god',VarArrayOf([form1.kartnls.Value,RGod]),[loCaseInsensitive]) then
              begin
                  datam.query1.close;
                  datam.query1.sql.clear;
                  datam.query1.sql.add('select * from uplatandfl where nls='+floattostr(form1.kartnls.value));
                  datam.query1.sql.add('and type=2 and summa>0 and god='+floattostr(RGod));
                  datam.query1.Prepare;
                  datam.Query1.open;
                  datam.query1.first;

                  if datam.query1.RecordCount<>1 then
                   begin
                    WriteLn(FXML,'<СумИтНалПер СумДохОбщ="'+FStr(x,2)+'" НалБаза="'+FStr(xtt,2)+'" НалИсчисл="'+FStr(xtt2,0)+
                       '" АвансПлатФикс="'+FStr(0,0)+'" '+
                     ' НалУдерж="'+FStr(xtt2,0)+'" НалПеречисл="'+FStr(supl,0)+'" НалУдержЛиш="0" СумНалПрибЗач="0"/>');
                   end
                      else//аванс
                   begin
                     Write(FXML,'<СумИтНалПер СумДохОбщ="'+FStr(x,2)+'" НалБаза="'+FStr(xtt,2)+'" НалИсчисл="'+FStr(xtt2,0)+
                     '" АвансПлатФикс="'+FStr(datam.query1.FieldByName('summa').asFloat,0)+'" '+
                       ' НалУдерж="'+FStr(xtt2,0)+'" НалПеречисл="'+FStr(supl,0)+'" НалУдержЛиш="'+FStr(0,0)+'" СумНалПрибЗач="0" >');

                     Write(FXML,' <УведФиксПлат НомерУвед="'+trim(datam.query1.FieldByName('numuved').asString)+
                             '" ДатаУвед="'+formatdatetime('dd.mm.yyyy',datam.query1.FieldByName('datuved').asdateTime)
                      +'" ИФНСУвед="'+trim(datam.query1.FieldByName('ifns').asString)+'"/>');

                      WriteLn(FXML,'</СумИтНалПер>');
                    end;

                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['D'+IntToStr(form1.nnspr2+19)].Value:=x;    //начисл
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['E'+IntToStr(form1.nnspr2+19)].Value:=xtt2; //исч
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['F'+IntToStr(form1.nnspr2+19)].Value:=xtt2; //удер
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['G'+IntToStr(form1.nnspr2+19)].Value:=supl; //перечис
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['H'+IntToStr(form1.nnspr2+19)].Value:=0; //излишне удержано
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['I'+IntToStr(form1.nnspr2+19)].Value:=0; //не удержано
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['J'+IntToStr(form1.nnspr2+19)].Value:=datam.query1.FieldByName('summa').asFloat; //фикс
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['K'+IntToStr(form1.nnspr2+19)].Value:=0; //приб
                     itdoxod:=itdoxod+x;
                     itisc:=itisc+xtt2;
                     ituderj:=ituderj+xtt2;
                     itupl:=itupl+supl;
                     itfix:=itfix+datam.query1.FieldByName('summa').asFloat;
                     
                  datam.query1.close;
                 end
               else
              begin //раздел 5 справки


              if form_58.ndflr5PR1.Value=1 then xisc:=form_58.ndflr5sisc.value else xisc:=xtt2;
              if form_58.ndflr5PR2.Value=1 then xuderj:=form_58.ndflr5sud.value else xuderj:=xtt2;
              
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['D'+IntToStr(form1.nnspr2+19)].Value:=x;    //начисл
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['E'+IntToStr(form1.nnspr2+19)].Value:=xisc; //исч
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['F'+IntToStr(form1.nnspr2+19)].Value:=xuderj; //удер
                     EREESTR.ActiveWorkBook.Sheets.Item[2].Range['G'+IntToStr(form1.nnspr2+19)].Value:=supl; //перечис

                      EREESTR.ActiveWorkBook.Sheets.Item[2].Range['H'+IntToStr(form1.nnspr2+19)].Value:=form_58.ndflr5suderj.value;
                      EREESTR.ActiveWorkBook.Sheets.Item[2].Range['I'+IntToStr(form1.nnspr2+19)].Value:=form_58.ndflr5snuderj.Value;;
                      EREESTR.ActiveWorkBook.Sheets.Item[2].Range['J'+IntToStr(form1.nnspr2+19)].Value:=form_58.ndflr5sfix.value;
                      EREESTR.ActiveWorkBook.Sheets.Item[2].Range['K'+IntToStr(form1.nnspr2+19)].Value:=form_58.ndflr5SPRIB.Value;


                     itdoxod:=itdoxod+x;
                     itisc:=itisc+xisc;
                     ituderj:=ituderj+xuderj;
                     itupl:=itupl+supl;
                     itfix:=itfix+form_58.ndflr5sfix.value;
                     itlis:=itlis+form_58.ndflr5suderj.value;
                     itprib:=itprib+form_58.ndflr5SPRIB.Value;
                     itneud:=itneud+form_58.ndflr5snuderj.Value;

                  WriteLn(FXML,'<СумИтНалПер СумДохОбщ="'+FStr(x,2)+'" НалБаза="'+FStr(xtt,2)+'" НалИсчисл="'+FStr(xisc,0)+
                  '" АвансПлатФикс="'+FStr(form_58.ndflr5sfix.value,0)+'" '+
                     ' НалУдерж="'+FStr(xuderj,0)+'" НалПеречисл="'+FStr(supl,0)+'" НалУдержЛиш="'+FStr(form_58.ndflr5suderj.value,0)
                      +'" СумНалПрибЗач="'+FStr(form_58.ndflr5SPRIB.Value,0)+'" />');


                
              end;

          //****
            xSumVic:=0;
             for j:=1 to 6 do
              begin
               xSumVic:=xSumVic+DST20[j,13];
              end;

           // if (Ftype=2) and (DRound(xSumVic+form1.kartIMVYC_SUMM.Value,2)<>0)  then WriteLn(FXML,'   <НалВычССИ>')  ;

            WriteLn(FXML,'   <НалВычССИ>')  ;

             for j:=1 to 6 do
              begin
               if (DS20[j]<>'') and (DST20[j,13]<>0) then
                 begin
                   if (FType=2) then WriteLn(FXML,' <ПредВычССИ КодВычет="'+DS20[j]+'" СумВычет="'+FStr(DST20[j,13],2)+'"/>');
                 end;
              end;
             x:=0;
             for j:=1 to 12 do x:=x+DImVyc[j];
             if (Ftype=2) and (x<>0)  then
               begin
                WriteLn(FXML,' <ПредВычССИ КодВычет="'+form1.kartIMVYC_KOD.asString+'" СумВычет="'+FStr(x,2)+'"/>')  ;
                WriteLn(FXML,'<УведВыч КодВидУвед="1" НомерУвед="'+Trim(form1.kartImVyc_Num.Value)+'" ДатаУвед="'
                   +FormatDateTime('dd.mm.yyyy',form1.kartImVyc_Dat.Value)+'" НОУвед="'+trim(form1.kartImVyc_Gni.Value)+'"/>');
              end;

              //**увед фикс платежи
              if form_58.ndflr5.Locate('nls;god',VarArrayOf([form1.kartnls.Value,RGod]),[loCaseInsensitive]) then
               begin
                 if trim(form_58.ndflr5num.Value)<>'' then
                      WriteLn(FXML,'<УведВыч КодВидУвед="3" НомерУвед="'+trim(form_58.ndflr5num.Value)+
                         '" ДатаУвед="'+formatdatetime('dd.mm.yyyy',form_58.ndflr5dat.Value)+'" НОУвед="'+
                             trim(form_58.ndflr5ifns.Value)+'"/>' );
               end;

              // if (Ftype=2) and (DRound(xSumVic+form1.kartIMVYC_SUMM.Value,2)<>0) then WriteLn(FXML,'   </НалВычССИ>')  ;
              WriteLn(FXML,'   </НалВычССИ>')  ;

              //***
              if form_58.ndflr5.Locate('nls;god',VarArrayOf([form1.kartnls.Value,RGod]),[loCaseInsensitive]) then
               begin
                 WriteLn(FXML,'<СумДохНеУд СумДохНеУдерж="'+FStr(form_58.ndflr5doxneud.value,2)+'" СумНеУдНал="'+FStr(form_58.ndflr5snuderj.Value,0)+'"/>');
               end;

        Writeln(FXML,'            <ДохВыч>');

      end;



  


    rtfdn:=false;

    for i:=1 to 12 do
     begin
      for j:=1 to 10 do
       begin
         s:='';
         if (qqD[j]<>'') and (qqMT[j,i]<>0) then
          begin
        //   k:=k+1;
            if (RGod>=2010) and (FTYPE=2) then
             begin
               if not ((qqV[j]<>'') and (qqDV[j,i]<>0)) then s:='              <СвСумДох Месяц="'+SNMes(i)+'" КодДоход="'+qqD[j]+'" СумДоход="'+Fstr(qqMT[j,i],2)+'"/>' ;
               if (qqV[j]<>'') and (qqDV[j,i]<>0) then s:='              <СвСумДох Месяц="'+SNMes(i)+'" КодДоход="'+qqD[j]+'" СумДоход="'+Fstr(qqMT[j,i],2)+'">' ;
             end;


           if (qqV[j]<>'') and (qqDV[j,i]<>0) then
            begin
             if RGod>=2010 then
              begin
               s:=s+#13+'                 <СвСумВыч КодВычет="'+qqV[j]+'" СумВычет="'+Fstr(qqDV[j,i],2)+'"/>'
                   +#13+'                 </СвСумДох>';
              end;
            end;
           if (s<>'') and (FTYPE=2) then Writeln(FXML,s);



          end;
       end;
     end;



    if (Ftype=2) and (RGod>=2010) then WriteLn(FXML,'            </ДохВыч>');

   nd:=nd+2;



   

   xSumVic:=0;
   for j:=1 to 6 do
    begin
     xSumVic:=xSumVic+DST20[j,13];
    end;



    x:=0;
    for j:=1 to 12 do x:=x+DImVyc[j];

    x:=0;
    for j:=1 to 12 do x:=x+DImVyc[j];




      x:=0;
      for j:=1 to 6 do x:=x+DST20[j,13];
      x:=0;
      for j:=1 to 12 do x:=x+DImVyc[j];

    x:=0;
    for j:=1 to 10 do x:=x+qqMT[j,13];

    supl:=FUplataNdfl(form1.kartNLS.Value,RGod,13,'')+FUplataNdfl(form1.kartNLS.Value,RGod,30,'');
    if not RtfI then supl:=FUplataNdfl(form1.kartNLS.Value,RGod,13,soktmo)+FUplataNdfl(form1.kartNLS.Value,RGod,30,soktmo);




    if FType=2 then
     begin


       if RGod>=2015 then
        begin
       //  for i:=1 to 12 do x:=x+DDoxod9[i];
         for i:=1 to 12 do DISC[13]:=DISC[13]+DPn9[i];
         tisc:=DISC[13];
         tud:=DISC[13];
         if form_58.ndflr5.Locate('nls;god',VarArrayOf([form1.kartnls.Value,RGod]),[loCaseInsensitive]) then
           begin
            tisc:=form_58.ndflr5sisc.value;
            tud:=form_58.ndflr5sud.value;
           end;
         supl:=supl+FUplataNdfl(form1.kartNLS.Value,RGod,9,'');

        end;




         WriteLn(FXML,'</СведДох>');
         WriteLn(FXML,'</СправДох>');


        end;

     end;




   END;
   {конец 13%}




 

  END;

 END;




end;


procedure TForm_58.JvXPButton20Click(Sender: TObject);
var gm1,gm2,i:Integer;
    fCol,gOk,gOk2:Boolean;
    x,x0,y0:real;
    sNAme:String;
    Nstroka:integer;
    NewValueArray: OLEVariant;
begin

 form58:=TForm58.Create(Self);
 //form58.tCheckBox1.Visible:=true;
 form58.fm1:=1;
 form58.fm2:=12;
 form58.ShowModal;
 fCol:=form58.tCheckBox1.Checked;
 form58.ComboBox1.Visible:=false;
 gm1:=form58.fm1;
 gm2:=form58.fm2;
 gOk:=form58.TOk;
 form58.free;
 if not gOk then exit;


 sNAme:='Export';



   if CreateExcel then
    begin
      AddWorkBook;
      AddSheet(sNAme);
      VisibleExcel(true);
      SetWindowState(2);
    end
    else
      begin
       MessageDlg('Err Excel',mtError,[mbOk],0);
       exit;
      end;

  NewValueArray := VarArrayCreate([1, 1, 1,3], varVariant);
  NStroka:=1;
  NewValueArray[1,1]:='Отчет за период: '+namemes[gm1]+' - '+namemes[gm2]+' '+floattostr(RGod)+'г.';
  NewValueArray[1,2]:=''; NewValueArray[1,3]:='' ;
  FTmp(sNAme,'A'+InttoStr(Nstroka),'C'+InttoStr(Nstroka),newValueArray,false,'Calibri',9);

    NStroka:=Nstroka+2;
    NewValueArray[1,1]:='Начисление';
    NewValueArray[1,2]:='Код дохода'; NewValueArray[1,3]:='Сумма' ;
    FTmp(sNAme,'A'+InttoStr(Nstroka),'C'+InttoStr(Nstroka),newValueArray,true,'Calibri',9);

 datam.Query1.Close;
 datam.Query1.SQL.Clear;
 datam.Query1.DatabaseName:=form1.DBDIR;
 datam.Query1.SQL.Add('select g.*, k.* from glnew g, kart k where g.nls=k.nls and g.WG='+Floattostr(RGod));
 datam.Query1.SQL.Add('and g.wm>='+floattostr(gm1));
 datam.Query1.SQL.Add('and g.wm<='+floattostr(gm2));
 datam.Query1.Prepare;
 datam.Query1.Open;

 datam.Query1.First;
 x0:=0; y0:=0;
 while not datam.query1.eof do
  begin
    x:=0;
    if datam.Query1.FieldByName('DAYRAB').asFloat<>0 then
     x:=DRound(datam.Query1.FieldByName('OKLAD').asFloat*DelenieCas(datam.Query1.FieldByName('DAYOTR').asFloat,
         datam.Query1.FieldByName('DAYRAB').asFloat,datam.Query1.FieldByName('DAYCAS').asInteger),form1.DRZn);
    x:=x+DRound(x*form1.configRK.Value/100,2);
    x0:=x0+x;
   datam.query1.next;
  end;

 if DRound(x0,2)<>0 then
  begin
    y0:=y0+x0;
    NStroka:=Nstroka+1;
    NewValueArray[1,1]:='Оклад/тариф';
    NewValueArray[1,2]:='2000'; NewValueArray[1,3]:=x0 ;
    FTmp(sNAme,'A'+InttoStr(Nstroka),'C'+InttoStr(Nstroka),newValueArray,true,'Calibri',9);
  end;

  form1.nacisl.first;
  while not form1.nacisl.Eof do
   begin
    datam.Query1.Close;
    datam.Query1.SQL.Clear;
    datam.Query1.SQL.Add('select g.* from obrt1new g, kart k where g.nls=k.nls and g.WG='+Floattostr(RGod));
    datam.Query1.SQL.Add('and g.wm>='+floattostr(gm1));
    datam.Query1.SQL.Add('and g.wm<='+floattostr(gm2));
    datam.Query1.SQL.Add('and g.kod='+floattostr(form1.nacislKOD.Value));
    datam.Query1.Prepare;
    datam.Query1.Open;
    datam.Query1.First;
    x0:=0 ;
    while not datam.query1.eof do
     begin
       x:=datam.Query1.FieldByName('KR').asFloat;
       if form1.NACISLRK.Value then x:=x+DRound(x*form1.configRK.Value/100,2);
       x0:=x0+x;
       datam.query1.next;
     end;

     if DRound(x0,2)<>0 then
      begin
       y0:=y0+x0;
       NStroka:=Nstroka+1;
       NewValueArray[1,1]:=form1.NACISLNAME.value;
       NewValueArray[1,2]:=form1.NACISLKODDOX.Value; NewValueArray[1,3]:=x0 ;
       FTmp(sNAme,'A'+InttoStr(Nstroka),'C'+InttoStr(Nstroka),newValueArray,true,'Calibri',9);
      end;


    form1.nacisl.next;
   end;

       NStroka:=Nstroka+1;
       NewValueArray[1,1]:='';
       NewValueArray[1,2]:=''; NewValueArray[1,3]:=y0 ;
       FTmp(sNAme,'A'+InttoStr(Nstroka),'C'+InttoStr(Nstroka),newValueArray,false,'Calibri',9);

  // for i:=1 to 3  do SetColumnAutoFit(sName,i)  ;


   SetColumnWidth(sName,1,30);
   SetColumnWidth(sName,2,9);
   SetColumnWidth(sName,3,11);
   PRowsHeight(sName,3,20);
   for i:=4 to NStroka do PRowsHeight(sName,i,18);
   SetVerticalAlignment(sName,'A3','C'+IntToStr(Nstroka),-4108);
   SetHorizontalAlignment(sName,'B4','B'+IntToStr(Nstroka),-4108);
   SetHorizontalAlignment(sName,'A3','C3',-4108);
   PZalivkaColor(sName,'A3','C3',WColorRep);

  //  SetShirina(sName);

   SetFormatRange(sName,'C3','C'+IntToStr(NStroka),'0.00');    

   
   SetWindowState(-4137);
   PUnAssigned;

end;

procedure TForm_58.JvXPButton21Click(Sender: TObject);
begin
 form3310:=tform3310.create(nil);
 form3310.showmodal;
 form3310.free;
end;

procedure TForm_58.N20231Click(Sender: TObject);
var tNls:Real;
    i:integer;
begin
  RPriznak:=trim(Edit9.TExt);


form1.kart.first;
 while not form1.kart.Eof do
  begin
   form1.kart.edit;
   form1.kart.fieldbyname('g').asString:='';
   form1.kart.post;
   form1.kart.next;
  end;
 form2.TMultiSelect:=True;
 datam.SetIndexKart('FAM');
 form2.JvXPButton4.Visible:=True;
 form2.DBGrid1.Columns[0].Visible:=True;  //Check in DBGrid
 Form2.ShowModal;
 tNls:=form1.kartNLS.Value;
 form2.DBGrid1.Columns[0].Visible:=false;  //Check in DBGrid
 form2.JvXPButton4.Visible:=True;
 form2.TMultiSelect:=False;

 if not form2.TOK then exit;

 tIsprDec2022:=false;
 if RGOD=2022 then if MessageDlg('Исключить из справок выплату з/п за вторую половину декабря 2022г, которая была проведена в январе 2023г (если такая есть)'+#13+
   'Предварительно должна быть проведена обработка (кнопка Декабрь 2022)',mtInformation,[mbYes,mbNo],0) = mrYes then tIsprDec2022:=true;


 form1.kart.first;
 i:=0;
 while not form1.kart.Eof do
  begin
   if form1.kart.FieldByName('G').asString='*' then
      begin
       i:=i+1;
       form_58.PSpr20192(form1.kartNls.Value,false);

      end;
   form1.kart.Next;
  end;

 if i=0 then
  begin
    form1.kart.locate('nls',tNls,[loCaseInsensitive]);
    form_58.PSpr20192(form1.kartnls.Value,false);
  end;

end;

end.
