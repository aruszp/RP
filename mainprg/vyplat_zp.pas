unit vyplat_zp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  OffBtn, StdCtrls, ExtCtrls,mainLib,reportFM,DB, DBTables, Grids, DBGrids,
  ComCtrls, RXCtrls, numTools, Buttons, ShellAPi, JvExControls,
  JvComponent, JvXPCore, JvXPButtons, Mask, ToolEdit, CurrEdit, Menus,ComObj,
  FR_Class, Reindex, excellib, nalkarta;

type
  TForm123 = class(TForm)
    Panel1: TPanel;
    vyplzp: TTable;
    vyplzpWM: TSmallintField;
    vyplzpWG: TSmallintField;
    vyplzpTYPE: TSmallintField;
    vyplzpSUMMA: TFloatField;
    vyplzpKOD: TSmallintField;
    DBGrid1: TDBGrid;
    Query1: TQuery;
    DataSource1: TDataSource;
    RadioGroup1: TRadioGroup;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    vyplzpSUMMA0: TFloatField;
    Label4: TLabel;
    Edit2: TEdit;
    RxLabel1: TRxLabel;
    Panel2: TPanel;
    RxLabel2: TRxLabel;
    Label1: TLabel;
    Edit1: TEdit;
    vyplzpNUMDOK: TFloatField;
    vyplzpDATDOK: TDateField;
    frReport1: TfrReport;
    vyplzpsumpropis: TStringField;
    vyplzpssumma: TStringField;
    vyplzpPROVEDEN: TSmallintField;
    vyplzpNOTE: TStringField;
    JvXPButton24: TJvXPButton;
    JvXPButton23: TJvXPButton;
    JvXPButton1: TJvXPButton;
    JvXPButton3: TJvXPButton;
    JvXPButton4: TJvXPButton;
    JvXPButton5: TJvXPButton;
    JvXPButton7: TJvXPButton;
    Label7: TLabel;
    Label8: TLabel;
    JvXPButton12: TJvXPButton;
    CurrencyEdit1: TCurrencyEdit;
    CurrencyEdit2: TCurrencyEdit;
    JvXPButton13: TJvXPButton;
    JvXPButton14: TJvXPButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    PopupMenu2: TPopupMenu;
    JvXPButton8: TJvXPButton;
    N3: TMenuItem;
    N4: TMenuItem;
    vyplzpIDP: TFloatField;
    JvXPButton40: TJvXPButton;
    JvXPButton41: TJvXPButton;
    JvXPButton42: TJvXPButton;
    Label2: TLabel;
    CurrencyEdit3: TCurrencyEdit;
    JvXPButton44: TJvXPButton;
    RadioButton4: TRadioButton;
    vyplzpocer: TStringField;
    PopupMenu3: TPopupMenu;
    N21998010620141: TMenuItem;
    N010620141: TMenuItem;
    JvXPButton28: TJvXPButton;
    JvXPButton280: TJvXPButton;
    vyplzpG: TStringField;
    N5: TMenuItem;
    JvXPButton99: TJvXPButton;
    ComboBox1: TComboBox;
    Label3: TLabel;
    JvXPButton268: TJvXPButton;
    JvXPButton59: TJvXPButton;
    JvXPButton295: TJvXPButton;
    Label5: TLabel;
    JvXPButton222: TJvXPButton;
    vavans: TTable;
    vavansNLS: TFloatField;
    vavansFIO: TStringField;
    vavansNAME: TStringField;
    vavansOKLAD: TFloatField;
    vavansDAYRAB: TFloatField;
    vavansDAYOTR: TFloatField;
    vavansRK: TFloatField;
    vavansNADBAVKA: TFloatField;
    vavansSUMMA: TFloatField;
    DataSource2: TDataSource;
    JvXPButton233: TJvXPButton;
    vdoxod: TTable;
    vdoxodKOD: TFloatField;
    vdoxodNAME: TStringField;
    vdoxodSUMMA: TFloatField;
    vdoxodSNALOG: TFloatField;
    DataSource3: TDataSource;
    vdoxodOI: TFloatField;
    vdoxodsumma2: TFloatField;
    vyplzpKODNAC: TFloatField;
    vyplzpSDOXOD: TFloatField;
    vyplzpSNALOG: TFloatField;
    vyplzpGOD: TFloatField;
    vyplzpMES: TFloatField;
    vdoxodsnalog2: TFloatField;
    vdoxodoi2: TFloatField;
    vyplzpDAT: TDateField;
    vyplzpIDSDOXOD: TFloatField;
    JvXPButton299: TJvXPButton;
    Button1: TButton;
    RadioGroup2: TRadioGroup;
    vyplzpnls: TFloatField;
    Label6: TLabel;
    vavansVICET: TFloatField;
    JvXPButton258: TJvXPButton;
    JvXPButton_222: TJvXPButton;
    JvXPButton298: TJvXPButton;
    N21: TMenuItem;
    N010620142: TMenuItem;
    JvXPButton9: TJvXPButton;
    JvXPButton555: TJvXPButton;
    JvXPButton288: TJvXPButton;
    JvXPButton315: TJvXPButton;
    vyplzpAVANS: TFloatField;
    CurrencyEdit4: TCurrencyEdit;
    Label9: TLabel;
    CheckBox199: TCheckBox;
    vyplzpAVDOXOD: TFloatField;
    vyplzpAVNALOG: TFloatField;
    vyplzpAVVICET: TFloatField;
    JvXPButton16: TJvXPButton;
    PopupMenu4: TPopupMenu;
    N6: TMenuItem;
    JvXPButton201: TJvXPButton;
    PopupMenu5: TPopupMenu;
    N8: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N9: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    JvXPButton404: TJvXPButton;
    PopupMenu6: TPopupMenu;
    N7: TMenuItem;
    N20: TMenuItem;
    N22: TMenuItem;
    JvXPButton20: TJvXPButton;
    procedure FormCreate(Sender: TObject);
    Procedure PlatVed;
    Procedure PlatVed2;
    function ZaprosOINew(xs:integer):Real;
    procedure FormShow(Sender: TObject);
    procedure ZaprosV1(tParam:integer);
        procedure ZaprosV2(mk:integer);
        procedure ZaprosV3;

    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure vyplzpsumpropisGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure vyplzpssummaGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    function ProvDAtVypl():Boolean;
    procedure DBGrid1DblClick(Sender: TObject);
    procedure JvXPButton24Click(Sender: TObject);
    procedure JvXPButton23Click(Sender: TObject);
    procedure JvXPButton1Click(Sender: TObject);
    procedure JvXPButton6Click(Sender: TObject);
    procedure JvXPButton2Click(Sender: TObject);
    procedure JvXPButton3Click(Sender: TObject);
    procedure JvXPButton4Click(Sender: TObject);
    procedure JvXPButton5Click(Sender: TObject);
    procedure JvXPButton7Click(Sender: TObject);
    procedure JvXPButton12Click(Sender: TObject);
    procedure JvXPButton13Click(Sender: TObject);
    procedure JvXPButton14Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure JvXPButton8Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure JvXPButton40Click(Sender: TObject);
    procedure JvXPButton41Click(Sender: TObject);
    procedure JvXPButton42Click(Sender: TObject);
    procedure JvXPButton44Click(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);
    procedure vyplzpocerGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure N21998010620141Click(Sender: TObject);
    procedure JvXPButton28Click(Sender: TObject);
    procedure JvXPButton280Click(Sender: TObject);
    procedure DrawGridCheckBox(Canvas: TCanvas; Rect: TRect; Checked: boolean);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N5Click(Sender: TObject);
    procedure JvXPButton99Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure JvXPButton268Click(Sender: TObject);
    procedure JvXPButton59Click(Sender: TObject);
    procedure JvXPButton295Click(Sender: TObject);
    procedure N010620141Click(Sender: TObject);
    procedure JvXPButton222Click(Sender: TObject);
    procedure JvXPButton233Click(Sender: TObject);
    procedure DeleteDoxodNac;
    procedure CheckBox1Click(Sender: TObject);
    procedure JvXPButton299Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure RadioGroup2Click(Sender: TObject);
    Function ZaprosOiDoxod(TNls:Real;tParam:Integer):Real;  //остатки по кодам доходов с учетом выплаты
    function ProvAvansKod(xNls:Real;sfio:String):Boolean;
    procedure JvXPButton258Click(Sender: TObject);
    procedure JvXPButton_222Click(Sender: TObject);
    procedure JvXPButton298Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N010620142Click(Sender: TObject);
    procedure JvXPButton9Click(Sender: TObject);
    procedure JvXPButton555Click(Sender: TObject);
    function Prov2Avans(xNls:Real):Integer;     //проверяет два и более аванса наличие
    procedure DeleteNullKr;
    procedure JvXPButton315Click(Sender: TObject);
    procedure ZaprosNOAvans;
    procedure CheckBox199Click(Sender: TObject);
    procedure Favans(sDoxod,sVicet:real;var sAvans,sNdfl:Real);
    procedure JvXPButton16Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure JvXPButton201Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure JvXPButton404Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure JvXPButton288Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure ControlMrot;
    procedure JvXPButton20Click(Sender: TObject);

  private
   xD1,xD2:TDAte;
   xTYPE:Integer;
    { Private declarations }
   isxrtf:string;
   xKoduderj:Integer;
   TXS:Integer;
   sSoobVyp:String;
  public
   fName:String;
   fDat:TDate;
   _TPKOD:Integer;
   NPVed:Integer;
   TYPEVYPLAT:Integer;

   tOstatok:Real;

   UWKOD:array[0..100] of Boolean;
   T_OWKOD:Boolean; //признак что фильтр по UWKOD
   TVYBOR:Boolean;
   TUpd:Boolean;
   TUpdNls:Real;
    { Public declarations }
  end;

var
  Form123: TForm123;

implementation

uses pevazp, nac_new_sp, memoedit, filtr, vyb_nac, ed_vyplzp, inf_basa,
  MyDATAMODULE, RasVedP, vypl_proveden, vvod_npved, vyb_mnogo, vybData,
  kassakn, FormTF, kartrab, klient, vavans, FSpr2006, vypl_doxod_casticno,
  Vyb_Mes, uplatadox333, uplatadoxod, uplatadoxvvod, upldoxv2, npved_edit,
  NPVed, vvodnpackapfr, PrikazSV, upld333_filter, FWait, avans_obrab,
  vyplatzp_vybor, avans_ed2023, ParamVicet, vybperiod, filial, doplMROT;

{$R *.DFM}




procedure TForm123.ControlMrot;
var xs:Real;
    xMrot,xMrot0:Real;
    xDayRab,xDayOtr:Real;
    kuc:Real;
    oldIndex:String;
begin
 xMrot:=0;
 form1.mrot.First;
 while not form1.mrot.Eof do
  begin
   if EncodeDate(RGod,RMes,1)>=form1.mrotDAT.Value then xMROT:=form1.mrotMROT.Value;
   form1.mrot.Next;
  end;
 xMrot:=xMrot+DRound(xMrot*form1.configRK.Value/100,2);
 xMrot0:=xMrot;

 if MessageDlg('Выполнить Контроль начислений по сотрудникам до МРОТ'+#13+
     'МРОТ='+FloattostrF(xMROT,ffNumber,12,2)+'руб'+#13+
      'В контроль попадают только сотрудники по трудовым договорам' ,mtInformation,[mbYes,mbNo],0) = mrNo then exit;


 form451:=TForm451.Create(nil);

 form451.donmrot.Active:=false;
 form451.donmrot.DatabaseName:=form1.DBDIR;
 if not form451.donmrot.Exists then form451.donmrot.CreateTable;
 form451.donmrot.Exclusive:=true;
 form451.donmrot.EmptyTable;
 form451.donmrot.active:=true;

 oldIndex:=form1.kart.IndexName;
 form1.kart.IndexName:='FAM';


 form1.kart.first;
 while not form1.kart.eof do
  begin
   xs:=0;
   datam.Query1.close;
   datam.query1.SQL.Clear;
   datam.query1.SQL.add('select * from glnew where nls='+floattostr(form1.kartnls.value));
   datam.query1.SQL.add('and wg='+floattostr(rgod));
   datam.query1.SQL.add('and wm='+floattostr(rmes));
   datam.query1.Prepare;
   datam.query1.open;
   if datam.query1.fieldbyname('dayrab').asFloat<>0 then
     xs:=mainlib.DRound(datam.query1.fieldbyname('oklad').asFloat*DelenieCas(datam.query1.fieldbyname('dayotr').asFloat,
               datam.query1.fieldbyname('dayrab').asFloat,datam.query1.fieldbyname('daycas').asFloat),form1.DRZn)
                                    else xs:=0;

   xs:=xs+DRound(xs*form1.configRK.Value/100,form1.DRZn) ;
   xDayRab:=datam.query1.fieldbyname('dayrab').asFloat   ;
   xDayOtr:=datam.query1.fieldbyname('dayOtr').asFloat   ;
   xMrot:=0;
   if xDayOtr>xDayRab then xDayOtr:=xDayRab;
   if xDayRab<>0 then xMrot:=DRound(xMrot0*xDayOtr/xDayRab,2);



   datam.Query1.close;
   datam.query1.SQL.Clear;
   datam.query1.SQL.add('select o.*, n.* from obrt1new o, nacisl n where o.kod=n.kod and o.nls='+floattostr(form1.kartnls.value));
   datam.query1.SQL.add('and o.wg='+floattostr(rgod));
   datam.query1.SQL.add('and o.wm='+floattostr(rmes));
   datam.query1.SQL.add('and n.DMROT=TRUE');

   datam.query1.Prepare;
   datam.query1.open;
   datam.query1.first;
   while not datam.query1.eof do
    begin
     xs:=xs+datam.query1.fieldbyname('kr').asfloat;
     if datam.query1.fieldbyname('rk').asBoolean then
       xs:=xs+DRound(datam.query1.fieldbyname('kr').asfloat*form1.configRK.Value/100,form1.DRZn) ;
     datam.query1.next;
    end;


  kuc:=1;

  form1.staj.First;
  while not form1.staj.Eof do
   begin
      if form1.stajD1.Value<=EncodeDAte(RGod,RMes,1) then kuc:=form1.stajN.Value;
    form1.staj.next;
   end;


 // if kuc<>1 then ShowMEssage(form1.kartfam.value+' xs='+floattostr(xs)+' мрот='+floattostr(xmrot)+#13+'kuc='+floattostr(kuc));


//  if (xMrot*kuc>xs) and (Ftypedog(form1.kartnls.value)=0) then
  if (Ftypedog(form1.kartnls.value)=0) then
   begin
    form451.donmrot.append;
    form451.donmrot.fieldbyname('nls').asFloat:=form1.kartNls.Value;
    form451.donmrot.fieldbyname('dayrab').asFloat:=xDayRab;
    form451.donmrot.fieldbyname('dayotr').asFloat:=xDayOtr;
    form451.donmrot.fieldbyname('fio').asstring:=form1.kartFam.Value+' '+form1.kartim.Value+' '+form1.kartot.Value ;
    if kuc<>1 then
         form451.donmrot.fieldbyname('fio').asstring:=form451.donmrot.fieldbyname('fio').asstring+' /k='+floattostr(kuc)+'/';
    form451.donmrot.FieldByNAme('mrot0').asFloat:=DRound(xMrot0,2);
    form451.donmrot.FieldByNAme('mrot').asFloat:=DRound(xMrot*kuc,2);
    form451.donmrot.FieldByNAme('summa').asFloat:=xs;
    form451.donmrot.FieldByNAme('donmrot').asFloat:=DRound(xMrot*kuc-xs,2);
    if DRound(form451.donmrot.FieldByNAme('donmrot').asFloat,2)<0 then form451.donmrot.FieldByNAme('donmrot').asFloat:=0;
    form451.donmrot.Post;
   end;



   form1.kart.next;
  end;

  form1.kart.IndexName:=oldIndex;

  form451.ShowModal;
  form451.Free;

end;


procedure TForm123.DeleteNullKr;
var yIdp,xNls,xIdp:Real;
    xIdsDoxod:Real;
    rtf:boolean;
begin


  yIdp:=query1.fieldbyname('idp').asFloat;
  rtf:=false;

  query1.First;
  while not query1.eof do
   begin
     if (query1.fieldbyname('proveden').asFloat=1) and (Dround(query1.fieldbyname('summa').asFloat,2)=0) then
       begin
        rtf:=true;
        xIdp:=query1.fieldbyname('idp').asFloat;
        xNls:=query1.fieldbyname('nls').asFloat;
        if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
         begin
          form1.kart.Locate('nls',xNls,[loCaseInsensitive]);
        //  Showmessage(form1.kartfam.value);

          if form1.obrt2.Locate('nls;numdok',VarArrayOf([xNls,xIdp]),[loCAseInsensitive]) then
           begin
            form1.DelSDoxodObrt2(form1.obrt2ID.Value);
            form1.obrt2.Delete;
           end;
             xIdsDoxod:=vyplzp.FieldByName('idsdoxod').asFloat;
             if form1.sdoxod.Locate('id',xIdsdoxod,[loCaseInsensitive]) then form1.sdoxod.Delete;

           vyplzp.delete;
          end;
        end;
    query1.Next;
   end;

   if rtf then
    begin
     vyplzp.FlushBuffers;
//     ZaprosV2(0);
//     JvXPButton12Click(nil);
       JvXPButton6Click(nil);
    end;
    
end;


Function TForm123.ZaprosOiDoxod(TNls:Real;tParam:Integer):Real;  //остатки по кодам доходов с учетом выплаты
var x,z,xSumma3,xSumma4:Real;
    tMes,tKod,tgod:Integer;
    xD:TDate;
    tGod2:Integer;
    s:String;
begin
   datam.qtmp.close;
   datam.qtmp.DatabaseName:=form1.dbdir;
   datam.qtmp.sql.clear;
   datam.qtmp.sql.add('select OW from kart where nls='+floattostr(Tnls));
   datam.qtmp.prepare;
   datam.qtmp.open;
   if DRound(datam.qtmp.fields[0].asFloat,2)>0 then tGod2:=RGod-1 else tGod2:=1999;
   //если есть остатков в карточке то запрос предыдущий год
   datam.qtmp.close;





   datam.Query2.DatabaseName:=form1.DBDIR;

   datam.Query1.Close;
   datam.Query1.DatabaseName:=form1.DBDIR;
   datam.Query1.SQL.Clear;
   datam.Query1.SQL.Add('SELECT * from glnew ');
   datam.Query1.SQL.Add('where ( (wm<='+IntToStr(RMes)+' and wg='+IntToStr(RGod)+') or wg<='+Inttostr(tGod2)+' ) and nls='+floattostr(TNls));
   datam.Query1.SQL.Add('and dayrab*dayotr*oklad>0');     //убрать задвоение
   if (tParam=1) and (not UWKOD[0]) then datam.Query1.SQL.Add('and nls<-100');
   datam.Query1.SQL.Add('order by datoklad');
   datam.Query1.Prepare;
   datam.Query1.Open;
   datam.Query1.first;


   xSumma4:=0; //сумма к выплате
   while not datam.Query1.Eof do
    begin
      x:=0;
      tMes:=datam.Query1.FieldByName('WM').asInteger;
      tGod:=datam.Query1.FieldByName('WG').asInteger;

      //  Showmessage(floattostr(tmes)+#13+floattostr(tgod));

      if datam.Query1.RecordCount>0 then
       if datam.Query1.FieldByName('dayrab').asFloat<>0 then
                x:=mainlib.DRound(datam.Query1.FieldByName('oklad').asFloat*
                  DelenieCas(datam.Query1.FieldByName('dayotr').asFloat,
                    datam.Query1.FieldByName('dayrab').asFloat,datam.Query1.FieldByName('DAYCAS').asFloat),form1.DRZn)  ;
                x:=x+DRound(x*form1.configRK.Value/100,2);
      x:=DRound(x,2);
      if datam.Query1.FieldByName('DATOKLAD').asDateTime<EncodeDate(2001,1,1) then xD:=EncodeDate(2001,1,1) else xD:=datam.Query1.FieldByName('DATOKLAD').asDateTime;
      z:=datam.Query1.FieldByName('snalog').asFloat;



     IF xD=EncodeDate(2001,1,1) THEN   //ЕСЛИ ДАТА ЗАПОЛНЕНА ТО ВСЕ ВЫПЛАЧЕНО !
      BEGIN
          xSumma3:=x-z;

           //начало обработки частичных выплат дохода
          datam.Query2.DataBaseName:=form1.DBDIR;
          datam.Query2.Close;
          datam.Query2.SQL.Clear;
          datam.Query2.SQL.Add('select * from sdoxod where sdoxod<>0 and  kodnac=0');
          datam.Query2.SQL.Add('and god='+IntToStr(tGod)+' and mes='+IntToStr(tMes));
          datam.Query2.SQL.Add('and nls='+FloatToStr(TNls));
          datam.Query2.Prepare;
          datam.Query2.Open;
          datam.Query2.first;
           while not datam.Query2.Eof do    //чяастичное проведение разбиваем
            begin
              xSumma3:=xSumma3-(datam.Query2.FieldByName('sdoxod').asFloat-datam.Query2.FieldByName('nalog').asFloat);
             datam.Query2.next;
            end;
            datam.Query2.Close;
          xSumma4:=xSumma4+xSumma3;

          if (DRound(xSumma3,2)<>0) and (tParam=0) then UWKOD[0]:=true;     //заполняем код выплаты что есть оклад


   //     if dround(xsumma3,2)<>0 then   showmessage(floattostr(tnls)+#13+'Оклад мес='+inttostr(tMes)+#13+floattostr(xSumma3));
      END;
     datam.Query1.Next;
    end;
   datam.Query1.Close;



 //*******************************************************************************

   datam.Query2.Close;
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('SELECT datprov,kod,wm,wg,sum(kr) as kr, sum(snalog) as snalog from obrt1new ');
   datam.Query2.SQL.Add('where ((wm<='+IntToStr(RMes)+' and wg='+IntToStr(RGod)+') or wg<=+'+intToStr(tGod2)+' ) and nls='+floattostr(tNls));
   datam.Query2.SQL.Add('group by datprov,kod,wm,wg');
   datam.Query2.Prepare;
   datam.Query2.Open;
   datam.Query2.First;
   while not datam.Query2.Eof do
    begin
     tKod:=datam.Query2.FieldByName('kod').asInteger;

     if (datam.Query2.FieldByName('wm').asInteger>0) and ((tParam=0) or (tParam=1) and (UWKOD[tKod])) then
     begin
      tMes:=datam.Query2.FieldByName('wm').asInteger;
      tGod:=datam.Query2.FieldByName('wg').asInteger;

      if datam.Query2.FieldByName('DATPROV').asDateTime<EncodeDate(2001,1,1) then xD:=EncodeDate(2001,1,1) else xD:=datam.Query2.FieldByName('DATPROV').asDateTime;
      IF xD=EncodeDate(2001,1,1) THEN
      BEGIN
       form1.NACISL.Locate('KOD',tKod,[loCaseInsensitive]);
      x:=0;
    //  if (form1.NACISLpn.Value<>1) then x:=datam.Query2.FieldByName('kr').asFloat;
      x:=datam.Query2.FieldByName('kr').asFloat;
      if form1.NACISLRK.Value then x:=x+DRound(x*form1.configRK.Value/100,2);
      x:=DRound(x,2);
      z:=datam.QUery2.fieldbyname('snalog').asFloat;

      xSumma3:=x-z;

       //начало обработки частичных выплат дохода
      datam.Query1.DataBaseName:=form1.DBDIR;
      datam.Query1.Close;
   datam.Query1.SQL.Clear;
   datam.Query1.SQL.Add('select * from sdoxod where sdoxod<>0  and kodnac='+floattostr(tKod));
   datam.Query1.SQL.Add('and god='+IntToStr(tGod)+' and mes='+IntToStr(tMes));
   datam.Query1.SQL.Add('and nls='+FloatToStr(TNls));
   datam.Query1.Prepare;
   datam.Query1.Open;
   datam.Query1.first;
   while not datam.Query1.Eof do    //чяастичное проведение разбиваем
    begin
       xSumma3:=xSumma3-(datam.Query1.FieldByName('sdoxod').asFloat-datam.Query1.FieldByName('nalog').asFloat);
     datam.Query1.next;
    end;
    datam.Query1.Close;

    if (DRound(xSumma3,2)<>0) and (tParam=0) then UWKOD[tKOd]:=true;

    xSumma4:=xSumma4+xSumma3;
  //  if dround(xsumma3,2)<>0 then  showmessage(floattostr(tnls)+#13+'Начисл. код ='+inttostr(tKod)+#13+'мес='+inttostr(tMes)+#13+floattostr(xSumma3));

   END;
   end;
  datam.Query2.Next;
 end;
 datam.query2.close;
//if dround(xsumma4,2)<>0 then showmessage('Итого: '+floattostr(tnls)+#13+floattostr(xSumma4));

//    ShowMessage(floattostr(xSumma4));


  ZaprosOiDoxod:=DRound(xSumma4,2);


end;



procedure TForm123.DeleteDoxodNac;  //удаляет ссылки на частичную выплату дохода при первом входе
begin
 if TYPEVYPLAT=2 then exit;
 Query1.first;
// ShowMessage('Удаляем');
 while not Query1.Eof do
  begin
   if (Query1.FieldbyNAme('proveden').asFloat<>1) and  (vyplzp.Locate('idp',Query1.FieldbyName('idp').asFloat,[loCaseInsensitive])) then
    begin
     vyplzp.edit;
     vyplzp.fieldbyname('snalog').asFloat:=0;
     vyplzp.fieldbyname('sdoxod').asFloat:=0;
     vyplzp.fieldbyname('god').asFloat:=0;
     vyplzp.fieldbyname('mes').asFloat:=0;
     vyplzp.fieldbyname('kodnac').asFloat:=0;
     vyplzp.fieldbyname('dat').asdatetime:=encodedate(1899,12,31);
     vyplzp.post;
    end;
   Query1.Next;
  end;
  Query1.first;
  ZaprosV2(0);
end;



procedure Tform123.DrawGridCheckBox(Canvas: TCanvas; Rect: TRect; Checked: boolean);
var
  DrawFlags: Integer;
begin
  Canvas.TextRect(Rect, Rect.Left + 1, Rect.Top + 1, ' ');
  DrawFrameControl(Canvas.Handle, Rect, DFC_BUTTON, DFCS_BUTTONPUSH or DFCS_ADJUSTRECT);
  DrawFlags := DFCS_BUTTONCHECK or DFCS_ADJUSTRECT;// DFCS_BUTTONCHECK
  if Checked then
    DrawFlags := DrawFlags or DFCS_CHECKED;
  DrawFrameControl(Canvas.Handle, Rect, DFC_BUTTON, DrawFlags);
end;



procedure TForm123.FormCreate(Sender: TObject);
begin
 sSoobVyp:='Номер платежной ведомости и комментарии можно корректировать через соответствующую кнопку в данном модуле <Плат.вед.>';

 NPVed:=20;
 xTYPE:=3; {пусто запросит}
end;

Procedure TForm123.PlatVed;
var i:Integer;
    F:textFile;
    x,x0:Real;
    isxrtf,outrtf:String;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
begin
 form93.Query1.DatabaseName:=form1.DBDIR;
 isxrtf:=GetCurrentDir()+'\ISXRTF\platVed.rtf';
 outrtf:=GetCurrentDir()+'\OUTRTF\_platVed.rtf';

 AssignFile(F,'dat.txt');
 Rewrite(F);
 Query1.first;
  i:=0;
  x0:=0;

  while not Query1.eof do
   begin
    rtf:=false;
    x:=Query1.FieldByName('SUMMA').asFloat ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=1 then rtf:=True;
     end;

  {  if Query1.FieldByName('datdok').asDAteTime<>xD then rtf:=False;
  }
    if Query1.fieldByName('PROVEDEN').asInteger<>1 then rtf:=False;

    IF (x<>0) and (rtf)  then
     begin
     x0:=x0+x;
     i:=i+1;
     if i=1 then
      begin
       Write(F,'|9:NAME:'+form1.config2NAME.Value);
       Write(F,'|9:PERIOD:'+namemes[RMEs]+' '+IntToStr(RGod)+'г.');
       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       WriteLN(F,'|9:RUKOV:'+xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.');
      end;

      form1.kart.Locate('nls',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]);
      write(F,'|8:NPP:'+IntToStr(i));
      write(F,'|8:FIO:'+form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value);
      writeLn(F,'|8:SUMMA:'+FloatToStrF(x,ffNumber,12,2));



      if i=form123.NPVed then
       begin
        writeLn(F,'|5:SUMMA:'+FloatToStrF(x0,ffNumber,12,2));
        writeLn(F,'|4:');

        i:=0;
        x0:=0;
       end;
     END;
    Query1.next;
   end;
  if i<>0 then writeLn(F,'|5:SUMMA:'+FloatToStrF(x0,ffNumber,12,2));


 CloseFile(F);

 if i<>0 then
  begin
   reportFM.rtf_CreateReport('dat.txt',isxrtf,outrtf,nil,i);
   form20.StartWord(outrtf);
  end
   else MessageDlg('Ничего не сформировано, проверьте'+#13+
            '- нет проведенных оборотов (сначала необходимо нажать на Провести выплату)'+#13+
            {- проверьте параметр Дата выплаты '+#13+  }
            '- в Справочнике удержаний (выплат) неверно установлены флаги документов по выплате из кассы',mtInformation,[mbOk],0);
end;


function Tform123.ZaprosOINew(xs:integer):Real;
var xOi,x:Real;
    xOi0:real;
    i:Integer;
begin
 // xs=-1 без учета статьи
 // xs>=0  ID статьи

 ZapolnMas;

 xoi:=0; xOi0:=0;
 if (xs=0) or (xs=-1) then xOi:=form1.kartOW.Value;
 xOi0:=form1.kartOW.Value;

 if xs>=0 then
  begin
   form65:=TForm65.Create(nil);
   form65.ZapolnDF;
   for i:=1 to RMes do xoi:=xoi-form65.DF[xs,6,i];
   form65.Free;
  end;


 for i:=1 to RMes do xOi0:=xOi0-DPn1[i]-DPn9[i]-DPn35[i];
 for i:=1 to RMes do
    begin
     if form1.GlDb.Locate('NLS;WM;WG',VarArrayOf([form1.kartNLS.VAlue,i,RGod]),[loCaseInsensitive]) then
      begin
       xOi0:=xOi0-form1.GlDbKASSA.Value;
       if form1.GlDbDAYRAB.Value<>0 then
          x:=DRound(form1.gldbOklad.Value*Deleniecas(form1.gldbDayOtr.Value,form1.gldbDayRab.Value,form1.GLDBDAYCAS.Value),form1.DRZn)
             else x:=0;
       xoi0:=xOi0+x+DRound(x*form1.configRK.Value/100,2);
      end;
    end;

 if (xs=0) or (xs=-1) then
 begin
   if xs=-1 then for i:=1 to RMes do xOi:=xOi-DPn1[i]-DPn9[i]-DPn35[i];
   for i:=1 to RMes do
    begin
     if form1.GlDb.Locate('NLS;WM;WG',VarArrayOf([form1.kartNLS.VAlue,i,RGod]),[loCaseInsensitive]) then
      begin
       xOi:=xOi-form1.GlDbKASSA.Value;
       if form1.GlDbDAYRAB.Value<>0 then
          x:=DRound(form1.gldbOklad.Value*Deleniecas(form1.gldbDayOtr.Value,form1.gldbDayRab.Value,form1.GLDBDAYCAS.Value),form1.DRZn)
             else x:=0;
       xoi:=xOi+x+DRound(x*form1.configRK.Value/100,2);
      end;
    end;
  end;


   form93.Query1.Close;
   form93.Query1.SQl.Clear;
   form93.Query1.SQl.Add('select o.kr, o.kod, o.xs, o.tavans from obrt2new o');
   form93.Query1.SQl.Add('where o.wm<='+IntToStr(RMEs)+' and o.wg='+Inttostr(RGod) );
   form93.Query1.SQl.Add('and o.nls='+floatToStr(form1.kartNls.Value));
   form93.Query1.Prepare;
   form93.Query1.Open;
   form93.Query1.First;

     while not form93.Query1.Eof do
      begin

       if form93.Query1.FieldByName('tavans').asFloat<>1 then
        begin
          xOi0:=xOi0-form93.Query1.FieldByName('KR').asFloat;
          if (xs=-1) then xOi:=xOi-form93.Query1.FieldByName('KR').asFloat;
          if (xs>0) and (xs=form93.Query1.FieldByName('xs').asFloat) then xOi:=xOi-form93.Query1.FieldByName('KR').asFloat;
          if (xs=0) then   //основн бюджет
           begin
            if (form93.Query1.FieldByName('xs').asFloat=-1) or (xs=form93.Query1.FieldByName('xs').asFloat) then xOi:=xOi-form93.Query1.FieldByName('KR').asFloat;
           end;
        end;
        
       form93.Query1.Next;
      end;


   form93.Query1.Close;
   form93.Query1.SQl.Clear;
   form93.Query1.SQl.Add('select n.rk, o.kr, o.kod, n.idstatya from nacisl n, obrt1new o');
   form93.Query1.SQl.Add('where n.kod=o.kod and o.wm<='+IntToStr(RMEs)+' and o.wg='+Inttostr(RGod) );
   form93.Query1.SQl.Add('and o.nls='+floatToStr(form1.kartNls.Value));
   form93.Query1.Prepare;
   form93.Query1.Open;
   form93.Query1.First;

     while not form93.Query1.Eof do
      begin
        x:=form93.Query1.FieldByName('KR').asFloat;
        if form93.Query1.FieldByName('RK').asBoolean then
           x:=x+DRound(form93.Query1.FieldByName('KR').asFloat*form1.configRK.Value/100,2);

        if (xs=-1) or (xs=form93.Query1.FieldByName('idstatya').asFloat) then xoi:=xOi+x;
        xoi0:=xOi0+x;

       form93.Query1.Next;
      end;


  xOi:=DRound(xOi,2);
  xOi0:=DRound(xOi0,2);

  if xOi>xOi0 then xOi:=xOi0;

//  ShowMessage(form1.kartFam.Value+' '+floattostr(xoi)+#13+floattostr(xOi0));

  ZaprosOINew:=xOi;


end;



procedure TForm123.FormShow(Sender: TObject);
 var xIdp:Real;
     i,npp:integer;
     sSoob:String;
     F:TextFile;
     s:String;
     st:real;

begin

  _TPKOD:=-1;


    if RMes<=11 then  N8.Caption:='Выплата з/п за '+namemes[RMEs]+' Вычеты будущего периода '+namemes[RMes+1]+' -->> '+namemes[RMes];
    if RMEs>=2 then N10.Caption:='Выплата з/п за '+namemes[RMEs-1]+' Вычеты текущего периода '+namemes[RMes]+' -->> '+namemes[RMes-1];
    if RMes=12 then N8.Visible:=false else N8.Visible:=true;
    if RMes=1 then N10.Visible:=false else N10.Visible:=true;


  form102.RxLabel1.Caption:='Предварительная проверка'   ;
  form102.ProgressBar1.Position:=0;
  form102.Show;
  form102.Refresh;
  npp:=0;
  form1.kart.first;
   while not form1.kart.eof do
    begin
       npp:=npp+1;
       form102.ProgressBar1.Position:=TRUNC(100*npp/form1.kart.RecordCount);
       form102.ProgressBar1.Refresh;
       form102.ProgressBar1.Repaint;

       if form1.kartSTATUS.Value='2' then st:=30 else st:=13;

         if (not form1.Proverka2023(form1.kartnls.value,RMes,st)) then
          begin
           //ShowMessage(form1.kartfam.value+#13+floattostr(RMes));
           form1.RaspredVicet;
          end;




     form1.kart.next;
    end;
  form102.close;

 TUpd:=false;
 TUpdNls:=-1;

 //проверить
 Form_58.DelIdVyplNull(0);

   form85:=tform85.create(nil);
   form85.DelErrOklad;//удаляет ошибочные записи sdoxod для которых doxod=0 and nalog<>0
   form85.free;


 Label6.Caption:='[TRA='+inttostr(form1.TREJIMAVTO)+']';

 ComboBox1.Items.Clear;
 ComboBox1.Items.Add('без учета статьи');
 TXS:=-1;
 form1.statya.Active:=true;
 form1.statya.First;
 while not form1.statya.eof do
  begin
   ComboBox1.Items.Add(form1.statyaname.value);
   form1.statya.next;
  end;
 ComboBox1.Text:=ComboBox1.Items[0];

 if TYPEVYPLAT=2 then ComboBox1.Enabled:=false else ComboBox1.Enabled:=true;
 if TYPEVYPLAT=2 then JvXpButton268.Enabled:=false else JvXpButton268.Enabled:=true;

 if form123.TYPEVYPLAT=2 then Label5.Visible:=True else Label5.Visible:=False;
 if form123.TYPEVYPLAT=2 then JvXpButton222.Visible:=True else JvXpButton222.Visible:=False;

 if form123.TYPEVYPLAT=2 then JvXpButton20.Visible:=false else JvXpButton20.Visible:=true;

 if (form123.TYPEVYPLAT=2) and (RGod>=2022) then CheckBox199.Visible:=True else CheckBox199.Visible:=False;
 if (form123.TYPEVYPLAT=2) and (RGod>=2022) then jvXPButton201.Visible:=false else jvXPButton201.Visible:=True;
// if RMes=1 then CheckBox201.Visible:=false;


 if form123.TYPEVYPLAT=2 then
  begin
   Radiobutton2.Checked:=true;
   RadioGroup1.Enabled:=false;
   Radiobutton1.Enabled:=false;
   Radiobutton2.Enabled:=false;
   Radiobutton3.Enabled:=false;
   Radiobutton4.Enabled:=false;
  end
    else
  begin
   Radiobutton2.Checked:=true;
   RadioGroup1.Enabled:=true;
   RadioButton1.Enabled:=true;
   RadioButton2.Enabled:=true;
   RadioButton3.Enabled:=true;
   RadioButton4.Enabled:=true;
  end;
// if TYPEVYPLAT=2 then JvXpButton59.Visible:=True else  JvXpButton59.Visible:=False;



 for i:=0 to 10 do form203.llist[i]:=true;
 for i:=0 to 10 do form203.zlist[i]:=0;

 {MessageDlg('Порядок работы:'+#13+
  '- проверьте суммы к выплате (при необходимости откорректируйте через изменение, удаление)'+#13+
             '- для проведения выплаты нажать на кнопку ПРОВЕСТИ '+#13+
             '- синим цветом помечаются проведенные операции',
              mtInformation,[mbOk],0  );
 }

 xD1:=EncodeDate(RGod,RMes,1);
 if RMes<>12 then xD2:=EncodeDate(RGod,RMes+1,1)-1;
 if RMes=12 then xD2:=EncodeDate(RGod,12,31);




 xKodUderj:=0;
 DBGrid1.Columns.Clear;
 DBGrid1.Columns.Add;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[0].FieldName:='G';
 DBGrid1.Columns[0].Title.Caption:=' ';
 DBGrid1.Columns[0].Width:=15;


 DBGrid1.Columns[1].FieldName:='FAM';
 DBGrid1.Columns[1].Title.Caption:='Фамилия';
 DBGrid1.Columns[1].Width:=100;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[2].FieldName:='SUMMA';
 DBGrid1.Columns[2].Title.Caption:='Сумма';
 DBGrid1.Columns[2].Width:=75;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[3].FieldName:='AVANS';
 DBGrid1.Columns[3].Title.Caption:='в т.ч. аванс';
 DBGrid1.Columns[3].Width:=75;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[4].FieldName:='NLS';
 DBGrid1.Columns[4].Title.Caption:='Итого';
 DBGrid1.Columns[4].Width:=70;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[5].FieldName:='KOD';
 DBGrid1.Columns[5].Title.Caption:='Место выплаты';
 DBGrid1.Columns[5].Width:=115;
 DBGrid1.Columns[5].Alignment:=taLeftJustify  ;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[6].FieldName:='NUMDOK';
 DBGrid1.Columns[6].Title.Caption:='Документ';
 DBGrid1.Columns[6].Width:=110;

 DBGrid1.Columns.Add;
 DBGrid1.Columns[7].FieldName:='NOTE';
 DBGrid1.Columns[7].Title.Caption:='Примеч';
 DBGrid1.Columns[7].Width:=300;


 {
 DBGrid1.Columns.Add;
 DBGrid1.Columns[7].FieldName:='KODNAC';
 DBGrid1.Columns[7].Title.Caption:='Код нач.';
 DBGrid1.Columns[7].Width:=60;
 DBGrid1.Columns.Add;
 DBGrid1.Columns[8].FieldName:='SDOXOD';
 DBGrid1.Columns[8].Title.Caption:='Доход';
 DBGrid1.Columns[8].Width:=60;
 DBGrid1.Columns.Add;
 DBGrid1.Columns[9].FieldName:='SNALOG';
 DBGrid1.Columns[9].Title.Caption:='Налог';
 DBGrid1.Columns[9].Width:=60;
 DBGrid1.Columns.Add;
 DBGrid1.Columns[10].FieldName:='MES';
 DBGrid1.Columns[10].Title.Caption:='Месяц';
 DBGrid1.Columns[10].Width:=60;
 DBGrid1.Columns.Add;
 DBGrid1.Columns[11].FieldName:='GOD';
 DBGrid1.Columns[11].Title.Caption:='Год';
 DBGrid1.Columns[11].Width:=60;
 DBGrid1.Columns.Add;
 DBGrid1.Columns[12].FieldName:='DAT';
 DBGrid1.Columns[12].Title.Caption:='Дата';
 DBGrid1.Columns[12].Width:=60;
 }



 


 if TYPEVYPLAT=2 then
   begin
    DBGrid1.Columns[3].Visible:=false;
    DBGrid1.Columns[4].Visible:=false;

    if RGOD>=2022 then
     begin
      DBGrid1.Columns.Add;
      DBGrid1.Columns[8].FieldName:='avdoxod';
      DBGrid1.Columns[8].Title.Caption:='Доход';
      DBGrid1.Columns[8].Width:=65;
      DBGrid1.Columns.Add;
      DBGrid1.Columns[9].FieldName:='avnalog';
      DBGrid1.Columns[9].Title.Caption:='НДФЛ';
      DBGrid1.Columns[9].Width:=50;
      DBGrid1.Columns.Add;
      DBGrid1.Columns[10].FieldName:='avvicet';
      DBGrid1.Columns[10].Title.Caption:='Вычет';
      DBGrid1.Columns[10].Width:=50;
      DBGrid1.Columns.Add;
      DBGrid1.Columns[11].FieldName:='PKOD';
      DBGrid1.Columns[11].Title.Caption:='Подразделение';
      DBGrid1.Columns[11].Width:=140;
      DBGrid1.Columns[11].Alignment:=taLeftJustify  ;
     end;

   { DBGrid1.Columns[7].Visible:=false;
    DBGrid1.Columns[8].Visible:=false;
    DBGrid1.Columns[9].Visible:=false;
    DBGrid1.Columns[10].Visible:=false;
    DBGrid1.Columns[11].Visible:=false;
    DBGrid1.Columns[12].Visible:=false;
   }
   end
    else
   begin

      DBGrid1.Columns.Add;
      DBGrid1.Columns[8].FieldName:='avvicet';
      DBGrid1.Columns[8].Title.Caption:='в.б.п.';
      DBGrid1.Columns[8].Width:=50;
      DBGrid1.Columns.Add;
      DBGrid1.Columns[9].FieldName:='PKOD';
      DBGrid1.Columns[9].Title.Caption:='Подразделение';
      DBGrid1.Columns[9].Width:=140;
      DBGrid1.Columns[9].Alignment:=taLeftJustify
    {
    DBGrid1.Columns[7].Visible:=false;
    DBGrid1.Columns[8].Visible:=false;
    DBGrid1.Columns[9].Visible:=false;
    DBGrid1.Columns[10].Visible:=false;
    DBGrid1.Columns[11].Visible:=false;
    DBGrid1.Columns[12].Visible:=false;
   }
   end;


 Query1.DataBAseName:=form1.DBDIR;
 form93.Query1.DataBAseName:=form1.DBDIR;
 {\\\}
 form123.Query1.DatabaseName:=form1.DBDIR;
 form123.query1.sql.clear;
 form123.query1.SQL.add('select max(idp) from vyplzp');
 form123.Query1.Prepare;
 form123.Query1.Open;
 xIdp:=form123.Query1.Fields[0].AsFloat;
 vyplzp.First;
 while not vyplzp.Eof do
  begin
    if vyplzpTYPE.Value<>2 then
     begin
      vyplzp.Edit;
      vyplzp.fieldByName('TYPE').asFloat:=0;
      vyplzp.Post;
     end;
    if vyplzpIDP.Value=0 then
     begin
      xIdp:=xIdp+1;
      vyplzp.Edit;
      vyplzp.fieldByName('idp').asFloat:=xIdp;
      vyplzp.Post;
     end;
    if vyplzpPROVEDEN.Value=0 then
     begin
      vyplzp.Edit;
      vyplzp.FieldByName('proveden').asInteger:=0; {для пустых значений}
      vyplzp.Post;
     end;
   vyplzp.Next;
  end;

{
 xTYPE:=3;
 ZaprosV2;
}

  fDat:=Date();



  if TYPEVYPLAT=2 then      {АВАНС}
    begin
     JvXPButton2Click(nil);
     fName:='заработная плата за первую половину '+ansilowerCase(Fnamemes(RMes));

     if (CheckBox199.Visible) and (FileExists(form1.DBDIR+'\uavans.txt')) and (RGod>=2022) then
      begin
       AssignFile(F,form1.DBDIR+'\uavans.txt');
       Reset(F);
       ReadLn(F,s);
       if s='UAVANS=TRUE' then CheckBox199.Checked:=true;
       if s='UAVANS=FALSE' then CheckBox199.Checked:=false;
       CloseFile(F);



  end;


    end;





  if TYPEVYPLAT=0 then      {ВЫПЛАТА}
    begin
     JvXPButton6Click(nil);
     fName:='заработная плата за '+ansilowerCase(namemes[RMes]);
     if TXS>=0 then fName:=fName+'/'+ComboBox1.Text+'/';
    end;





 if  TYPEVYPLAT=2 then form123.Caption:='Выплата аванса отчетный период '+ansilowerCase(namemes[RMes])
                      else form123.Caption:='Выплата заработной платы отчетный период '+ansilowerCase(namemes[RMes]);
 if TYPEVYPLAT=2 then form123.JvXPButton40.Visible:=False else form123.JvXPButton40.Visible:=True;
 if TYPEVYPLAT=2 then form123.JvXPButton41.Visible:=False else form123.JvXPButton41.Visible:=True;
 if TYPEVYPLAT=2 then form123.JvXPButton44.Visible:=False else form123.JvXPButton44.Visible:=True;

 if TYPEVYPLAT=2 then form123.N7.Enabled:=False else form123.N7.Enabled:=True;

 datam.Query2.DatabaseName:=form1.DBDIR;


 form1.PFormLoad('form123',Form123);

 form123.DeleteDoxodNac;


 Form123.Button1Click(nil);//проверка недействующих ссылок


 datam.qtmpstaj.close;
 datam.qtmpstaj.DatabaseName:=form1.DBDIR;
 datam.qtmpstaj.sql.Clear;
 datam.qtmpstaj.sql.add('select count(o.id) from obrt2new o where o.wg='+Floattostr(RGod));
 datam.qtmpstaj.sql.add('and o.wm='+floattostr(RMEs));
 datam.qtmpstaj.sql.add('and o.tavans=1 ') ;
 datam.qtmpstaj.prepare;
 datam.qtmpstaj.open;
 
 sSoob:='';
 if datam.qtmpstaj.Fields[0].asFloat>0 then sSoob:='Имеются необработанные записи Аванса'+#13+
    'Аванс необходимо обработать перед итоговой выплатой заработной платы после проведения всех начислений за отчетный период. Кнопка <АВАНС>';

// if TYPEVYPLAT=0 then sSoob:=sSoob+#13+'Для проведения выборочных выплат (отпускные, больничные ....) используйте кнопку <ФИЛЬТР>';

 if trim(sSoob)<>'' then MessageDlg(sSoob,mtWarning,[mbOk],0);

 datam.qtmpstaj.close;

 TVYBOR:=true;

 if TYPEVYPLAT=0 then
   begin


     form3045:=tform3045.create(nil);
      form3045.ShowModal;
      if form3045.JvRadioGroup1.ItemIndex=1 then
       begin
        TVYBOR:=false; //галочки в списке фильтр
        N7Click(nil);
       end;
      form3045.free;
   
   end;

 if RGOD>=2023 then //уплату НДФЛ убрал отсюда с 2023
  begin
   jvXpButton41.Visible:=false;
   jvXpButton258.Visible:=false;
  end;


end;


procedure TForm123.ZaprosNOAvans;
begin
 datam.qtmpstaj.close;
 datam.qtmpstaj.DatabaseName:=form1.DBDIR;
 datam.qtmpstaj.sql.Clear;
 datam.qtmpstaj.sql.add('select sum(o.kr) from obrt2new o where o.nls in (select k.nls from kart k) and o.wg='+Floattostr(RGod));
 datam.qtmpstaj.sql.add('and o.wm='+floattostr(RMEs));
 datam.qtmpstaj.sql.add('and o.tavans=1') ;
 datam.qtmpstaj.Prepare;
 datam.qtmpstaj.Open;
 CurrencyEdit4.Value:=DRound(datam.qtmpstaj.Fields[0].asFloat,2);
 datam.qtmpstaj.close;
end;


procedure TForm123.Zaprosv2(mk:integer);
var xOi,xAvans:Real;
    s:String;
begin

 ZaprosNOAvans;

 datam.qtmpstaj.close;
 datam.qtmpstaj.DatabaseName:=form1.DBDIR;
 datam.qtmpstaj.sql.Clear;
 datam.qtmpstaj.sql.add('select o.nls, sum(o.kr) from obrt2new o where o.wg='+Floattostr(RGod));
 datam.qtmpstaj.sql.add('and o.wm='+floattostr(RMEs));
 datam.qtmpstaj.sql.add('and o.tavans=1') ;
 datam.qtmpstaj.sql.add('group by nls');
 datam.qtmpstaj.Prepare;
 datam.qtmpstaj.Open;




 vyplzp.first;
 while not vyplzp.eof do
  begin
   xAvans:=0;

   if datam.qtmpstaj.Locate('nls',vyplzpNls.Value,[loCaseInsensitive]) then xAvans:=datam.qtmpstaj.fields[1].asFloat;

   if vyplzpPROVEDEN.Value<>1 then
    begin
     xOi:=vyplzpSUMMa0.Value;
     if RadioButton1.Checked then xOi:=INT(xOi);
     if RadioButton3.Checked then xOi:=INT(xOi/10)*10;
     if RadioButton4.Checked then xOi:=DRound(xOi,0);

     vyplzp.edit;
     vyplzp.fieldByName('Summa').asFloat:=DRound(xOi,2);
     vyplzp.fieldbyname('avans').asFloat:=xAvans;
   {  vyplzp.fieldbyname('snalog').asFloat:=0;
     vyplzp.fieldbyname('sdoxod').asFloat:=0;
     vyplzp.fieldbyname('god').asFloat:=0;
     vyplzp.fieldbyname('mes').asFloat:=0;
     vyplzp.fieldbyname('kodnac').asFloat:=0;
    }
     vyplzp.post;
    end;
   vyplzp.Next;
  end;
 vyplzp.FlushBuffers;
 Query1.Close;
 qUERY1.sqL.Clear;
 Query1.sql.add('select k.nls, k.fam,k.im,k.ot, k.pkod, o.proveden,o.avans,o.type,o.summa, o.kod, o.wm, o.wg, o.numdok, o.datdok,o.idp, o.note, o.g, o.kodnac, o.sdoxod, snalog, god, mes, dat');
 Query1.sql.Add(',o.avdoxod, o.avnalog, o.avvicet from kart k, vyplzp o' );
 Query1.sql.add('where k.nls=o.nls and o.wm='+IntToStr(RMes)+' and o.wg='+intToStr(RGod));
 if mk=1 then Query1.sql.Add('and o.G='+#39+'*'+#39)  ; // выделены зеленым цеветом +
 if form118.KartFilter<>'' then
 Query1.sql.add(' and '+form118.kartfilter);
 if RadioGroup2.ItemIndex=2 then Query1.sql.Add('and o.proveden<>1');
 if RadioGroup2.ItemIndex=1 then Query1.sql.Add('and o.proveden=1');
{
 if not CheckBox2.Checked then Query1.sql.add(' and o.summa<>0 ');
}


 if (TYPEVYPLAT=0) and (T_OWKOD=true) then Query1.sql.add(' and o.summa<>0 '); // не выводим суммы нулевык если фильтр по кодам выплат
 {
 Query1.sql.add(' and o.proveden<>1');
}
 Query1.sql.add(' and o.type='+IntToStr(XTYPE));

 if _TPKOD>=0 then  Query1.sql.add(' and k.pkod='+IntToStr(_TPKOD));   //подразделение

 Query1.sql.add('order by k.fam, k.im');
 Query1.Prepare;
 Query1.Open;

{ s:='';
 if T_OWKOD then s:='true' else s:='false';
 ShowMessage('T_OWKOD='+s+#13+'TYPEVYPLAT='+floattostr(TYPEVYPLAT));
}


 datam.qtmpstaj.close;


end;

procedure TForm123.ZaprosV1(tParam:integer);
var i:Integer;
    xKod:Integer;
    xIdp,x,y:Real;
    npp:integer;
begin
 Query1.close;
 Query1.DatabaseName:=form1.DBDIR;
 query1.sql.clear;
 query1.SQL.add('select max(idp) from vyplzp');
 Query1.Prepare;
 Query1.Open;
 xIdp:=Query1.Fields[0].AsFloat;
 Query1.Close;

 x:=0;

 if tParam=0 then  for i:=0 to 100 do UWKOD[i]:=false;
 if tParam=0 then T_OWKOD:=false else T_OWKOD:=true;   //признак фильтра для upl1333

  form102.RxLabel1.Caption:='Ожидайте'   ;

  form102.ProgressBar1.Position:=0;
  form102.Show;
  form102.Refresh;
  npp:=0;

form1.kart.first;

 while not form1.kart.eof do
  begin

   
    npp:=npp+1;
    form102.ProgressBar1.Position:=TRUNC(100*npp/form1.kart.RecordCount);
    form102.ProgressBar1.Refresh;
    form102.ProgressBar1.Repaint;

    xKod:=0;
    try
     xKod:=StrToInt(Trim(form1.kartD1C.Value));
    except
     xKod:=0;
    end;

    datam.Query1.Close;
    datam.query1.sql.clear;
    datam.query1.SQL.add('select idp from vyplzp where wm='+FloatToStr(RMEs)+' and wg='+FloatToStr(RGod)+' and type=0 and proveden=0 and nls='+form1.kartNLS.asString);
    datam.Query1.Prepare;
    datam.Query1.Open;
    if datam.query1.RecordCount>1 then
      begin
       {убиваем все двойные записи}
       datam.Query1.Close;
       datam.query1.sql.clear;
       datam.query1.SQL.add('DELETE from vyplzp where wm='+FloatToStr(RMEs)+' and wg='+FloatToStr(RGod)+' and type=0 and proveden=0 and nls='+form1.kartNLS.asString);
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
      end;
    datam.Query1.Close;



    IF form1.TREJIMAVTO=1
          THEN
            begin

             x:=ZaprosOINew(TXS);

             y:=ZaprosOIDoxod(form1.kartNLS.Value,tParam); //остатки по кодам начислений-выплата доход

          //   x:=y; // исправлено 02.10.2020 чтобы аванс не включался только по sdoxod

             if y<x then
                begin
                 {
                 MessageDlg(form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+
                                'Остаток к выплате по налоговым регистрам (дата выплаты) '+floattostrf(y,ffnumber,12,2)+'р.'+#13+' меньше итого остатка к выплате по оборотам '+
                                 floattostrf(x,ffnumber,12,2)+'р.'+#13+'Возможно в будущих месяцах есть ссылка на выплату дохода',mtWarning,[mbOk],0);
                  }
                 x:=y;//меньшую сумму к выплате  т.к. может выплата вперед быть в месяце март за февраль и если вернуться в февраль
                end;
            end
            ELSE x:=ZaprosOINew(TXS);   //просто остатки

    if x<0 then x:=0;



     if not vyplzp.Locate('NLS;WM;WG;TYPE;PROVEDEN',
             VarArrayOf([form1.kartNLS.Value,RMes,RGod,0,0]),[loCAseInsensitive]) then
      begin
       if x>0 then
        begin
         xIdp:=xIdp+1;
         vyplzp.Append;
         vyplzp.FieldByName('idp').asFloat:=xIdp;
         vyplzp.FieldByName('type').asFloat:=0;
         vyplzp.FieldByName('nls').asFloat:=form1.kartNls.Value;
         vyplzp.FieldByName('numdok').asFloat:=0;
         vyplzp.FieldByName('wm').asFloat:=RMEs;
         vyplzp.FieldByName('wg').asFloat:=RGod;
         vyplzp.FieldByName('summa0').asFloat:=x;
         vyplzp.FieldByName('kod').asInteger:=xKod;
         vyplzp.FieldByName('proveden').asInteger:=0;
         vyplzp.post;
        end;
      end
       else
      begin
       if vyplzpPROVEDEN.Value=0 then
        begin
         vyplzp.Edit;
         vyplzp.Fieldbyname('summa0').asFloat:=x;
         vyplzp.post;
        end;
      end;


      
   form1.kart.next;
  end;
 form102.RxLabel1.Caption:='Формирую отчет, ожидайте'   ;
 form102.close;




end;

procedure TForm123.RadioButton1Click(Sender: TObject);
begin
 ZaprosV2(0);
end;

procedure TForm123.RadioButton2Click(Sender: TObject);
begin
 ZaprosV2(0);
end;

procedure TForm123.RadioButton3Click(Sender: TObject);
begin
 ZaprosV2(0);
end;

procedure TForm123.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var s1:String;
    nDok:Integer;
    x:Real;
begin
 s1:=Column.Field.Text;
  DBGrid1.Canvas.Brush.Color:=$00F4F4F4;
  DBGrid1.Canvas.Font.Style:=[];
  DBGrid1.Canvas.Fillrect(Rect);
  DBGrid1.Canvas.Font.Color:=clBlack;
  DBGrid1.Canvas.Font.Name:='Calibri';
  DBGrid1.Canvas.Font.Size:=9;


{  if Column.FieldName = 'G' then // CheckInDBGrid
    if Query1.FieldByName('G').asString='*' then
      DrawGridCheckBox(DBGrid1.Canvas, Rect, true)
    else
      DrawGridCheckBox(DBGrid1.Canvas, Rect, false);
}

  if Query1.fieldByName('PROVEDEN').asInteger=1 then DBGrid1.Canvas.Font.Color:=clBlack else DBGrid1.Canvas.Font.Color:=clRed ;

   if (AnsiUpperCase(Column.FieldNAme)='SUMMA') then s1:=FloatToStrF(Query1.FieldByName('SUMMA').asFloat,ffNumber,12,2);
   if (AnsiUpperCase(Column.FieldNAme)='AVDOXOD') then s1:=FloatToStrF(Query1.FieldByName('AVDOXOD').asFloat,ffNumber,12,2);


  if (AnsiUpperCase(Column.FieldNAme)='AVANS') then
     begin
      // DBGrid1.Canvas.Font.Style:=[fsItalic];

      if (Query1.FieldByName('proveden').asFloat=1) or (DRound(Query1.FieldByName('AVANS').asFloat,2)=0) then s1:=''
         else
          begin
            s1:=FloatToStrF(Query1.FieldByName('AVANS').asFloat,ffNumber,12,2);
            DBGrid1.Canvas.Font.Color:=clRed;
          end;
     end;

   if (AnsiUpperCase(Column.FieldNAme)='PKOD') then
     begin
      if form1.filial.locate('pkod',query1.fieldbyname('pkod').asfloat,[locaseinsensitive]) then s1:=' '+ansilowercase(form1.filialname.value) else s1:='';
     end;

   if (AnsiUpperCase(Column.FieldNAme)='NLS') then
     begin
      // DBGrid1.Canvas.Font.Style:=[fsItalic];

      x:=Query1.FieldByName('SUMMA').asFloat-Query1.FieldByName('AVANS').asFloat;

      if (Query1.FieldByName('proveden').asFloat=1) or (DRound(Query1.FieldByName('AVANS').asFloat,2)=0) then s1:=FloatToStrF(Query1.FieldByName('SUMMA').asFloat,ffNumber,12,2)
         else
          begin
            s1:=FloatToStrF(x,ffNumber,12,2);
            if x<0 then DBGrid1.Canvas.Font.Color:=clRed;
          end;
     end;

   
   
   if ansiuppercase(Column.FieldNAme)='DAT' then
     begin
      s1:='';
      if Query1.FieldByName('DAT').asDateTime>EncodeDate(2000,1,1) then s1:=FormatDateTime('dd.mm.yyyy',Query1.FieldByName('DAT').asDateTime);
     end;


   if (Ansiuppercase(Column.FieldNAme)='G') then if Query1.FieldByName('G').asString='*' then s1:='[+]' else s1:='';


   if (Column.FieldNAme='FAM') or (Column.FieldNAme='fam') then
    begin
     s1:=Query1.fieldByName('Fam').asString+' '+Copy(Query1.fieldByName('Im').asString,1,1)+'.'+
             Copy(Query1.fieldByName('Ot').asString,1,1)+'.';
    end;


  if (Column.FieldNAme='KOD') or (Column.FieldNAme='kod') then
   begin
    if form1.uderj.Locate('KOD',Query1.fieldByName('KOD').asInteger,[loCaseInsensitive]) then
      s1:=' '+AnsiLowerCase(form1.uderjNAME.Value) else s1:=' <не определено>';
   end;

   if (Column.FieldNAme='NOTE') or (Column.FieldNAme='note') then
   begin
    if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]) then
      s1:=vyplzpNote.Value else s1:='';
   end;

  if Query1.fieldByName('G').asString='*' then DBGrid1.Canvas.Font.Color:=clGreen;


  if (Column.FieldNAme='NUMDOK') or (Column.FieldNAme='numdok') then
   begin
    if form1.uderj.Locate('KOD',Query1.fieldByName('KOD').asInteger,[loCaseInsensitive]) then
     begin
      try
       nDok:=StrToInt(Trim(form1.uderjDBSPRAV.asString));
      except
       nDok:=0;
      end;
      if Query1.FieldByName('datdok').asDateTime>EncodeDate(2000,1,1) then
                            s1:=DateToStr(Query1.FieldByName('datdok').asDateTime)
                             else s1:='-';
      if Query1.FieldByName('numdok').asInteger<>0 then
       begin
        s1:='№'+IntToStr(Query1.FieldByName('numdok').asInteger)+' от '+DateToStr(Query1.FieldByName('datdok').asDateTime);
        if nDok=1 then s1:='р/о №'+IntToStr(Query1.FieldByName('numdok').asInteger)+' '+DateToStr(Query1.FieldByName('datdok').asDateTime);
        if nDok=2 then s1:='п/п №'+IntToStr(Query1.FieldByName('numdok').asInteger)+' '+DateToStr(Query1.FieldByName('datdok').asDateTime);
       end;
     end
       else s1:='<>';

   end;

   if Column.Alignment=taRightJustify then
          DBGrid1.Canvas.TextOut(Rect.Right-2-DBGrid1.Canvas.TextWidth(s1),
               Rect.Top+2,s1) else
                     DBGrid1.Canvas.TextOut(Rect.Left+2,Rect.Top+2,s1);


end;

Procedure TForm123.PlatVed2;
var i:Integer;
    F:textFile;
    x,x0:Real;
    isxrtf,outrtf:String;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
begin
 form93.Query1.DatabaseName:=form1.DBDIR;
 isxrtf:=GetCurrentDir()+'\ISXRTF\platVed2.rtf';
 outrtf:=GetCurrentDir()+'\OUTRTF\_platVed2.rtf';

 AssignFile(F,'dat.txt');
 Rewrite(F);
 Query1.first;
  i:=0;
  x0:=0;

  while not Query1.eof do
   begin
    rtf:=false;
    x:=Query1.FieldByName('SUMMA').asFloat ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=2 then rtf:=True;
     end;

  { if Query1.FieldByName('datdok').asDAteTime<>xD then rtf:=False;
  }
    if Query1.fieldByName('PROVEDEN').asInteger<>1 then rtf:=False;


    IF (x<>0) and (rtf)  then
     begin
     x0:=x0+x;
     i:=i+1;
     if i=1 then
      begin
       Write(F,'|9:NAME:'+form1.config2NAME.Value);
       Write(F,'|9:DAT:'+FormatDateTime('dd.mm.yyyy',Date()));
       Write(F,'|9:PERIOD:'+namemes[RMEs]+' '+IntToStr(RGod)+'г.');
       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       WriteLN(F,'|9:RUKOV:'+xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.');
      end;

       form1.kart.Locate('nls',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]);

      write(F,'|8:NPP:'+IntToStr(i));
      write(F,'|8:FIO:'+form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value);
      write(F,'|8:SUMMA:'+FloatToStrF(x,ffNumber,12,2));
      if Query1.fieldByName('numdok').asFloat<>0 then
         writeLn(F,'|8:REKVIZIT:'+'№'+Query1.fieldByName('numdok').asString+', '+
          DateToStr(Query1.fieldByName('datdok').asDateTime))
           else writeLn(F,'|8:REKVIZIT:');




     END;
    Query1.next;
   end;
  Write(F,'|5:SUMMA:'+FloatToStrF(x0,ffNumber,12,2));
  mainLib.GetFIO(form1.configGLBUH.Value,xFam,xIm,xOt);
  WriteLN(F,'|5:GLBUH:'+xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.');

 CloseFile(F);

if i<>0 then
  begin
   reportFM.rtf_CreateReport('dat.txt',isxrtf,outrtf,nil,i);
   form20.StartWord(outrtf);
  end
   else MessageDlg('Ничего не сформировано, проверьте:'+#13+
            '- нет проведенных оборотов (сначала необходимо нажать на Провести выплату)'+#13+
            '- нет операций с признаком пл/пор (все операции по кассе)'+#13+
           { '- проверьте параметр Дата выплаты'+#13+
           } '- в Справочнике удержаний (выплат) неверно установлены флаги документов по выплате через Банк',mtInformation,[mbOk],0);

end;

procedure TForm123.vyplzpsumpropisGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
var s:String;
    i:Integer;
    rubl,kopey:Real;
begin
 rubl:= Int(vyplzpSUMMA.Value);
 kopey:=Round((Frac(vyplzpSUMMA.Value)*100));
 s:=NumTools.NumeralToPhrase(FloatToStr(rubl));
 s:=s+' '+NumTools.GeniCase(FloatToStr(rubl),'рубль','рубля','рублей');
 if kopey<10 then s:=s+' 0'  else s:=s+' ';
 s:=s+FloatToStr(kopey);
 s:=s+' '+NumTools.GeniCase(FloatToStr(kopey),'копейка','копейки','копеек');
 Text:=s;
end;

procedure TForm123.vyplzpssummaGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
var s:String;
    i:Integer;
begin
 Str(vyplzpSUMMA.Value:12:2,s);
 s:=Trim(s);
 for i:=1 to Length(s) do
  begin
   if (s[i]=',') or (s[i]='.') then s[i]:='-';
  end;
 Text:=s;
end;


procedure TForm123.Favans(sDoxod,sVicet:real;var sAvans,sNdfl:Real);
var sBasa,st:Real;
begin
  sBasa:=DRound(sDoxod-sVicet,2);
  if sBasa<0 then sBasa:=0;
  if form1.kartSTATUS.Value='2' then st:=30 else st:=13;
  sNdfl:=DRound(sBasa*st/100,0);
  sAvans:=Dround(sDoxod-sNdfl,2);
end;

procedure TForm123.ZaprosV3;
var i,k,m,ndotp:Integer;
    xKod:Integer;
    xIdp,x,y:Real;
    xProc:Real;
    rtf1,rtf2,rtf5,rtf6,_rtfav:Boolean;
    rtfotp:boolean;
    tmpd1,tmpd2:TDate;
    xDatP:TDAte;
    dd,mm,yy:Word  ;
    kNadbavka:Real;
    nn:Integer;
    xOklad,xdayRab,xDayOtr:Real;
    xVicet,xVicet0:Real;
    skodv:String;
    xcas1,xcas2:Real;
    XTEKTYPEOKLAD:Real;
    TKarantin:Boolean;
    sDoxod,sVicet,sBasa,sNdfl,st,rvicet:Real;
    oldtavans,newtavans:Real;
    ik:integer;
begin

 vavans.DatabaseName:=form1.DBDIR;
 vavans.Active:=false;
 DeleteFile(form1.DBDIR+'\vavans.dbf');
 DeleteFile(form1.DBDIR+'\vavans.mdx');
 if not vavans.Exists then vavans.CreateTable;
 vavans.Active:=true;
 Reindex.ReindexTab(vavans,form1.DBDIR+'\vavans','fio','');
 vavans.IndexName:='fio';


 Query1.Close;
 Query1.DatabaseName:=form1.DBDIR;
 query1.sql.clear;
 query1.SQL.add('select max(idp) from vyplzp');
 Query1.Prepare;
 Query1.Open;
 xIdp:=Query1.Fields[0].AsFloat;
 Query1.Close;

  xProc:=50;
//  MessageDlg('Сумма аванса берется из карточки сотрудника,'+#13+'в случае если в карточке не установлено - берется 50% от оклада',mtInformation,[mbOk],0);

 // MessageDlg('Сумма аванса берется из карточки сотрудника',mtInformation,[mbOk],0);



 form1.staj.MasterFields:='NLS';
 form1.staj.MasterSource:=form1.DataSource1;


 form1.kart.first;

 while not form1.kart.eof do
  begin
   // Showmessage('Вход:'+form1.kartfam.value+#13+floattostr(form1.kartnls.value));

    form1.GetPravo(RMes);


    xKod:=0;
    try
     xKod:=StrToInt(Trim(form1.kartD1C2.Value));
    except
     xKod:=0;
    end;


    {ПРОВЕРКА УВОЛЕННЫХ !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}


     rtf6:=false;
     if (form1.staj.RecordCount<=0) and (DRound(form1.kartAVANS.Value,2)<>0) then   //для арендодателей - у них оплата по договору без приема на работу
         begin
           if MessageDlg(form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value+#13+
              'Невозможно определить период работы сотрудника (нет сведений о приеме на работу)'+#13+
                'Включать в аванс ?' ,mtWarning,[mbYes,mbNo],0) = mrYes  then rtf6:=true;
          end;

       form1.staj.first;
       rtf1:=False;
       rtf2:=false;
       rtf5:=false;  //rtf5 - для устроенных в этом месяце !
       k:=0; xDatP:=date();
       tmpd1:=EncodeDAte(RGod,RMes,1);
       if RMes<>12 then tmpd2:=EncodeDAte(RGod,RMes+1,1)-1
                   else tmpd2:=EncodeDAte(RGod,12,31) ; {аванс - если внутри месяца сотрудник - числится}
       kNadbavka:=0;
       xDayRab:=0; xDayOtr:=0;
       while not form1.staj.eof do
        begin
           if (tmpd1>=form1.stajD1.Value)
                  and (form1.stajD2.Value<=EncodeDate(1910,1,1))
                      then kNadbavka:=form1.stajPNADBAVKA.Value+form1.stajPNADBAV2.Value+form1.stajPNADBAV3.Value;
            if (tmpd1>=form1.stajD1.Value) and (tmpd1<=form1.stajD2.Value)
                       then kNadbavka:=form1.stajPNADBAVKA.Value+form1.stajPNADBAV2.Value+form1.stajPNADBAV3.Value;



          if form1.stajTYPE.Value<>2 then
           begin
            k:=k+1;
            if (form1.stajD1.Value>=tmpd1) and (form1.stajD1.Value<=tmpd2) and (k=1) then
               begin
                rtf5:=true;
                xDatP:=form1.stajD1.Value;
               end;
            if (tmpd1>=form1.stajD1.Value)
                  and (form1.stajD2.Value<=EncodeDate(1910,1,1)) then rtf1:=true;
            if (tmpd1>=form1.stajD1.Value) and (tmpd1<=form1.stajD2.Value) then rtf1:=true;

            if (tmpd2>=form1.stajD1.Value)
                  and (form1.stajD2.Value<=EncodeDate(1910,1,1)) then rtf2:=true;
            if (tmpd2>=form1.stajD1.Value) and (tmpd2<=form1.stajD2.Value) then rtf2:=true;



           end;

          form1.staj.Next;
        end;

    //  Showmessage('Выход:'+form1.kartfam.value+#13+floattostr(form1.kartnls.value));



    {КОНЕЦ ПРОВЕРКИ УВОЛЕННЫХ !!!!!!!!!!!!!!!!!!!!!!!!!!}

    datam.kart2.Locate('nls',form1.kartnls.value,[loCaseinsensitive]);

    if (datam.kart2TYPEAVANS.Value=3) or (datam.kart2TYPEAVANS.Value=5) or (datam.kart2TYPEAVANS.Value=12) or (datam.kart2TYPEAVANS.Value=8)
        or (datam.kart2TYPEAVANS.Value=9) or (datam.kart2TYPEAVANS.Value=7) or (datam.kart2TYPEAVANS.Value=10) then
     begin
      oldtavans:=datam.kart2.fieldbyname('typeavans').asfloat;
      datam.kart2.edit;
      if datam.kart2TYPEAVANS.Value=3 then datam.kart2.fieldbyname('typeavans').asfloat:=2;
      if datam.kart2TYPEAVANS.Value=5 then datam.kart2.fieldbyname('typeavans').asfloat:=4;
      if datam.kart2TYPEAVANS.Value=12 then datam.kart2.fieldbyname('typeavans').asfloat:=11;
      if datam.kart2TYPEAVANS.Value=8 then datam.kart2.fieldbyname('typeavans').asfloat:=2;
      if datam.kart2TYPEAVANS.Value=7 then datam.kart2.fieldbyname('typeavans').asfloat:=6;
      if datam.kart2TYPEAVANS.Value=9 then datam.kart2.fieldbyname('typeavans').asfloat:=4;
       if datam.kart2TYPEAVANS.Value=10 then datam.kart2.fieldbyname('typeavans').asfloat:=4;
      datam.kart2.post;
      newtavans:=datam.kart2.fieldbyname('typeavans').asfloat;
      MessageDlg(form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+'Изменен вид расчета аванса'+#13+
       datam.DTAVANS[trunc(oldtavans)]+#13+
         'на новый: '+#13+datam.DTAVANS[trunc(newtavans)]+#13+'С 2023 года аванс рассчитывается сразу с вычетами и с НДФЛ'
           ,mtinformation,[mbOk],0);
     end;

    xOklad:=mainlib.LoadOklad(RMEs,RGod,form1.kartNLS.Value);

    XTEKTYPEOKLAD:=LoadTypeOklad(RMes,RGod,form1.kartNLS.Value);


    if (rtf1 and rtf2) or (rtf5) or (rtf6) then
      begin
       datam.kart2.Locate('nls',form1.kartnls.value,[loCaseinsensitive]);
       if datam.kart2TYPEAVANS.Value=1 then
        begin
          x:=xOklad;
          x:=DRound(x*form1.kartAVANS.Value*tkUc/100,2);
         // ShowMessage(form1.kartfam.value+#13+floattostr(tKUc));
        end
         else
           x:=form1.kartAVANS.Value;

       // if (form1.kartAVANS.Value<>0) then x:=form1.kartAVANS.Value
       //    else
       //      x:=DRound(form1.kartOKLAD.Value*xProc/100,2);
      end
       else
        begin
         x:=0;
        end;

    if Form1.kartDAYCAS.Value then x:=0;
    DecodeDate(xDatP,yy,mm,dd);
    if (rtf5) and (dd>15) then x:=0;
    if (rtf5) and (x<>0)  then
     begin
      x:=DRound(x*(15-dd+1)/15,2);     //пропорция
     end;

       {проверка отпускников}
    tmpd1:=EncodeDAte(RGod,RMes,1);
    tmpd2:=EncodeDAte(RGod,RMes,15);
    datam.query1.close;
    datam.query1.sql.clear;
    datam.query1.sql.add('select * from tabel where nls='+floattostr(form1.kartnls.value));
  //  datam.query1.sql.add('and kod=14');
    datam.query1.sql.add('and kod in(14,15,18,19,21,24,25,26,29,30,31,35,48,39)');
    datam.query1.sql.add('and dat>='+#39+formatdatetime('dd.mm.yyyy',tmpd1)+#39);
    datam.query1.sql.add('and dat<='+#39+formatdatetime('dd.mm.yyyy',tmpd2)+#39);
    datam.query1.Prepare;
    datam.Query1.open;
    datam.query1.First;
    ndotp:=0;
    rtfotp:=false;
    while not datam.query1.eof do
     begin
      ndotp:=ndotp+1;
      datam.query1.next;
     end;
     if ndotp<>0 then
      begin
       //пропорция
       x:=DRound(x*(15-ndotp)/15,2);
       rtfotp:=true;
      end;
    datam.query1.close;

        form1.idkalend.Locate('ID',form1.kartIDKALEND.Value,[loCaseInsensitive]);
        form1.kalend1.Locate('GOD;MES;TYPE',VarArrayOf([RGod,RMes,form1.kartIDKALEND.Value]),[loCaseInsensitive]);
        //факт дни отраб с 01 по 5 число
            {проверка отпускников}

        m:=0;  xcas1:=0;    xcas2:=0;
        for i:=1 to 15 do
         begin
          tmpd1:=EncodeDAte(RGod,RMes,i);
          xcas2:=sumhhmm(xcas2,form1.kalend1.Fieldbyname('N'+IntToStr(i)).asFloat);
          datam.query1.close;
          datam.query1.sql.clear;
          datam.query1.sql.add('select * from tabel where nls='+floattostr(form1.kartnls.value));
          datam.query1.sql.add('and cas>0 and dat='+#39+formatdatetime('dd.mm.yyyy',tmpd1)+#39);
          datam.query1.Prepare;
          datam.Query1.open;
          if datam.Query1.RecordCount>0 then
           begin
            m:=m+1;
            xcas1:=sumhhmm(xcas1,datam.query1.fieldbyname('cas').asfloat);
           end;
         end;
        datam.Query1.close;

    k:=form1.kalend1DAYRAB.Value;
    xcas2:=form1.kalend1HOURS.Value;

    //****
    TKarantin:=FKarantin();
    if (TKarantin) and (form1.kalend1GOD.Value=2020) then
     begin
      if form1.kalend1MES.Value=3 then
       begin
        k:=21;
        if (form1.kalend1TYPE.Value=4) or (form1.kalend1TYPE.Value=5) then xcas2:=168;
        if (form1.kalend1TYPE.Value=3) or (form1.kalend1TYPE.Value=7) then xcas2:=151.12;
       end;
      if form1.kalend1MES.Value=4 then
       begin
        k:=22;
        if (form1.kalend1TYPE.Value=4) or (form1.kalend1TYPE.Value=5) then xcas2:=175;
        if (form1.kalend1TYPE.Value=3) or (form1.kalend1TYPE.Value=7) then xcas2:=157.24;
       end;
      if form1.kalend1MES.Value=5 then
       begin
        k:=17;
        if (form1.kalend1TYPE.Value=4) or (form1.kalend1TYPE.Value=5) then xcas2:=135;
        if (form1.kalend1TYPE.Value=3) or (form1.kalend1TYPE.Value=7) then xcas2:=121.24;
       end;
     end;
     //****

    if (XTEKTYPEOKLAD=1)  then  //оклад по часам
      begin
       xDAyRab:=VremyToFloat(xcas2);
       xDayOtr:=VremyToFloat(xcas1);
      end
       else
     if (XTEKTYPEOKLAD=0) then //оклад по дням
      begin
       xDayRab:=k;
       xDayOtr:=m;
      end;
     if (XTEKTYPEOKLAD=3)  then  //часовая ставка
      begin
       xDAyRab:=1;
       xDayOtr:=VremyToFloat(xcas1);
      end
       else
     if (XTEKTYPEOKLAD=2) then //дневная ставка
      begin
       xDayRab:=1;
       xDayOtr:=m;
      end;


   {}
     k:=Trunc(xDayRab);
     m:=Trunc(xDayOtr);
   {}


    xVicet:=0;
    xVicet0:=0;

    if (datam.kart2TYPEAVANS.VAlue>=2) and (datam.kart2TYPEAVANS.VAlue<=10) and (k<>0) then      //факт. дни
      begin

        xVicet:=0;
        if (datam.kart2typeavans.Value>=8) and (datam.kart2typeavans.Value<=10) then
          begin
           xVicet:=FVic1(RMes)+FKartVic2(RMes,skodv)+FIjdev2(RMes,skodv);
           // Showmessage(floattostr(xVicet)+#13+form1.kartfam.value);
          end;
        xVicet0:=xVicet;
        
        x:=0;
        if (datam.kart2typeavans.Value=2) or (datam.kart2typeavans.Value=3) or (datam.kart2typeavans.Value=8) then
         begin
         // ShowMessage(form1.kartfam.value+#13+floattostr(form1.kartOKLAD.value)+#13+floattostr(xDayOtr)+#13+floattostr(xDayRab));
          if xDayRab<>0 then x:=DRound(form1.kartOKLAD.Value*xDayOtr/xDayRab,2);
          x:=DRound(x*(1+form1.configRK.Value/100),2);
          if datam.kart2typeavans.Value=3 then x:=DRound(x*0.87,0);
          xVicet:=x-xVicet;  //база
          if xVicet<0 then xVicet:=0;
          if datam.kart2typeavans.Value=8 then x:=DRound(x-xVicet*0.13,0) ;
         end;


        if (datam.kart2typeavans.Value=4) or (datam.kart2typeavans.Value=5) or (datam.kart2typeavans.Value=9) then
         begin

       // showmessage(form1.kartFAM.value+#13+floattostr(form1.kartOKLAD.Value)+#13+floattostr(kNadbavka)+#13+floattostr(m)+#13+floattostr(k)+#13+
        //  floattostr(xDayRab)+#13+floattostr(xDayOtr));

          if (XTEKTYPEOKLAD=1) or (XTEKTYPEOKLAD=3) then //оклад по часам  , часовая ставка
            begin
             if xDayRab<>0 then x:=DRound(INT((form1.kartOKLAD.Value*(1+(form1.configRK.Value+kNadbavka)/100)*Deleniecas(xDayOtr,xDayRab,1))),0);
            end
              else
            begin
             if k<>0 then x:=DRound(INT((form1.kartOKLAD.Value*(1+(form1.configRK.Value+kNadbavka)/100)*m/k)),0);
            end;

          if datam.kart2typeavans.Value=5 then x:=DRound(x*0.87,0);
          xVicet:=x-xVicet;  //база
          if xVicet<0 then xVicet:=0;
          if datam.kart2typeavans.Value=9 then x:=DRound(x-xVicet*0.13,0)
         end;

        if (datam.kart2typeavans.Value=6) or (datam.kart2typeavans.Value=7) or (datam.kart2typeavans.Value=10) then
         begin
          if k<>0 then x:=DRound(INT(form1.kartOKLAD.Value*m/k),0);
          if datam.kart2typeavans.Value=7 then x:=DRound(x*0.87,0);
          xVicet:=x-xVicet;  //база
          if xVicet<0 then xVicet:=0;
          if datam.kart2typeavans.Value=10 then x:=DRound(x-xVicet*0.13,0);
         end;




      end;


    //  Showmessage(floattostr(xcas1)+#13+floattostr(xcas2));


   if (datam.kart2typeavans.Value=11) or (datam.kart2typeavans.Value=12) then
    begin
     Form2.LocateGlDb(Form1.kartNls.Value,RMEs,RGod);
     x:=mainlib.DRound(form1.GlDbOKLAD.Value*DelenieCas(form1.GlDbDAYOTR.Value,form1.GlDbDAYRAB.Value,form1.GlDbDAYCAS.Value),form1.DRZn) ;
     x:=x+DRound(x*form1.configRK.Value/100,2);

     datam.qtmp2.close;
     datam.qTmp2.DatabaseName:=form1.DBDIR;
     datam.qtmp2.sql.clear;
     datam.qtmp2.sql.add('select * from obrt1new where wg='+floattostr(RGod)+' and wm='+floattostr(RMES)+' and nls='+floattostr(datam.kart2NLS.Value));
     datam.qtmp2.prepare;
     datam.qtmp2.open;
     while not datam.qtmp2.eof do
      begin
        form1.NACISL.Locate('kod',datam.qtmp2.fieldbyname('kod').asInteger,[loCaseInsensitive]);
        if form1.NACISLKODDOX.Value='2000' then
          begin
           y:=datam.qtmp2.fieldbyname('kr').asFloat;
           if form1.NACISLRK.Value then y:=y+DROund(y*form1.configRK.VAlue/100,2);
           x:=x+y;
          end;
       datam.qtmp2.Next;
      end;

     if datam.kart2typeavans.Value=12 then x:=x-DRound(x*0.13,0);

    end;


    if (form1.kart.FieldByName('tabel').asFloat=19) or (form1.kart.FieldByName('tabel').asFloat=21) or (form1.kart.FieldByName('tabel').asFloat=48) then
     begin
      if form1.sptabel2.Locate('ID',form1.kartTabel.Value,[loCaseInsensitive]) and (x<>0) then
       begin
        if MessageDlg(form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+
          'В соответствии с настройками Аванса начислена сумма: '+FloattostrF(x,ffNumber,12,2)+' руб.'+#13+
           'Код Табеля в настройках сотрудника ['+form1.sptabel2NAME.Value+']'+#13+
              'Обнулить аванс ?',mtWarning,[mbYes,mbNo],0) = mrYes then x:=0;

        
       end;
     end;

     sDoxod:=0; sBasa:=0; sNdfl:=0; sVicet:=0;

  //  showmessage(form1.kartfam.value);

 _rtfav:=false;   //заполнение доход, вычеты, ндфл  только если не проведен аванс
 if not vyplzp.Locate('NLS;WM;WG;TYPE',VarArrayOf([form1.kartNLS.Value,RMEs,RGod,2]),[loCAseInsensitive]) then
   begin
    _rtfav:=TRUE;
   end
    else
   begin
    if vyplzpPROVEDEN.Value=0 then _rtfav:=TRUE;
   end;

//  ShowMessage(form1.kartfam.value+#13+floattostr(x)+#13+floattostr(datam.kart2TYPEAVANS.Value));

   if (datam.kart2KAVANS.Value<>0) and (datam.kart2KAVANS.Value<>1) then
       begin
        x:=DRound(x*datam.kart2KAVANS.VAlue,2);
       // vavans.FieldByName('name').asString:=vavans.FieldByName('name').asString+', k='+Floattostr(datam.kart2KAVANS.VAlue);
       end;


  IF ((TUPD) and (TUpdNLS=-1)) OR ((TUPD) and (TUpdNLS=form1.kartNLS.Value))  OR (_rtfav) THEN     //расчет дохода, вычетов либо из обработки аванса TUPD либо если не проведен аванс
   BEGIN

     if (CheckBox199.Checked) and (RGod>=2022) then  // с аванса удержание НДФЛ
       begin
        if (datam.kart2TYPEAVANS.VAlue<=2) or (datam.kart2TYPEAVANS.VAlue=4) or (datam.kart2TYPEAVANS.VAlue=6)
                or(datam.kart2TYPEAVANS.VAlue=11) then
                  begin

                   // ZapolnMas;
                   // ObrabObrtNalKart;
                    sVicet:=FVic1(RMes)+FKartVic2(RMes,skodv)+FIjdev2(RMes,skodv);  //право всего, но если есть расчет факт берем факт pvbaseRVICET
                    if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then
                                         sVicet:=DRound(datam.pvbasePRAVO.Value-datam.pvbaseFIXVICET.Value,2);

                    if trim(form1.kartIMVYC_KOD.Value)<>'' then
                     begin
                     // showmessage(form1.kartfam.value+#13+'имущ.вычет');
                      ZapolnMas;
                      ObrabObrtNalKart;
                      sVicet:=sVicet+DImvyc[RMes];
                     end;

                    if sVicet<0 then sVicet:=0;
                  //  ShowMessage(form1.kartfam.value+#13+floattostr(sVicet));

                    rVicet:=0;
                    datam.qtmp.close;
                    datam.qtmp.databasename:=form1.dbdir;
                    datam.qtmp.sql.clear;
                    datam.qtmp.sql.add('select sum(rvicet) from sdoxod where nls='+floattostr(form1.kartnls.value));
                    datam.qtmp.sql.add('and mes='+floattostr(RMes)+' and god='+floattostr(RGod));
                    datam.qtmp.prepare;
                    datam.qtmp.open;
                    rvicet:=datam.qtmp.fields[0].asfloat;   //уже применено вычетов
                    datam.qtmp.close;
                    sVicet:=DRound(sVicet-rvicet,2); //остаток оставляем
                    if sVicet<0 then svicet:=0;

                   sDoxod:=x;
                   sBasa:=DRound(x-sVicet,2);
                   if sBasa<0 then sBasa:=0;
                   if form1.kartSTATUS.Value='2' then st:=30 else st:=13;
                   sNdfl:=DRound(sBasa*st/100,0);
                   x:=Dround(x-sNdfl,2);
                   xvicet0:=sVicet;
                   // x:=DRound(x-DRound(x*0.13,0),2);
                  end
       end;

     if (not CheckBox199.Checked) and (RGod>=2022) then  // с аванса удержание НДФЛ обратный расчет
       begin
        if (datam.kart2TYPEAVANS.VAlue<=2) or (datam.kart2TYPEAVANS.VAlue=4) or (datam.kart2TYPEAVANS.VAlue=6)
                or(datam.kart2TYPEAVANS.VAlue=11) then
                  begin
                    // ZapolnMas;
                    // ObrabObrtNalKart;
                    sVicet:=FVic1(RMes)+FKartVic2(RMes,skodv)+FIjdev2(RMes,skodv);  //право
                    if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then
                                       sVicet:=DRound(datam.pvbasePRAVO.Value-datam.pvbaseFIXVICET.Value,2);
                    if sVicet<0 then sVicet:=0;

                    rVicet:=0;
                    datam.qtmp.close;
                    datam.qtmp.databasename:=form1.dbdir;
                    datam.qtmp.sql.clear;
                    datam.qtmp.sql.add('select sum(rvicet) from sdoxod where nls='+floattostr(form1.kartnls.value));
                    datam.qtmp.sql.add('and mes='+floattostr(RMes)+' and god='+floattostr(RGod));
                    datam.qtmp.prepare;
                    datam.qtmp.open;
                    rvicet:=datam.qtmp.fields[0].asfloat;   //уже применено вычетов
                    datam.qtmp.close;
                    sVicet:=DRound(sVicet-rvicet,2); //остаток оставляем
                    if sVicet<0 then svicet:=0;

                   if form1.kartSTATUS.Value='2' then st:=0.30 else st:=0.13;
                   if x>sVicet then
                    begin
                     sDoxod:=Dround((x-st*sVicet)/(1-st),2);
                     sNdfl:=DRound(sDoxod-x,0);
                     sDoxod:=DRound(x+sNdfl,2);
                    end
                     else
                    begin
                     sndfl:=0;
                     sdoxod:=x;
                    end;
                   xvicet0:=sVicet;
                  end
       end;
   END;


     vavans.append;
     vavans.fieldbyname('nls').asFloat:=form1.kartnls.value;
     vavans.fieldbyname('Fio').asString:=form1.kartfam.value+' '+Copy(form1.kartim.Value,1,1)+'.'+Copy(form1.kartOt.Value,1,1)+'.';
     nn:=datam.kart2TYPEAVANS.AsInteger;
     if nn>1 then vavans.FieldByName('name').asString:=datam.FNameAvans(nn);
     if nn=0 then vavans.FieldByName('name').asString:=datam.FNameAvans(nn)+' '+FloattostrF(form1.kartAvans.Value,ffNumber,12,2);
     if nn=1 then vavans.FieldByName('name').asString:=datam.FNameAvans(nn)+' '+FloattostrF(form1.kartAvans.Value,ffNumber,12,2)+'%';


     vavans.FieldByName('oklad').asFloat:=xOklad;
     vavans.FieldByName('nadbavka').asFloat:=kNadbavka;
     vavans.FieldByName('rk').asFloat:=form1.configRK.Value;
     vavans.FieldByName('dayrab').asFloat:=DRound(xDayRab,2);
     vavans.FieldByName('dayotr').asFloat:=DRound(XDayOtr,2);
     vavans.FieldByName('vicet').asFloat:=xvicet0;
     vavans.FieldByName('Summa').asFloat:=x;
     vavans.post;

     if not vyplzp.Locate('NLS;WM;WG;TYPE',
            VarArrayOf([form1.kartNLS.Value,RMEs,RGod,2]),[loCAseInsensitive]) then
      begin
       if x>0 then
        begin
         xIdp:=xIdp+1;
         vyplzp.Append;
         vyplzp.FieldByName('idp').asFloat:=xIdp;
         vyplzp.FieldByName('type').asFloat:=2;
         vyplzp.FieldByName('nls').asFloat:=form1.kartNls.Value;
         vyplzp.FieldByName('numdok').asFloat:=0;
         vyplzp.FieldByName('wm').asFloat:=RMEs;
         vyplzp.FieldByName('wg').asFloat:=RGod;
         vyplzp.FieldByName('summa0').asFloat:=x;

         if rtfotp then MessageDlg(form1.kartFam.Value+' '+form1.kartIm.Value+' '+form1.kartOt.Value+#13+
         'Обнаружено в табеле отпуск/больничный/прочее отсутствие в первой половине месяца '+floattostr(ndotp)+'дн. из 15дн.'+#13+
         'Сумма к выплате аванса установлена равной '+Floattostr(x)+' руб.'+#13+
            'Вы можете скорректировать вручную данную сумму в соответствии с Трудовым договором/Внутренним положением',mtInformation,[mbOk],0);

         if rtf5 then MessageDlg(form1.kartFam.Value+' '+form1.kartIm.Value+' '+form1.kartOt.Value+#13+'Дата принятия на работу '+dATETOsTR(xDatP)+#13+
         'Сумма к выплате аванса установлена равной '+Floattostr(x)+' руб. как пропорция '+Floattostr(15-dd+1)+'дн./15дн.'+#13+
            'Вы можете скорректировать вручную данную сумму в соответствии с Трудовым договором/Внутренним положением',mtInformation,[mbOk],0);


         vyplzp.FieldByName('kod').asInteger:=xKod;
         vyplzp.FieldByName('proveden').asInteger:=0;
         if RGod>=2022 then
          begin
           vyplzp.fieldbyname('avdoxod').asfloat:=sDoxod;
           vyplzp.fieldbyname('avnalog').asfloat:=sNdfl;
           vyplzp.FieldByName('avvicet').asFloat:=xvicet0;
          end;
         vyplzp.post;
        end;
      end
       else
      begin
       if x<0 then x:=0;

       if (vyplzpPROVEDEN.Value=1) and (RGod>=2022) and (TUpd) then //не обработан аванс
         begin
           if form1.obrt2.Locate('nls;numdok',VarArrayOf([form1.kartnls.value,vyplzpIDP.Value]),[loCAseInsensitive]) then
            begin
             if form1.obrt2TAVANS.Value=1 then  //только не обработана запись аванса, иначе глюк
              begin
               vyplzp.edit;
               vyplzp.fieldbyname('avdoxod').asfloat:=sDoxod;
               vyplzp.fieldbyname('avnalog').asfloat:=sNdfl;
               vyplzp.FieldByName('avvicet').asFloat:=xvicet0;
               vyplzp.post;
              end;
            end
         end;


       if vyplzpPROVEDEN.Value=0 then
        begin
         vyplzp.Edit;

          if (RGod>=2022) then
          begin
           vyplzp.fieldbyname('avdoxod').asfloat:=sDoxod;
           vyplzp.fieldbyname('avnalog').asfloat:=sNdfl;
           vyplzp.FieldByName('avvicet').asFloat:=xvicet0;
          end;

         vyplzp.Fieldbyname('summa0').asFloat:=x;
         vyplzp.post;
             if rtfotp then MessageDlg(form1.kartFam.Value+' '+form1.kartIm.Value+' '+form1.kartOt.Value+#13+
         'Обнаружено в табеле отпуск/больничный/прочее отсутствие в первой половине месяца '+floattostr(ndotp)+'дн. из 15дн.'+#13+
         'Сумма к выплате аванса установлена равной '+Floattostr(x)+' руб.'+#13+
            'Вы можете скорректировать вручную данную сумму в соответствии с Трудовым договором/Внутренним положением',mtInformation,[mbOk],0);

        end;


     end;


     

   form1.kart.next;
  end;

end;

function TForm123.ProvDAtVypl():Boolean;
var rtf:Boolean;
begin

{ if (dateTimePicker1.DateTime<xD1) or (dateTimePicker1.DateTime>xD2) then
  begin
    MessageDlg('Расчетный период установлен '+namemes[RMEs]+#13+
                'Дата выплаты должна быть в промежутке '+
                  #13+DateToStr(xD1)+' '+DAteToStr(xD2)+#13+
                   'Исправьте дату выплаты',mtError,[mbOk],0);
    dateTimePicker1.DateTime:=xD;
    DateTimePicker1.SetFocus;
    rtf:=False;
  end
   else
    begin
     xD:=dateTimePicker1.DateTime;
     rtf:=True;
    end;
 }
      
     rtf:=True;


 ProvDatVypl:=rtf;
end;



procedure TForm123.DBGrid1DblClick(Sender: TObject);
begin
 JvXPButton23Click(Sender);
end;

procedure TForm123.JvXPButton24Click(Sender: TObject);
var OldIdp,oldNls,xIdsdoxod:Real;
    x,y:real;
begin
 if Query1.RecordCount<=0 then exit;
 oldIdp:=query1.fieldbyname('idp').asFloat;
 oldNls:=query1.fieldbyname('nls').asFloat;
 if MessageDlg('Исключить из выплаты '+Query1.fieldByname('FAM').asString+' '+Query1.fieldByname('IM').asString+
             ' '+Query1.fieldByname('OT').asString+#13+
               'Сумма = '+FloatToStr(Query1.fieldByname('SUMMA').asFloat),mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 if vyplzp.Locate('IDP',oldIdp,[loCaseInsensitive]) then
  begin
   form1.kart.Locate('nls',oldNls,[loCaseInsensitive]);
   if form1.obrt2.Locate('nls;numdok',VarArrayOf([oldNls,oldIdp]),[loCAseInsensitive]) then
    begin
     MessageDlg('Удалена и связаная операция выплаты (удержания) в карточке начислений/удержаний'+#13+
       'Сумма='+FloatToStr(form1.obrt2KR.Value),mtInformation,[mbOk],0);

     form1.DelSDoxodObrt2(form1.obrt2ID.Value);

     form1.obrt2.Delete;
    end;


   xIdsDoxod:=vyplzp.FieldByName('idsdoxod').asFloat;
   if form1.sdoxod.Locate('id',xIdsdoxod,[loCaseInsensitive]) then form1.sdoxod.Delete;

   vyplzp.Delete;
   vyplzp.FlushBuffers;
   ZaprosV2(0);
  end;
  query1.Locate('idp',oldIdp-1,[loCaseInsensitive]);
  JvXPButton12Click(nil);

end;

procedure TForm123.JvXPButton23Click(Sender: TObject);
var xNls,xIdp,xst:Real;
begin
 if Query1.RecordCount<=0 then exit;
 if Query1.FieldByNAme('proveden').asFloat=1 then
  begin
    MessageDlg('Исправление невозможно, уже проведено !'+#13+
           'Снимите сначала признак проведения (отменить проведение)'+#13+
            'затем только возможная корректировка',mtWarning,[mbOk],0);
    exit;
  end;
 xNls:=Query1.fieldByName('nls').asFloat;
 xIdp:=Query1.fieldByName('idp').asFloat;
 if not vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then exit;

 if (RGod>=2022) and (TYPEVYPLAT=2)
      // and (CheckBox199.Checked)
      then
   begin
    form3309:=tform3309.create(nil);
    form1.kart.locate('nls',xNls,[locaseinsensitive]);
    form3309.jvCalcEdit1.Value:=vyplzpAVDOXOD.Value;
    form3309.jvCalcEdit2.Value:=vyplzpAVVICET.Value;
    if form1.kartSTATUS.Value='2' then xst:=30 else xst:=13;
    form3309.stpn:=xst;
    form3309.showmodal;
    if form3309.TOK then
      begin
       vyplzp.edit;
       vyplzp.fieldbyname('avdoxod').asfloat:=form3309.jvCAlcEdit1.Value;
       vyplzp.fieldbyname('avnalog').asfloat:=form3309.jvCAlcEdit3.Value;
        vyplzp.fieldbyname('avvicet').asfloat:=form3309.jvCAlcEdit2.Value;
       vyplzp.fieldbyname('summa0').asfloat:=form3309.jvCAlcEdit4.Value;
       vyplzp.post;
      end;
    form3309.free;
   end
    else
   begin
    form124.ShowModal;
   end;
 ZaprosV2(0);
 Query1.Locate('idp',xIdp,[loCaseInsensitive]);
 JvXPButton12Click(nil);


end;

procedure TForm123.JvXPButton1Click(Sender: TObject);
begin
 form123.Close;
end;

procedure TForm123.JvXPButton6Click(Sender: TObject);
begin
  {НЕ ТРОГАТЬ !!!!!!! Осталась функция от кнопки удаленной}

 ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]

 xTYPE:=0; {для ZaprosV2}
 ZaprosV2(0);


 JvXPButton12Click(nil);
end;

procedure TForm123.JvXPButton2Click(Sender: TObject);
begin
   {НЕ ТРОГАТЬ !!!!!!! Осталась функция от кнопки удаленной}
  ZaprosV3;
  xTYPE:=2; {аванс, для ZaprosV2}
  ZaprosV2(0);
  JvXPButton12Click(nil);
end;

procedure TForm123.JvXPButton3Click(Sender: TObject);
var E:OleVariant;
    NStrok:Integer;
    NewValueArray:OleVariant;
    xFam,xIm,xOt:String;
    i:Integer;
    xItog:Real;
    fNameXls:String;
    nPved:String;
    xDat:TDate;
begin

 if Query1.RecordCount<=0 then
  begin
   MessageDlg('Отсутствует проведенние выплат',mtInformation,[mbOk],0);
   exit;
  end;

   if Query1.FieldByName('proveden').asInteger<>1 then
    begin
       MessageDlg('Данная выплата не проведена',mtInformation,[mbOk],0);
       exit;
    end;

   nPVed:='';
   xDat:=EncodeDate(2000,1,1);
   if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]) then
    begin
      nPVed:=vyplzpNote.Value  ;
      xDat:=vyplzpDatDok.Value
    end;

   if MessageDlg('Сформировать:'+#13+npVed+#13+FormatDatetime('dd.mm.yyyy',xDat),mtInformation,[mbYes,mbNo],0) = mrNo then exit;


 fNameXls:=mainlib.GetNameXLSn('Экспорт','exp')+'.xls';
 if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\export.xls',form1.DBCurr+'\TMP_XLS\'+fNameXLS) then
   begin
    MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
    exit;
   end;

 E:=CreateOleObject('Excel.Application');
 E.WorkBooks.Open(form1.DBCurr+'\TMP_XLS\'+fNameXLS);

 E.Visible:=True;
 E.Application.WindowState:=2;

 E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:=form1.config2NAME.VAlue+' '+'ИНН '+form1.config2INN.VAlue;
 E.ActiveWorkBook.Sheets.Item[1].Range['A2'].Value:='Платежная '+nPVed;

 // E.ActiveWorkBook.Sheets.Item[1].Range['A4'].Value:='Платежная ведомость № '+'________'+' за '+ansilowercase(namemes[RMes])+' '+Floattostr(RGod)+'г.';
    E.ActiveWorkBook.Sheets.Item[1].Range['A4'].Value:='Дата выплаты: '+FormatDatetime('dd.mm.yyyy',xDat);


 mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
 E.ActiveWorkBook.Sheets.Item[1].Range['A6'].Value:=form1.configRUKOVDOLJN.Value+' '+xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';


 NewValueArray := VarArrayCreate([1, 1, 1, 4], varVariant);

 NewValueArray[1,1]:='№ п/п';
 NewValueArray[1,2]:='ФИО';
 NewValueArray[1,3]:='Сумма';
 NewValueArray[1,4]:='Роспись в получении';

 NStrok:=8;
 if not form1.TCheckOO then
      E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok)]:=NewValuearray
       else
         excellib.PExcelVyvod(E,1,'A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok),NewValuearray) ;


 Query1.First;
 i:=0;  xItog:=0;
 while not Query1.Eof do
  begin
    vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]);
    if (Query1.FieldByName('proveden').asInteger=1) and (nPved=vyplzpNote.Value) and (xDat=vyplzpDatDok.Value) then
     begin
      i:=i+1;
      NStrok:=NStrok+1;
      NewValueArray[1,1]:=i;
      form1.kart.Locate('nls',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]);
      NewValueArray[1,2]:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;
      NewValueArray[1,3]:=FloattostrF(Query1.FieldByNAme('summa').asFloat,ffNumber,12,2);
      xItog:=xItog+Query1.FieldByNAme('summa').asFloat;
      NewValueArray[1,4]:='';
      if not form1.TCheckOO then
          E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok)]:=NewValuearray
              else
                  excellib.PExcelVyvod(E,1,'A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok),NewValuearray) ;
     end;
   Query1.Next;
  end;


      try
         for i:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A'+IntToStr(8),'D'+IntToStr(Nstrok)].Borders[i].LineStyle:=1;
       except
       end;

      NStrok:=NStrok+1;
      NewValueArray[1,1]:='';
      NewValueArray[1,2]:='';
      NewValueArray[1,3]:=FloattostrF(xItog,ffNumber,12,2);
      NewValueArray[1,4]:='';
    if not form1.TCheckOO then
         E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok)]:=NewValuearray
          else
             excellib.PExcelVyvod(E,1,'A'+IntToStr(Nstrok),'D'+IntToStr(Nstrok),NewValuearray) ; ;

   E.ActiveWorkbook.Sheets.Item[1].Columns[1].ColumnWidth:=5;
   E.ActiveWorkbook.Sheets.Item[1].Columns[2].ColumnWidth:=35;
   E.ActiveWorkbook.Sheets.Item[1].Columns[3].ColumnWidth:=10;
   E.ActiveWorkbook.Sheets.Item[1].Columns[4].ColumnWidth:=25;


  Nstrok:=Nstrok+2;
  E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(Nstrok)].value:='Итого по платежной ведомости выплачено: _________________________________';
  Nstrok:=Nstrok+2;
  E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(Nstrok)].value:='Депонировано: _________________________________';
  Nstrok:=Nstrok+2;
  E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(Nstrok)].value:='Кассир: ______________________'+form1.configKASSIR.AsString;


  try
   E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(1)+':D'+IntToStr(Nstrok)].font.name:='Calibri';
   E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(1)+':D'+IntToStr(Nstrok)].font.size:=10;
   E.ActiveWorkbook.Sheets.Item[1].Range['A'+inttostr(1),'A'+inttostr(Nstrok)].Rows.RowHeight:=19;
   E.ActiveWorkbook.Sheets.Item[1].Range['A'+inttostr(3),'A'+inttostr(3)].Rows.RowHeight:=5;
   E.ActiveWorkbook.Sheets.Item[1].Range['A'+inttostr(5),'A'+inttostr(5)].Rows.RowHeight:=5;
   E.ActiveWorkbook.Sheets.Item[1].Range['A'+inttostr(7),'A'+inttostr(7)].Rows.RowHeight:=5;
  except
  end;

 E.Application.WindowState:=-4137;
 E:=UnAssigned;

 EXIT;


 if MessageDlg('Разбивать на страницы ?',mtInformation,[mbYes,mbNo],0) = mrYes
     then form123.NPVed:=20 else form123.NPVed:=10000;

 if ProvDatVypl then PlatVed;

end;

procedure TForm123.JvXPButton4Click(Sender: TObject);
begin

 if not ProvDatVypl then exit;

 if Query1.RecordCount<=0 then exit;

 if Query1.fieldByName('PROVEDEN').asInteger<>1 then
  begin
   MessageDlg('Данная выплата по сотруднику не проведена'+#13+'(кнопка Провести)',mtInformation,[mbOk],0);
   exit;
  end;

 if not vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then exit;


 PopupMenu3.Popup(Panel1.Left+JvXPButton4.Left+JvXPButton4.Width+form123.left,Panel1.Top+JvXPButton4.Top+JvXPButton4.Height+form123.top);



end;

procedure TForm123.JvXPButton5Click(Sender: TObject);
begin

 if ProvDatVypl then PlatVed2;
end;

procedure TForm123.JvXPButton7Click(Sender: TObject);
var i:Integer;

    x,x0:Real;

    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
    xNum,oldxNum:Integer;
    npp:Integer;
    xNls,xIDP:Real;
    xNlsPlat,xNlsPlat2:Real;
begin

 Edit1.Text:='1';
 //ShowMessage(Edit1.text);

 xNlsPlat:=form1.GetPlatNls(form1.config2Namekrat.Value,form1.config2INN.Value,form1.config2KPP.Value,form1.config2RSC.Value,form1.config2KSC.Value,
     form1.config2Bik.Value,form1.config2BankName.Value,form1.config2Name.Value); //Плательщик в п/п
 if xNlsPlat<0 then
   begin
    MessageDlg('Не заполнены данные Плательщика'+#13+form1.config2Name.Value,mtInformation,[mbOk],0);
    exit;
   end;
 form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;
 xNlsPlat2:=form1.GetPlatNls(form1.kartNamepoluc.asString,form1.kartBankINN.Value,form1.kartBankKPP.Value,form1.kartBankRSC.Value,
   form1.kartBankKSC.Value,
     form1.kartBankBik.Value,form1.KartBankName.Value,form1.kartNamepoluc.Value); //Получатель в п/п
 if xNlsPlat2<0 then
  begin
   MessageDlg('Не заполнены данные Получателя'+#13+form1.kartFam.Value+' '+form1.kartIm.Value+' '+form1.kartOt.VAlue+#13+
     'Заполнение реквизитов производится в карточке сотрудника - > Реквизиты платежей - > Банк',mtInformation,[mbOk],0);
   exit;
  end;

if not ProvDatVypl then exit;

if Query1.RecordCount<=0 then exit;

 if Query1.fieldByName('PROVEDEN').asInteger<>1 then
  begin
   MessageDlg('Данная выплата по сотруднику не проведена'+#13+'(кнопка Провести)',mtInformation,[mbOk],0);
   exit;
  end;


xIDP:=Query1.FieldByNAme('IDP').asFloat;

  i:=0;
  x0:=0;
  try
   xNum:=StrToInt(Trim(Edit1.text));
  except
   xNum:=1;
  end;
  npp:=0;

    rtf:=false;
    x:=Query1.FieldByName('SUMMA').asFloat ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=2 then rtf:=True;
     end;

    if not rtf then
     begin
      MessageDlg('Невозможно формирование для данного места выплаты'+#13+'По данной операции в настройках Справочника удержаний не выставлен признак, что документом является платежное поручение',mtError,[mbOk],0);
      exit;
     end;

 form93.Query1.DatabaseName:=form1.DBDIR;



    IF (x<>0) and (rtf)  then
     begin
       npp:=npp+1;

       form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;

       vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);

      // ShowMessage(vyplzpNUMDOK.asString);

      oldxNum:=Trunc(vyplzpNumdok.asFloat);

       if vyplzpNUMDOK.Value=0 then
        begin
         xNum:=Trunc(form1.GetMaxNPP());
         oldxNum:=Trunc(vyplzpNumdok.asFloat);
         vyplzp.Edit;
         if xNum>0 then vyplzp.fieldByName('numdok').asFloat:=xNum;
        { if vyplzpDATDOK.Value<EncodeDAte(2000,1,1) then vyplzp.fieldByName('datdok').asDAteTime:=xD;
        } vyplzp.Post;

       // ShowMEssage(floattostr(vyplzp.fieldByName('numdok').asFloat));

         xNum:=xNum+1;
        end;
        form124.ShowModal;
       if form124.TOk then
        begin
         frReport1.LoadFromFile('plpor.frf');

         form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;
         vyplzp.Filter:='IDP='+FloatToStr(Query1.fieldByName('idp').asFloat);
         vyplzp.Filtered:=True;
         frReport1.ShowReport;
         vyplzp.Filter:='';
         vyplzp.Filtered:=False;
        end
         else
        begin
           vyplzp.Edit;
           vyplzp.fieldByName('numdok').asFloat:=oldxNum;
           vyplzp.Post;
        end;
     END;


 if not form124.TOk then EXIT;

 // Edit1.TExt:=IntToStr(xNum);

  xNls:=Query1.fieldByName('nls').asFloat;


  //создаем п/п
   form1.plpor.Databasename:=form1.DBDIR;
   form1.plpor.Open;
   form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;
     vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);

  if not form1.plpor.Locate('Summa;num;dat',VarArrayOf([vyplzp.FieldByName('summa').asFloat,vyplzp.FieldByName('numdok').asFloat,
     vyplzp.FieldByName('datdok').asDateTime]),[loCaseInsensitive]) then
  begin
   if MessageDlg('Сохранить сформированние платежное поручение в базе платежных поручений',mtInformation,[mbYes,mbNo],0) = mrYes then
    begin
     form1.plpor.Append;
     form1.plpor.FieldByName('KOD').asFloat:=0;
     form1.plpor.FieldByName('Summa').asFloat:=vyplzp.FieldByName('summa').asFloat;
     form1.plpor.FieldByName('NDS').asFloat:=18;
     form1.plpor.FieldByName('NUM').asFloat:=vyplzp.FieldByName('numdok').asInteger;
     form1.plpor.FieldByName('NDSAVTO').asBoolean:=False;
     form1.plpor.FieldByName('NLS1').asFloat:=xNlsPlat;
     form1.plpor.FieldByName('NLS2').asFloat:=xNlsPlat2;
     form1.plpor.FieldByName('DAT').asdateTime:=vyplzp.FieldByName('datdok').asDateTime;
     form1.plpor.FieldByName('DOPPOLE').asString:=' ';
     form1.plpor.FieldByName('VID').asString:='Электронно';
     form1.plpor.FieldByName('OCER').asString:='6';
     form1.plpor.FieldByName('VIDOPLAT').asString:='01';
     form1.plpor.fieldbyname('NAZPLAT1').asString:=vyplzp.FieldByName('note').asString+' Для зачисления на счет '+form1.kartLSC.Value;
     form1.plpor.fieldbyname('NAZPLAT2').asString:='Получатель '+form1.kartFam.Value+' '+form1.kartIm.Value+' '+form1.kartOt.Value+' Без НДС.';
     form1.plpor.Post;
     form1.plpor.FlushBuffers;
     MessageDlg('Выполнено сохранение',mtInformation,[mbOk],0);
    end;
 end;
   form1.plpor.close;


    ZaprosV2(0);
  Query1.Locate('nls',xNls,[loCaseInsensitive]);



end;

procedure TForm123.JvXPButton12Click(Sender: TObject);
var x,x1,xNls:Real;
begin

 CurrencyEdit1.Value:=0;
 CurrencyEdit2.Value:=0;
 CurrencyEdit3.Value:=0;


 if Query1.RecordCount<=0 then exit;
 
 x:=0; x1:=0;

 xNls:=Query1.fieldByName('nls').asFloat;
 DBGrid1.DataSource:=nil;

 Query1.First;
 while not Query1.Eof do
  begin
    x:=x+Query1.FieldByName('summa').Value;
    if Query1.FieldByNAme('proveden').Value=1 then x1:=x1+Query1.FieldByName('summa').Value;
   Query1.Next;
  end;
 CurrencyEdit1.Value:=x;
 CurrencyEdit2.Value:=x1;
 CurrencyEdit3.Value:=x-x1;
 Query1.Locate('nls',xNls,[loCaseInsensitive]);
 DBGrid1.DataSource:=DataSource1;

end;

procedure TForm123.JvXPButton13Click(Sender: TObject);
begin

  Form54:=TForm54.Create(Self);
  form54.JvXPButton2.Enabled:=False;
  if TYPEVYPLAT=2 then
   begin
     form54.RadioButton1.Checked:=True;
     form54.RadioButton2.Enabled:=False;
     form54.RadioButton4.Enabled:=False;
   end
    else
   begin
     form54.RadioButton2.Checked:=True;
     form54.RadioButton1.Enabled:=False;
     form54.RadioButton3.Enabled:=False;
   end ;

  Form54.ShowModal;
      if form54.TOk=0 then
        begin
         form54.Free;
         Exit;
        end;
      if form54.TOk=2 then
       begin
        form54.FZT53(0); //вывод аванс, касса разделяет переключтаелем
        form54.Free;
        Exit;
       end;


end;

procedure TForm123.JvXPButton14Click(Sender: TObject);
begin

 if (RGod>=2023) and (CheckBox199.Visible) and (not CheckBox199.Checked) then
  begin
  // MessageDlg('Выставите флаг <Удержать НДФЛ с аванса> (с 2023 г.)',mtInformation,[mbOk],0);
  // exit;
  end;

 form40.tavans.active:=true;
 PopupMenu1.Popup(JvXPButton14.Left+JvXPButton14.Width+form123.left,JvXPButton14.Top+JvXPButton14.Height+form123.top)  ;
 form40.tavans.active:=false;


end;

function Tform123.ProvAvansKod(xNls:Real;sfio:String):Boolean;
var rtf:Boolean;
begin

 //************

 if form1.config2PRAVANS.Value=1 then  //для Митрошкиной Е. аванс по Подразделениям
   begin

     Form_58.SmenaKodAvansFil(RMes,RGod,form1.kartNls.Value)  ; //аванс по подразделениям перегруппировать
     EXIT;
   end;

 //**************


// ShowMEssage('Ok');

 rtf:=true;

                         datam.qTmp.Close;
                         datam.qTmp.DatabaseName:=form1.DBDIR;
                         datam.qTmp.SQL.Clear;
                         datam.qTmp.SQL.Add('select g.* from glnew g, kart k, sdoxod s where s.nls=g.nls and k.nls=g.nls ');
                         datam.qTmp.SQL.Add(' and s.sdoxod<>0 and g.wm<='+floattostr(Rmes));
                         datam.qTmp.SQL.Add(' and g.wm=s.mes and s.kodnac=0');
                         datam.qTmp.SQL.Add(' and g.wg='+floattostr(rGod));
                         datam.qTmp.SQL.Add(' and s.god='+floattostr(rGod));
                         datam.qTmp.SQL.Add(' and g.nls='+floattostr(xNls));
                         datam.qTmp.SQL.Add(' and g.oklad*g.dayotr=0');
                         datam.qTmp.prepare;
                         datam.qTmp.open;
                         datam.qTmp.First;
                         while not datam.qTmp.eof do
                          begin


                            if MessageDlg('Обнаружена выплата аванса,'+#13+
                               'но начисление по окладу/тарифу отсутствует'+#13+
                                    '(выплата аванса по умолчанию списывается с оклада)'+#13+
                                      namemes[datam.qTmp.fieldbyname('wm').asInteger]+#13+
                                       sfio+#13+
                                           'Попробовать привязать выплаченный аванс к другим кодам начислений ?',mtInformation,[mbYes,mbNo],0) = mrYes
                                             then
                                                begin
                                                 if not form_58.SmenaKodAvans(datam.qTmp.fieldbyname('wm').asfloat,RGod,xNls) then
                                                   begin
                                                      if MessageDlg('Другие доступные коды начислений не найдены для привязки'+#13+
                                                       'Все равно продолжить проведение выплаты ?',mtWarning,[mbYes,mbNo],0) = mrNo then  //меняем код аванса на начисление любое с кодом 2000
                                                              rtf:=false;
                                                   end;
                                                end
                                                 else
                                                   begin
                                                    MessageDlg('Выплата отменена',mtInformation,[mbOk],0);
                                                      rtf:=false;
                                                   end;



                            
                           datam.qTmp.next;
                          end;
                         datam.qTmp.close;
     ProvAvansKod:=rtf;                    
end;


procedure TForm123.N1Click(Sender: TObject);
var rtf:Boolean;
    OldIDP,xIdSdoxod:Real;
    s:String;
    xDat:TDate;
    i,xKod:integer;
    xNote:String;
    xNPVed:Real;
    xid,xIdVypl:real;
    TProvAvans2:Integer;
    uIdp,xnls:Real;
begin

 if Query1.RecordCount<=0 then EXIT;

 uIdp:=Query1.fieldByName('IDP').asFloat;
 xnls:=Query1.fieldByName('nls').asFloat;

 if not vyplzp.Locate('IDP',Query1.fieldByName('IDP').asFloat,[loCaseInsensitive]) then
  begin
   exit;
  end;


 if vyplzp.FieldByNAme('proveden').asFloat=1 then
  begin
    MessageDlg('Уже проведено !'+#13+
           'Для внесения изменений снимите сначала признак проведения (отменить проведение)'+#13+
            'затем только возможная корректировка',mtWarning,[mbOk],0);
    exit;
  end;

  if (not T_OWKOD) and (TYPEVYPLAT=0) then
    begin
     if DRound(vyplzpAVANS.Value,2)<>0 then
      begin
        MessageDlg('Итоговая выплата з/п не возможна по данному Сотруднику, т.к. имеется не обработанный Аванс'+#13+
           'Для проведения промежуточных выплат (отпускные, больничные, премиальные) установите Фильтр по кодам начислений и повторите данную операцию.'+#13+
            'Но имейте ввиду, что по итогам месяца сумма не выплаченных начислений должно будет хватать, чтобы обработать Аванс',mtWarning,[mbOk],0);
        exit;
            
      end;
    end;


  TProvAvans2:=0;


 if form123.TYPEVYPLAT=0 then    //выплата з/п
  begin
   TProvAvans2:=Prov2Avans(query1.fieldbyname('nls').asFloat);

   if TProvAvans2=1 then //проверка два аванса наличие
    begin
     MessageDlg('По сотруднику '+FZaprosFio(query1.fieldbyname('nls').asFloat)+' обнаружена выплата Аванса, но начислений за текущий период не хватает',mtWarning,[mbOk],0);
     EXIT;
    end;

   if TProvAvans2=2 then //проверка два аванса наличие
     begin
       {if MessageDlg('По сотруднику '+FZaprosFio(query1.fieldbyname('nls').asFloat)+' обнаружена выплата Аванса, необходимо перепровести выплаты'+#13+
          'Перепровести ?',mtWarning,[mbYes,mbNo],0) = mrNo then Exit;
       }
        form1.PereprovAvans(query1.fieldbyname('nls').asFloat,0,0);
    end;
  end;



                     //аванс проведен но нет выплаты !
                       if (vyplzpSumma.Value<>0) and (TProvAvans2=0) then
                        begin
                        { if not ProvAvansKod(vyplzpNls.Value,Query1.fieldByname('FAM').asString+' '+Query1.fieldByname('IM').asString+' '+
                                         Query1.fieldByname('OT').asString) then EXIT;
                        }
                       end;




 form98:=TForm98.Create(nil);

   form98.Label1.Caption:=Query1.fieldByname('FAM').asString+' '+Query1.fieldByname('IM').asString+' '+Query1.fieldByname('OT').asString;
   datam.Query1.DataBaseName:=form1.DBDIR;
   datam.Query1.Close;
   datam.Query1.SQL.Clear;
   datam.Query1.SQL.Add('select n.kod,count(o.kr) from obrt1new o, nacisl n where nls='+FloatToStr(query1.FieldByNAme('nls').asFloat)+' and pn<>1 and o.kod=n.kod ');
   datam.Query1.SQL.Add('and WG='+IntToStr(RGod)+' and WM='+IntToStr(RMes)+' group by n.kod');
   datam.Query1.Prepare;
   datam.Query1.Open;
   form98.RxCheckListBox1.Items.Clear;
   form98.rxCheckListBox1.Items.Add(FloatToStr(0)+'  - оклад');

    datam.Query1.First;
       while not datam.Query1.Eof do
           begin
                if form1.NACISL.Locate('KOD',datam.Query1.Fields[0].asFloat,[loCaseInsensitive])
                  then
                     form98.rxCheckListBox1.Items.Add(floattostr(form1.nacislkod.value)+'  - '+ansilowercase(form1.nacislName.Value));

            datam.Query1.Next;
           end;
   datam.Query1.Close;

   if form123.TYPEVYPLAT=0 then
    begin
     form98.CheckBox1.Checked:=True;



     for i:=0 to form98.RxCheckListBox1.Items.Count-1 do  form98.RxCheckListBox1.Checked[i]:=True;
     if Query1.FieldByName('sdoxod').asFloat<>0 then
      begin
       form98.CheckBox1.Checked:=False;  //частичная выплата
       for i:=0 to form98.RxCheckListBox1.Items.Count-1 do  form98.RxCheckListBox1.Checked[i]:=False;
      end;
    end;

   if form123.TYPEVYPLAT=2 then
    begin
     if form1.TREJIMAVTO<>1 then form98.CheckBox2.Enabled:=False;
     form98.CheckBox3.Enabled:=False;
     form98.CheckBox1.Checked:=False;
     form98.CheckBox1.Visible:=False;
     form98.RxCheckListBox1.Visible:=False;
     form98.SpeedButton1.Visible:=false;
     form98.SpeedButton2.Visible:=false;
     for i:=0 to form98.RxCheckListBox1.Items.Count-1 do  form98.RxCheckListBox1.Checked[i]:=False;
    end;

 form98.CheckBox3.Checked:=False;

 try
  if form123.TYPEVYPLAT=2 then xDat:=EncodeDAte(RGod, RMEs, trunc(form1.config2DV1.Value));
  if (form123.TYPEVYPLAT=0) and (RMEs<>12) then xDat:=EncodeDAte(RGod, RMEs+1, trunc(form1.config2DV2.Value));
  form123.fDat:=xDat;
 except
 end;

 form98.ShowModal;
 form98.CheckBox3.Enabled:=True;
 form98.CheckBox2.Enabled:=True;


 xDat:=form98.DateEdit1.Date;
 xNote:=Trim(form98.Edit5.Text);
 if form98.TOk=0 then
   begin
    form98.free;
    exit;
   end;




 oldIDP:=Query1.FieldByNAme('IDP').asFloat;

 if not ProvDatVypl then exit;

 s:='Дата выплаты  ';
 s:=s+ FormatDateTime('dd.mm.yyyy',xDat)+#13+'Примечание: '+xNote;

 {
 if MessageDlg('Провести расходные операции по Сотруднику '+#13+Query1.FieldByNAme('FAM').asString+' '+copy(Query1.FieldByNAme('IM').asString,1,1)+'.'+
              copy(Query1.FieldByNAme('OT').asString,1,1)+'.'+#13+s,mtWarning,[mbYes,mbNo],0)=mrNo then
                begin
                  form98.Free;
                  exit;
                end;
 }

   datam.Query2.DataBaseName:=form1.DBDIR;
   datam.Query2.Close;                           //номер плат.ведомости
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('select max(id) from sdoxod');
   datam.Query2.Prepare;
   datam.Query2.Open;
   xIdsdoxod:=datam.Query2.Fields[0].asFloat;
   datam.Query2.Close;

   datam.Query2.DataBaseName:=form1.DBDIR;
   datam.Query2.Close;                           //номер плат.ведомости
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('select max(npved) from obrt2new where wg='+FloatToStr(RGod));
   datam.Query2.Prepare;
   datam.Query2.Open;
   xNPVed:=datam.Query2.Fields[0].asFloat+1;
   form228:=Tform228.Create(nil);
   form228.RxCalcEdit1.Value:=xNPVed;
   form228.ShowModal;
   if form228.TOk then xNPVed:=form228.RxCalcEdit1.Value else xNPVed:=-1;;
   form228.Free;
   if xNPVed=-1 then exit;


 rtf:=True;
    if (Query1.fieldByName('KOD').asInteger<=0) and
            (Query1.fieldByName('SUMMA').asFloat<>0)  then rtf:=false;
 if not rtf then
  begin
   MessageDlg('По данному сотруднику'+#13+'не определено место выплаты',mtError,[mbOk],0);
   form98.Free;
   EXIT;
  end;

  IF (Query1.FieldByName('SUMMA').asFloat<>0) and (Query1.FieldByName('PROVEDEN').asFloat=0) then
     begin

      form1.Obrt2.Append;
      xIdVypl:=form1.Obrt2ID.Value;
     // ShowMessage(floattostr(form1.Obrt2ID.Value));
      form1.Obrt2.FieldByName('NLS').asFloat:=Query1.fieldByName('NLS').asFloat;
      form1.Obrt2.FieldByName('KR').asFloat:=Query1.fieldByName('SUMMA').asFloat;
      form1.Obrt2.FieldByName('WM').asFloat:=RMes;
      form1.Obrt2.FieldByName('WG').asFloat:=RGod;
      form1.Obrt2.FieldByName('KOD').asFloat:=Query1.fieldByName('KOD').asFloat;
      form1.obrt2.FieldByName('DATPROV').asDateTime:=xDat;
      form1.obrt2.FieldByName('NUMDOK').asFloat:=Query1.FieldByName('IDP').asFloat;
      form1.Obrt2.FieldByName('npved').asFloat:=xNPVed;
      form1.Obrt2.Fieldbyname('xs').asFloat:=TXS;

      if TYPEVYPLAT=0 then form1.Obrt2.Fieldbyname('TAVANS').asFloat:=0 else form1.Obrt2.Fieldbyname('TAVANS').asFloat:=1;

      form1.Obrt2.Post;

  //    showmessage(floattostr(xidvypl));

    IF (form98.CheckBox2.Checked) and (form1.TREJIMAVTO=1) THEN     //дата выплаты дохода
     BEGIN
      IF TYPEVYPLAT=0 THEN
        BEGIN
         form1.FProvedenVypl2(query1.fieldbyname('nls').asFloat,xIdVypl,Query1.fieldByName('SUMMA').asFloat,xDat,2,true)  ;
        END
         ELSE
        BEGIN
         //аванс
          datam.Query1.close;
          datam.query1.sql.clear;
          datam.query1.sql.add('select max(id) from sdoxod');
          datam.query1.prepare;
          datam.query1.Open;
          xid:=datam.query1.Fields[0].asfloat;
          datam.query1.close;

          xid:=xid+1;
       {
          form1.sdoxod.Append;
          form1.sdoxod.fieldbyname('id').asFloat:=xid;
          form1.sdoxod.fieldbyname('tavans').asFloat:=1;
          form1.sdoxod.fieldbyname('TPRAV2').asFloat:=1;
          form1.sdoxod.fieldbyname('idvypl').asFloat:=xIdVypl;
          form1.sdoxod.fieldbyname('nls').asFloat:=query1.fieldbyname('nls').asFloat;
          form1.sdoxod.fieldbyname('kodnac').asFloat:=0;
          form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
          form1.sdoxod.fieldbyname('mes').asFloat:=RMes;
          form1.sdoxod.fieldbyname('god').asFloat:=RGod;
          form1.sdoxod.fieldbyname('sdoxod').asFloat:=query1.fieldByName('summa').asFloat;
          form1.sdoxod.fieldbyname('nalog').asFloat:=0;
          form1.sdoxod.post;
        }

          MessageDlg('Выплата Аванса проведена, но не разнесена по налоговым регистрам'+#13+
              'Для окончательной выплаты аванса необходимо выполнить <Обработать аванс>, в модуле <Выплата заработной платы>'+#13+
                'Данная операция производится после проведения всех начислений перед итоговой выплатой заработной платы за отчетный период ',
                   mtInformation,[mbOk],0);


         END; ;
      END;



      if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asFloat,[loCaseInsensitive]) then
       begin
        vyplzp.Edit;
        vyplzp.fieldByName('proveden').asInteger:=1;
        vyplzp.fieldByName('datdok').asDateTime:=xDat;
        if vyplzp.fieldbyname('sdoxod').asFloat<>0 then
          begin
           vyplzp.fieldByName('dat').asDateTime:=xDat;
           form1.sdoxod.Append;
           form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
           form1.sdoxod.fieldbyname('sdoxod').asFloat:=vyplzp.fieldbyname('sdoxod').asFloat;
           form1.sdoxod.fieldbyname('mes').asFloat:=vyplzp.fieldbyname('mes').asFloat;
           form1.sdoxod.fieldbyname('god').asFloat:=vyplzp.fieldbyname('god').asFloat;
           form1.sdoxod.fieldbyname('nalog').asFloat:=vyplzp.fieldbyname('snalog').asFloat;
           form1.sdoxod.fieldbyname('nls').asFloat:=vyplzp.fieldbyname('nls').asFloat;
           form1.sdoxod.fieldbyname('kodnac').asFloat:=vyplzp.fieldbyname('kodnac').asFloat;
           form1.sdoxod.fieldbyname('id').asFloat:=xIdSdoxod+1;
           form1.sdoxod.Post;
           vyplzp.fieldbyname('idsdoxod').asFloat:=form1.sdoxod.FieldByName('id').asFloat;
          end;
        vyplzp.FieldByName('note').asString:='Вед.№'+FloatToStr(xNPVed)+', '+xNote;
        vyplzp.post;


       end;

     END;

 ZaprosV2(0);
 Query1.Locate('IDP',OldIDP,[loCaseInsensitive]);

 JvXPButton12Click(nil);

IF (form1.TREJIMAVTO<>1) and (form98.CheckBox2.Checked) THEN
 BEGIN
 {if form123.TYPEVYPLAT=0 then
  begin
   datam.Query1.Close;
   datam.Query1.SQL.Clear;
   datam.Query1.SQL.Add('UPDATE glnew SET datoklad='+#39+FormatDatetime('dd.mm.yyyy',xD)+#39);
   datam.Query1.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
   datam.Query1.SQL.Add('and nls='+FloatToStr(Query1.fieldByName('NLS').asFloat));
   datam.Query1.SQL.Add('and datoklad<'+#39+'01.01.2000'+#39);
   datam.Query1.Prepare;
   datam.Query1.ExecSQL;
   datam.Query1.Close;


   datam.Query1.Close;
   datam.Query1.SQL.Clear;
   datam.Query1.SQL.Add('UPDATE obrt1new SET datprov='+#39+FormatDatetime('dd.mm.yyyy',xD)+#39);
   datam.Query1.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
   datam.Query1.SQL.Add('and nls='+FloatToStr(Query1.fieldByName('NLS').asFloat));
   datam.Query1.SQL.Add('and datprov<'+#39+'01.01.2000'+#39);
   datam.Query1.Prepare;
   datam.Query1.ExecSQL;
   datam.Query1.Close;
  end;
 }
 {дата выплаты дохода}


 //*****************************************************
  for i:=0 to form98.RxCheckListBox1.Items.Count-1 do
   begin
    if form98.RxCheckListBox1.Checked[i] then
      begin
        xKod:=StrToInt(Trim(Copy(form98.rxCheckListBox1.Items[i],1,2)));

         if  xKod=0 then
           begin
            datam.Query2.Close;
            datam.Query2.SQL.Clear;
            datam.Query2.SQL.Add('UPDATE glnew SET datoklad='+#39+FormatDatetime('dd.mm.yyyy',form98.DateEdit1.Date)+#39);
            datam.Query2.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
            datam.Query2.SQL.Add('and nls='+FloatToStr(Query1.FieldByname('NLS').asFloat));
            if form98.CheckBox1.Checked then datam.Query2.SQL.Add('and (not datoklad>'+#39+'01.01.2005'+#39+')');
            datam.Query2.Prepare;
            datam.Query2.ExecSQL;
            datam.Query2.Close;
           end;

           if xKod<>0 then
             begin
              datam.Query2.Close;
              datam.Query2.SQL.Clear;
              datam.Query2.SQL.Add('UPDATE obrt1new SET datprov='+#39+FormatDatetime('dd.mm.yyyy',form98.DateEdit1.Date)+#39);
              datam.Query2.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
              datam.Query2.SQL.Add('and nls='+FloatToStr(Query1.FieldByname('NLS').asFloat));
              if form98.CheckBox1.Checked then datam.Query2.SQL.Add('and (not datprov>'+#39+'01.01.2005'+#39+')');
              datam.Query2.SQL.Add('and kod='+FloatToStr(xKod));
              datam.Query2.Prepare;
              datam.Query2.ExecSQL;
              datam.Query2.Close;
             end;


      end;
   end;
 END;

  if (TYPEVYPLAT=0) AND (form1.TREJIMAVTO<>1) AND (form98.CheckBox3.Checked) then if MessageDlg('Выполнено !'+#13+sSoobVyp+#13+'Перейти к процедуре заполнения перечисления НДФЛ ?',mtInformation,[mbYes,mbNo],0)=mrYes
     then Form1.N65Click(nil);;


 if (TYPEVYPLAT=0) AND (form1.TREJIMAVTO=1) and (form98.CheckBox3.Checked) then
   begin
    //ндфл заполнение
     if RGOD<=2022 then jvXpButton41Click(nil);
   end;


 form98.Free;





 // MessageDlg('Выполнено',mtInformation,[mbOk],0);

  DeleteNullKr;


  Query1.Locate('idp',uIdp,[loCaseInsensitive]);

  if (TYPEVYPLAT=0) then
   begin
    form1.kart.locate('nls',xnls,[locaseinsensitive]);
    form1.RaspredVicet;
   end;

  MessageDlg('Выполнено'+#13+sSoobVyp,mtInformation,[mbOk],0);



end;

procedure TForm123.N2Click(Sender: TObject);
var rtf:Boolean;
    s:String;
     xDat:TDate;
    xNote:String;
    i,xKod:integer;
    xNPVed, xIdSdoxod:Real;
    xid,xIdVypl:real;
    TProvAvans2:Integer;
    xnls:real;
begin

 query1.first;
 rtf:=true;
 while not query1.eof do
  begin
   if (not T_OWKOD) and (TYPEVYPLAT=0) and (Query1.FieldByNAme('PROVEDEN').asFloat<>1) then
     begin
      if DRound(Query1.FieldByNAme('AVANS').asFloat,2)<>0 then
       begin
        rtf:=false;
       end;
    end;
   query1.next;
  end;
 query1.first;

 if not rtf then
  begin
    MessageDlg('Итоговая выплата з/п не возможна, т.к. имеются не обработанные Авансы'+#13+
    '1) Для проведения определенных записей в данном списке сотрудников можно выбрать их (кнопка INS) и выполнить <Провести только отмеченные>'+#13+
    '2) Для проведения промежуточных выплат (отпускные, больничные, премиальные) установите Фильтр по кодам начислений и повторите данную операцию.'+
            ' В данном случае имейте ввиду, что по итогам месяца сумма не выплаченных начислений должно будет хватать, чтобы обработать Аванс',mtWarning,[mbOk],0);
        exit;
   end;


 if not ProvDatVypl then exit;


    //аванс проведен но нет выплаты !


 TProvAvans2:=0;

  Query1.First;
  while not Query1.Eof do
  begin
    vyplzp.Locate('IDP',Query1.fieldByName('IDP').asFloat,[loCaseInsensitive]);

   if form123.TYPEVYPLAT=0 then    //выплата з/п
     begin
      TProvAvans2:=Prov2Avans(query1.fieldbyname('nls').asFloat);

       if (vyplzp.FieldByNAme('proveden').asFloat<>1) then
        begin

         if TProvAvans2=1 then //проверка два аванса наличие
          begin
           MessageDlg('По сотруднику '+FZaprosFio(query1.fieldbyname('nls').asFloat)+' обнаружена выплата Аванса, но начислений за текущий период не хватает',mtWarning,[mbOk],0);
           EXIT;
          end;

         if TProvAvans2=2 then //проверка два аванса наличие
           begin
            { if MessageDlg('По сотруднику '+FZaprosFio(query1.fieldbyname('nls').asFloat)+' обнаружена выплата Аванса, необходимо перепровести выплаты'+#13+
                'Перепровести ?',mtWarning,[mbYes,mbNo],0) = mrNo then Exit;
            }
              form1.PereprovAvans(query1.fieldbyname('nls').asFloat,0,0);
          end;
       end;

     end;

    if (vyplzpSumma.Value<>0) and (vyplzp.FieldByNAme('proveden').asFloat<>1) and (TProvAvans2=0) then
      begin
        //showmessage(floattostr(vyplzpNls.Value));
       { if not ProvAvansKod(vyplzpNls.Value,Query1.fieldByname('FAM').asString+' '+Query1.fieldByname('IM').asString+' '
                                         +Query1.fieldByname('OT').asString) then EXIT;;
         }
      end;
    Query1.next;
   end;

 Query1.First;

 form98:=TForm98.Create(nil);
 form98.Label1.Caption:='Все сотрудники';

   datam.Query2.DataBaseName:=form1.DBDIR;
   datam.Query2.Close;
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('select max(id) from sdoxod');
   datam.Query2.Prepare;
   datam.Query2.Open;
   xIdSdoxod:=datam.Query2.Fields[0].asFloat;
   datam.Query2.Close;

   datam.Query2.DataBaseName:=form1.DBDIR;
   datam.Query2.Close;                           //номер плат.ведомости
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('select max(npved) from obrt2new where wg='+FloatToStr(RGod));
   datam.Query2.Prepare;
   datam.Query2.Open;
   xNPVed:=datam.Query2.Fields[0].asFloat+1;
   form228:=Tform228.Create(nil);
   form228.RxCalcEdit1.Value:=xNPVed;
   form228.ShowModal;
   if form228.TOk then xNPVed:=form228.RxCalcEdit1.Value else xNPVed:=-1;;
   form228.Free;
   if xNPVed=-1 then exit;


   datam.Query2.DataBaseName:=form1.DBDIR;
   datam.Query2.Close;
   datam.Query2.SQL.Clear;
   datam.Query2.SQL.Add('select n.kod,count(o.kr) from obrt1new o, nacisl n where pn<>1 and o.kod=n.kod ');
   datam.Query2.SQL.Add('and WG='+IntToStr(RGod)+' and WM='+IntToStr(RMes)+' group by n.kod');
   datam.Query2.Prepare;
   datam.Query2.Open;
   form98.RxCheckListBox1.Items.Clear;
   form98.rxCheckListBox1.Items.Add(FloatToStr(0)+'  - оклад');
    datam.Query2.First;
       while not datam.Query2.Eof do
           begin
              if form1.NACISL.Locate('KOD',datam.Query2.Fields[0].asFloat,[loCaseInsensitive]) then
                   form98.rxCheckListBox1.Items.Add(floattostr(form1.nacislkod.value)+'  - '+ansilowercase(form1.nacislName.Value));
            datam.Query2.Next;
           end;

 if form123.TYPEVYPLAT=0 then
    begin
     if form123.T_OWKOD then form98.CheckBox3.Checked:=False else form98.CheckBox3.Checked:=True;
     for i:=0 to form98.RxCheckListBox1.Items.Count-1 do  form98.RxCheckListBox1.Checked[i]:=True;
    end;

   if form123.TYPEVYPLAT=2 then
    begin
     form98.CheckBox1.Checked:=False;
     if form1.TREJIMAVTO<>1 then form98.CheckBox2.Enabled:=False;
     form98.CheckBox3.Enabled:=False;
     form98.CheckBox1.Visible:=False;
     form98.RxCheckListBox1.Visible:=False;
     form98.SpeedButton1.Visible:=false;
     form98.SpeedButton2.Visible:=false;
     for i:=0 to form98.RxCheckListBox1.Items.Count-1 do  form98.RxCheckListBox1.Checked[i]:=False;
    end;


 try
  if form123.TYPEVYPLAT=2 then xDat:=EncodeDAte(RGod, RMEs, trunc(form1.config2DV1.Value));
  if (form123.TYPEVYPLAT=0) and (RMEs<>12) then xDat:=EncodeDAte(RGod, RMEs+1, trunc(form1.config2DV2.Value));
  form123.fDat:=xDat;
 except
 end;   

 form98.ShowModal;
 form98.CheckBox3.Enabled:=True;
 form98.CheckBox2.Enabled:=True;



 xDat:=form98.DateEdit1.Date;
 xNote:=Trim(form98.Edit5.Text);
  if form98.TOk=0 then
   begin
    form98.free;
    exit;
   end;


 s:='Дата выплаты ';
 s:=s+ FormatDateTime('dd.mm.yyyy',xDat)+#13+'Примечание: '+xNote;

 {
 if MessageDlg('Провести расходные операции по карточкам Сотрудников ?'+#13+s+#13+'Проведены будут только непроведенные записи' ,mtWarning,[mbYes,mbNo],0)=mrNo
   then
    begin
     form98.Free;
     exit;
    end;
 }
 
 Query1.first;
 rtf:=True;
 while not Query1.Eof do
  begin
    if (Query1.fieldByName('KOD').asInteger<=0) and
            (Query1.fieldByName('SUMMA').asFloat<>0)  then rtf:=false;
   Query1.Next;
  end;
 if not rtf then
  begin
   MessageDlg('По некоторым сотрудникам'+#13+'не определено место выплаты',mtError,[mbOk],0);
   form98.Free;
   EXIT;
  end;

  Query1.First;
  while not Query1.Eof do
  begin
    vyplzp.Locate('IDP',Query1.fieldByName('IDP').asFloat,[loCaseInsensitive]);
    xnls:=Query1.fieldByName('nls').asFloat;


    if (Query1.FieldByName('SUMMA').asFloat<>0) and (vyplzp.FieldByName('PROVEDEN').asFloat=0) then
     begin
      form1.Obrt2.Append;
      xIdVypl:=form1.Obrt2ID.Value;
      form1.Obrt2.FieldByName('NLS').asFloat:=Query1.fieldByName('NLS').asFloat;
      form1.Obrt2.FieldByName('KR').asFloat:=Query1.fieldByName('SUMMA').asFloat;
      form1.Obrt2.FieldByName('WM').asFloat:=RMes;
      form1.Obrt2.FieldByName('WG').asFloat:=RGod;
      form1.Obrt2.FieldByName('KOD').asFloat:=Query1.fieldByName('KOD').asFloat;
      form1.obrt2.FieldByName('DATPROV').asDateTime:=xDat;
      form1.obrt2.FieldByName('NUMDOK').asFloat:=Query1.FieldByName('IDP').asFloat;
      form1.Obrt2.fieldByName('npved').asFloat:=xNPVed;
      form1.Obrt2.Fieldbyname('xs').asFloat:=TXS;

      if TYPEVYPLAT=0 then form1.Obrt2.Fieldbyname('TAVANS').asFloat:=0 else form1.Obrt2.Fieldbyname('TAVANS').asFloat:=1;

      form1.Obrt2.Post;
   //   showmessage(floattostr(xidvypl));

     if (query1.FieldByName('PROVEDEN').asFloat<>1) and (Query1.fieldByName('SUMMA').asFloat>0) then
      begin
  IF (form98.CheckBox2.Checked) and (form1.TREJIMAVTO=1) THEN
      BEGIN
      if (TYPEVYPLAT=0) then
      begin
       form1.FProvedenVypl2(query1.fieldbyname('nls').asFloat,xIdVypl,Query1.fieldByName('SUMMA').asFloat,xDat,2,true)  ;
      end
      else
        begin
         //аванс
        datam.Query1.close;
        datam.query1.sql.clear;
        datam.query1.sql.add('select max(id) from sdoxod');
        datam.query1.prepare;
        datam.query1.Open;
        xid:=datam.query1.Fields[0].asfloat;
        datam.query1.close;

         xid:=xid+1;
       {
         form1.sdoxod.Append;
         form1.sdoxod.fieldbyname('id').asFloat:=xid;
         form1.sdoxod.fieldbyname('tavans').asFloat:=1;
         form1.sdoxod.fieldbyname('TPRAV2').asFloat:=1;
         form1.sdoxod.fieldbyname('idvypl').asFloat:=xIdVypl;
         form1.sdoxod.fieldbyname('nls').asFloat:=query1.fieldbyname('nls').asFloat;
         form1.sdoxod.fieldbyname('kodnac').asFloat:=0;
         form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
         form1.sdoxod.fieldbyname('mes').asFloat:=RMes;
         form1.sdoxod.fieldbyname('god').asFloat:=RGod;
         form1.sdoxod.fieldbyname('sdoxod').asFloat:=query1.fieldByName('summa').asFloat;
         form1.sdoxod.fieldbyname('nalog').asFloat:=0;
         form1.sdoxod.post;
        } 
        end;
     END;
    end;



      if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asFloat,[loCaseInsensitive]) then
       begin
        vyplzp.Edit;
       {
        vyplzp.fieldByName('summa').asFloat:=vyplzpSUMMA.VAlue-Query1.fieldByName('SUMMA').asFloat;
        vyplzp.fieldByName('summa0').asFloat:=vyplzpSUMMA0.VAlue-Query1.fieldByName('SUMMA').asFloat;
       }
        vyplzp.fieldByName('proveden').asInteger:=1;
        vyplzp.fieldByName('g').asString:='';
        vyplzp.fieldByName('datdok').asDateTime:=xDat;
        if vyplzp.fieldbyname('sdoxod').asFloat<>0 then
           begin
            vyplzp.fieldByName('dat').asDateTime:=xDat;
             form1.sdoxod.Append;
             form1.sdoxod.fieldbyname('dat').asDateTime:=xDat;
             form1.sdoxod.fieldbyname('sdoxod').asFloat:=vyplzp.fieldbyname('sdoxod').asFloat;
             form1.sdoxod.fieldbyname('mes').asFloat:=vyplzp.fieldbyname('mes').asFloat;
             form1.sdoxod.fieldbyname('god').asFloat:=vyplzp.fieldbyname('god').asFloat;
             form1.sdoxod.fieldbyname('nalog').asFloat:=vyplzp.fieldbyname('snalog').asFloat;
             form1.sdoxod.fieldbyname('nls').asFloat:=vyplzp.fieldbyname('nls').asFloat;
             form1.sdoxod.fieldbyname('kodnac').asFloat:=vyplzp.fieldbyname('kodnac').asFloat;
             xIdSdoxod:=xIdsDoxod+1;
             form1.sdoxod.fieldbyname('id').asFloat:=xIdSdoxod;
             form1.sdoxod.Post;
             vyplzp.fieldbyname('idsdoxod').asFloat:=form1.sdoxod.FieldByName('id').asFloat;
            end;
         vyplzp.FieldByName('note').asString:='Вед.№'+FloatToStr(xNPVed)+', '+xNote;
        vyplzp.post;
       end;

     end;


  IF (form1.TREJIMAVTO<>1) and (form98.CheckBox2.Checked) THEN
   BEGIN
    { if form123.TYPEVYPLAT=0 then
      begin
       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('UPDATE glnew SET datoklad='+#39+FormatDatetime('dd.mm.yyyy',xD)+#39);
       datam.Query1.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
       datam.Query1.SQL.Add('and nls='+FloatToStr(Query1.fieldByName('NLS').asFloat));
       datam.Query1.SQL.Add('and datoklad<'+#39+'01.01.2000'+#39);
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
       datam.Query1.Close;


       datam.Query1.Close;
       datam.Query1.SQL.Clear;
       datam.Query1.SQL.Add('UPDATE obrt1new SET datprov='+#39+FormatDatetime('dd.mm.yyyy',xD)+#39);
       datam.Query1.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
       datam.Query1.SQL.Add('and nls='+FloatToStr(Query1.fieldByName('NLS').asFloat));
       datam.Query1.SQL.Add('and datprov<'+#39+'01.01.2000'+#39);
       datam.Query1.Prepare;
       datam.Query1.ExecSQL;
       datam.Query1.Close;

     end;
    }
     {дата выплаты дохода}
    
      for i:=0 to form98.RxCheckListBox1.Items.Count-1 do
       begin
        if form98.RxCheckListBox1.Checked[i] then
          begin
           xKod:=StrToInt(Trim(Copy(form98.rxCheckListBox1.Items[i],1,2)));

           if  xKod=0 then
            begin
             datam.Query2.Close;
             datam.Query2.SQL.Clear;
             datam.Query2.SQL.Add('UPDATE glnew SET datoklad='+#39+FormatDatetime('dd.mm.yyyy',form98.DateEdit1.Date)+#39);
             datam.Query2.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
             datam.Query2.SQL.Add('and nls='+FloatToStr(Query1.FieldByname('NLS').asFloat));
             if form98.CheckBox1.Checked then datam.Query2.SQL.Add('and (not datoklad>'+#39+'01.01.2005'+#39+')');
             datam.Query2.Prepare;
             datam.Query2.ExecSQL;
             datam.Query2.Close;
            end;

            if xKod<>0 then
              begin
               datam.Query2.Close;
               datam.Query2.SQL.Clear;
               datam.Query2.SQL.Add('UPDATE obrt1new SET datprov='+#39+FormatDatetime('dd.mm.yyyy',form98.DateEdit1.Date)+#39);
               datam.Query2.SQL.Add('where wm='+IntToStr(RMes)+' and wg='+IntToStr(RGod));
               datam.Query2.SQL.Add('and nls='+FloatToStr(Query1.FieldByname('NLS').asFloat));
               if form98.CheckBox1.Checked then datam.Query2.SQL.Add('and (not datprov>'+#39+'01.01.2005'+#39+')');
               datam.Query2.SQL.Add('and kod='+FloatToStr(xKod));
               datam.Query2.Prepare;
               datam.Query2.ExecSQL;
               datam.Query2.Close;
              end;
           end;
         end;

     END;

      if (TYPEVYPLAT=0) then
       begin
        form1.kart.locate('nls',xnls,[locaseinsensitive]);
        form1.RaspredVicet;
       end;


   Query1.Next;
  end;

 ZaprosV2(0);

 JvXPButton12Click(nil);



   DeleteNullKr;

 if (TYPEVYPLAT=2) then  MessageDlg('Выплата Аванса проведена, но не разнесена по налоговым регистрам'+#13+
              'Для окончательной выплаты аванса необходимо выполнить <Обработать аванс>, в модуле <Выплата заработной платы>'+#13+
                'Данная операция производится после проведения всех начислений перед итоговой выплатой заработной платы за отчетный период ',
                   mtInformation,[mbOk],0);


 if (TYPEVYPLAT=0) AND (form1.TREJIMAVTO<>1) AND (form98.CheckBox3.Checked) then
     if MessageDlg('Выполнено !'+#13+sSoobVyp+#13+'Перейти к процедуре заполнения перечисления НДФЛ ?',mtInformation,[mbYes,mbNo],0)=mrYes
          then Form1.N65Click(nil);;

 if (TYPEVYPLAT=0) AND (form1.TREJIMAVTO=1) and (form98.CheckBox3.Checked) then
   begin
    //ндфл заполнение

    MessageDlg('Выполнено !'+#13+sSoobVyp,mtInformation,[mbOk],0);

    if RGOD<=2022 then jvXpButton41Click(nil);



   end;



  form98.Free;


end;

procedure TForm123.JvXPButton8Click(Sender: TObject);
begin
 PopupMenu2.Popup(JvXPButton8.Left+JvXPButton8.Width+form123.left,JvXPButton8.Top+JvXPButton8.Height+form123.top)  ;
end;

procedure TForm123.N3Click(Sender: TObject);
var oldIdp:Real;
    xIdsDoxod:Real;
begin
 if Query1.RecordCount<=0 then exit;
 oldIdp:=Query1.Fieldbyname('idp').asFloat;
 if MessageDlg('Отменить проведение операции '+Query1.fieldByname('FAM').asString+' '+Query1.fieldByname('IM').asString+
             ' '+Query1.fieldByname('OT').asString+#13+
               'Сумма = '+FloatToStr(Query1.fieldByname('SUMMA').asFloat),mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 if vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then
  begin


   form1.kart.Locate('nls',vyplzpnls.Value,[loCaseInsensitive]);
   if form1.obrt2.Locate('nls;numdok',VarArrayOf([vyplzpNls.Value,vyplzpIDP.Value]),[loCAseInsensitive]) then
    begin
     form1.DelSDoxodObrt2(form1.obrt2ID.Value);
     form1.obrt2.Delete;
     MessageDlg('Отменена операция выплаты (удержания) !',mtInformation,[mbOk],0);
    end;

   vyplzp.edit;
   vyplzp.fieldbyname('proveden').asFloat:=0;
   vyplzp.fieldbyname('summa0').asFloat:=vyplzp.fieldbyname('summa').asFloat;  //могло быть проведено не вся сумма
   vyplzp.fieldbyname('numdok').asFloat:=0;
   vyplzp.fieldbyname('datdok').asDateTime:=EncodeDAte(1990,1,1);
   vyplzp.FieldByName('note').asString:='';
   vyplzp.post;
   xIdsDoxod:=vyplzp.FieldByName('idsdoxod').asFloat;
   if form1.sdoxod.Locate('id',xIdsdoxod,[loCaseInsensitive]) then form1.sdoxod.Delete;

   vyplzp.FlushBuffers;
   ZaprosV2(0);
  end;
  query1.Locate('idp',oldIdp,[loCaseInsensitive]);

   JvXPButton12Click(nil);







end;

procedure TForm123.N4Click(Sender: TObject);
var xidsdoxod:Real;
begin
 if Query1.RecordCount<=0 then exit;
 if MessageDlg('Отменить проведение операции по всем сотрудникам ?'
                 ,mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 Query1.First;
while not Query1.Eof do
begin
 if vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then
  begin

   form1.kart.Locate('nls',vyplzpnls.Value,[loCaseInsensitive]);
   if form1.obrt2.Locate('nls;numdok',VarArrayOf([vyplzpNls.Value,vyplzpIDP.Value]),[loCAseInsensitive]) then
    begin
     form1.DelSDoxodObrt2(form1.obrt2ID.Value);
     form1.obrt2.Delete;
    end;

   vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) ;

   vyplzp.edit;
   vyplzp.fieldbyname('proveden').asFloat:=0;
   vyplzp.fieldbyname('summa0').asFloat:=vyplzp.fieldbyname('summa').asFloat;  //могло быть проведено не вся сумма
   vyplzp.fieldbyname('numdok').asFloat:=0;
   vyplzp.fieldbyname('datdok').asDateTime:=EncodeDAte(1990,1,1);
   vyplzp.FieldByName('note').asString:='';
   vyplzp.post;
   xIdsDoxod:=vyplzp.FieldByName('idsdoxod').asFloat;
   if form1.sdoxod.Locate('id',xIdsdoxod,[loCaseInsensitive]) then form1.sdoxod.Delete;


  end;
query1.next;
end;
vyplzp.FlushBuffers;
ZaprosV2(0);

 JvXPButton12Click(nil);


end;

procedure TForm123.JvXPButton40Click(Sender: TObject);
begin
 Form1.N68Click(nil);
end;

procedure TForm123.JvXPButton41Click(Sender: TObject);
var oldRMes:Integer;
begin
 if (form1.TREJIMAVTO=1)  then
   begin

    
  //  Form1.JvXPBar4Items3Click(nil);
  //  exit;


    //ндфл заполнение
    form806:=Tform806.Create(nil);
    form806.Label1.Caption:='Переход к уплате НДФЛ за месяц';
    form806.ShowModal;
    oldRMes:=RMes;
    if form806.tMes=0 then
      begin
       form806.free;
       exit;
      end;
    RMEs:=form806.tMes;  
    form806.Free;
    form85:=Tform85.Create(Self);
    form89:=Tform89.Create(Self);
    form90:=Tform90.Create(Self);
    form85.TWSoob:=false; //не выводить сообщение про редактирование
    form85.OnShow(nil);
    Form85.N7Click(nil);
    Messagedlg('Уплату НДФЛ можно повторить за любой период через кнопку <Перечисление НДФЛ по сотрудникам>',mtInformation,[mbOk],0);
    form85.free;
    form89.free;
    form90.free;
    RMes:=oldRMes;
    exit;
   end;

 oldRMes:=RMes;
 Form1.N65Click(nil);
 RMes:=oldRMes;
end;

procedure TForm123.JvXPButton42Click(Sender: TObject);
begin
 form1.N17.Click;
end;

procedure TForm123.JvXPButton44Click(Sender: TObject);
var xProc:Real;
begin
 if Query1.RecordCount<=0 then exit;

 form228:=tform228.Create(nil);
 xProc:=100;
 form228.RxCalcEdit1.Value:=xProc;
 form228.Label1.Caption:='Процент выплаты от расчетной суммы';
 form228.ShowModal;
 xProc:=form228.RxCalcEdit1.Value;

 if not form228.TOk then
   begin
    form228.Free;
    exit;
   end;

 form228.free;

 Query1.First;
 while not Query1.Eof do
  begin
    if Query1.FieldByNAme('proveden').asFloat<>1 then
      begin
       if not vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then exit;
       vyplzp.Edit;
       vyplzp.FieldByNAme('summa0').asFloat:=DRound(vyplzpSUMMA0.Value*xProc/100,2);
       vyplzp.Post;
      end;

   Query1.Next;
  end;

  Query1.First;
  MessageDlg('Выполнено !',mtInformation,[mbOk],0);
  ZaprosV2(0);
  JvXpButton12Click(nil);

end;

procedure TForm123.RadioButton4Click(Sender: TObject);
begin
  ZaprosV2(0);
end;

procedure TForm123.vyplzpocerGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
var s:String;  
begin

 if form123.vyplzpDATDOK.Value<EncodeDAte(2014,1,1) then s:='6' else s:='3';
 Text:=s;
end;

procedure TForm123.N21998010620141Click(Sender: TObject);
var i:Integer;
    x,x0:Real;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
    xNum:Integer;
    npp:Integer;
    xNls:Real;
    E:OleVariant;
    fNameXls:String;
begin
  i:=0;
  x0:=0;
  try
   xNum:=StrToInt(Trim(Edit2.text));
  except
   xNum:=1;
  end;
  npp:=0;

    rtf:=false;
    x:=Query1.FieldByName('SUMMA').asFloat ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=1 then rtf:=True;
     end;

    if not rtf then
     begin
      MessageDlg('Невозможно формирование для данного места выплаты'+#13+'По данной операции в настройках Справочника удержаний  не выставлен признак, что документом является Расходный кассовый ордер',mtError,[mbOk],0);
      exit;
     end;

 form93.Query1.DatabaseName:=form1.DBDIR;

       npp:=npp+1;
       form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;

       vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
       if vyplzpNUMDOK.Value=0 then
        begin
         vyplzp.Edit;
         vyplzp.fieldByName('numdok').asFloat:=xNum;
        { if vyplzpDATDOK.Value<EncodeDAte(2000,1,1) then vyplzp.fieldByName('datdok').asDateTime:=xD;
        } vyplzp.Post;
         xNum:=xNum+1;
        end;
       form124.ShowModal;

       if not form124.TOk then exit;

      fNameXLS:=GetNameXLSn('ko2','ko2')+'.xls'  ;
      if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\ko2.xls',
               GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
      begin
       MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
       exit;
      end;

      E:=CreateOleObject('Excel.Application');
      E.Visible:=True;
      E.Application.WindowState:=2;
      E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXLS);

      try
       E.ActiveWindow.DisplayGridlines:=False;
      except
      end;


       E.ActiveWorkBook.Sheets.Item[1].Range['A5'].Value:=form1.config2NAME.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AI6'].Value:=form1.config2OKPO.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['O34'].Value:=form1.configRUKOVDOLJN.Value;
       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['AD34'].Value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';
       mainLib.GetFIO(form1.configGLBUH.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['S38'].Value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';
       E.ActiveWorkBook.Sheets.Item[1].Range['S58'].Value:=form1.configKASSIR.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['F25'].Value:=SumPropis(x);

       if not form124.CheckBox2.Checked then E.ActiveWorkBook.Sheets.Item[1].Range['AI12'].Value:=FormatdateTime('dd.mm.yyyy',vyplzpDATDOK.VAlue);
       if not form124.CheckBox1.Checked then E.ActiveWorkBook.Sheets.Item[1].Range['AD12'].Value:=FloatToStr(vyplzpNUMDOK.Value);

       E.ActiveWorkBook.Sheets.Item[1].Range['B53'].Value:=form1.kartNAMEDOC.Value+','+
              form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartPASS.Value+
              ', выдан '+form1.kartDATVYD.Text+' '+form1.kartVYDAN.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['J22'].Value:=vyplzpNOTE.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['G18'].Value:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['AD16'].Value:=FloatToStrF(x,ffNumber,12,2);
       E.Visible:=True;
       E.Application.WindowState:=-4137;
       E:=UnAssigned;



 Edit2.TExt:=IntToStr(xNum);



  xNls:=Query1.fieldByName('nls').asFloat;
  ZaprosV2(0);
  Query1.Locate('nls',xNls,[loCaseInsensitive]);



end;

procedure TForm123.JvXPButton28Click(Sender: TObject);
begin
 form423:=TForm423.Create(nil);
 form423.ShowModal;
 form423.Free;
end;

procedure TForm123.JvXPButton280Click(Sender: TObject);
begin
   form54:=TForm54.Create(Self);
   form54.JvXpButton5Click(nil);
   form54.Free;

end;

procedure TForm123.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (Ord(Key)=45) then JvXPButton99Click(nil);
 if Key=13 then JvXPButton23Click(Sender);

end;

procedure TForm123.N5Click(Sender: TObject);
begin
 ZaprosV2(1);
 N2Click(nil);
 ZaprosV2(0);
end;

procedure TForm123.JvXPButton99Click(Sender: TObject);
var xId:Real;
begin
  if Query1.RecordCount<=0 then exit;
  if not vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then exit;
  xId:=vyplzpIDP.Value;

  if (vyplzpProveden.Value=1) and (vyplzp.fieldbyname('G').asString='') then
   begin
    ShowMEssage('Нельзя выделить запись, которая уже проведена');
    exit;
   end;

  vyplzp.edit;
  if vyplzpG.Value='*' then vyplzp.fieldbyname('G').asString:='' else vyplzp.fieldbyname('G').asString:='*';
  vyplzp.post;
  Zaprosv2(0);
  Query1.Locate('IDP',xId,[loCaseInsensitive]);
  query1.next;

end;

procedure TForm123.FormClose(Sender: TObject; var Action: TCloseAction);
var tErr:boolean;
    F:TextFile;
begin

 //проверка
 query1.first;
 terr:=false;
 while not query1.eof do
  begin
    if vyplzp.Locate('IDP',query1.fieldByName('idp').asFloat,[loCaseInsensitive]) then
     begin
      form1.kart.Locate('nls',vyplzpnls.Value,[loCaseInsensitive]);
      if form1.obrt2.Locate('nls;numdok',VarArrayOf([vyplzpNls.Value,vyplzpIDP.Value]),[loCAseInsensitive]) then
       begin
        if form1.Obrt2TAVANS.Value<>1 then
         begin
          if (not FProvVypl(query1.fieldbyname('nls').asFloat,form1.obrt2.fieldbyname('id').asFloat,query1.fieldbyname('summa').asFloat))
              or (DRound(query1.fieldbyname('summa').asFloat-form1.obrt2Kr.Value,2)<>0) then
                begin
                    terr:=true;
                    MessageDlg('Что-то пошло не так по данной выплате, рекомендуется удалить и перепровести'+#13+
                      form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value+#13+
                       'Сумма в ведомости='+floattostrf(query1.fieldbyname('summa').asFloat,ffNumber,12,2)+#13+
                        'Сумма в карточке выплат='+floattostrf(form1.obrt2Kr.Value,ffNumber,12,2)+#13+
                        'Сумма регистрах 6-НДФЛ='+floattostrf(FSumVypl(query1.fieldbyname('nls').asFloat,form1.obrt2.fieldbyname('id').asFloat),ffNumber,12,2),mtWarning,[mbOk],0);
                end;
         end;
       end;
     end;
   query1.next;
  end;
 
 

 form1.statya.Active:=false;
 form1.PFormSave('form123',Form123);

 if CheckBox199.Visible then
  begin
   AssignFile(F,form1.DBDIR+'\uavans.txt');
   Rewrite(F);
   if (CheckBox199.Checked) then WriteLn(F,'UAVANS=TRUE') else WriteLn(F,'UAVANS=FALSE');
   CloseFile(F);
  end;

end;

procedure TForm123.JvXPButton268Click(Sender: TObject);
var xs:integer;
begin
 xs:=-1;
 if form1.statya.Locate('name',ComboBox1.Text,[loCaseInsensitive]) then xs:=Trunc(form1.statyaID.Value);
 if MessageDlg('Применить расчеты с учетом статьи'+#13+comboBox1.Text+#13+'id='+floattostr(xs),mtInformation,[mbYes,mbNo],0) = mrNo then exit;;
 TXS:=xs;

 JvXPButton6Click(nil);
 fName:='выплата заработной платы за '+ansilowerCase(namemes[RMes]);
 if TXS>=0 then fName:=fName+'/'+ComboBox1.Text+'/';


end;

procedure TForm123.JvXPButton59Click(Sender: TObject);
var xKod:Integer;
    xIdp:Real;
    x:real;
    i:Integer;
    rtf:Boolean;
    x0,x1,tNls:Real;
    pVicet,rVicet:real;
    skodv:string;
begin
// if MessageDlg('Добавить новую запись в реестр выплат ?',mtInformation,[mbYes,mbNo],0) = mrNo then exit;




 IF form123.TYPEVYPLAT=2 then     //аванс
 BEGIN
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

   form1.kart.First;
   i:=0;
   while not form1.kart.Eof do
    begin
      if form1.kartG.Value='*' then i:=i+1;
     form1.kart.next;
    end;
   if i=0 then
    begin
     form1.kart.locate('nls',tNls,[loCaseInsensitive]);
     form1.kart.edit;
     form1.kart.fieldbyname('g').asString:='*';
     form1.kart.post;
    end;
    
   form1.kart.first;
   while not form1.kart.Eof do
    begin
      if form1.kart.FieldByName('g').Value='*' then
       begin
             datam.kart2.Locate('nls',form1.kartnls.value,[loCaseinsensitive]);
             if datam.kart2TYPEAVANS.Value=1 then
              begin
               x0:=mainlib.LoadOklad(RMEs,RGod,form1.kartNLS.Value);
               x0:=DRound(x0*form1.kartAVANS.Value/100,2);   //всего аванс
             end
            else
             x0:=form1.kartAVANS.Value;    //всего аванс


           datam.Query1.Close;
           datam.Query1.SQL.Clear;
           datam.Query1.SQL.Add('select * from vyplzp where type=2 and nls='+floattostr(form1.kartnls.value));
           datam.query1.sql.add('and wm='+floattostr(RMes)) ;
           datam.query1.sql.add('and wg='+floattostr(RGod)) ;
           datam.query1.Prepare;
           datam.query1.open;
           x1:=0; //уже выплачено
           rtf:=true;
           while not datam.query1.eof do
            begin
             if datam.Query1.FieldByName('proveden').asInteger>=0 then
              begin
                MessageDlg(form1.kartFam.Value+' '+form1.kartim.Value+' '+form1.kartOt.Value+#13+'Уже есть запись по выплате аванса',mtWarning,[mbOk],0);
                Query1.Locate('idp',datam.Query1.FieldByName('idp').asFloat,[loCaseInsensitive]);
                rtf:=false;
              end;
              x1:=x1+datam.query1.fieldbyname('summa').asfloat;
            datam.query1.next;
          end;
          datam.query1.Close;
          x:=x0-x1;
          if x<0 then x:=0;
           { if x=0 then if
             MessageDlg('Согласно карточке сотрудника Аванс='+FloattostrF(x0,ffNumber,12,2)+#13+
              'В реестре выплат найдены суммы='+FloattostrF(x1,ffNumber,12,2)+#13+
               'Сумма аванса к доплате='+FloattostrF(x,ffNumber,12,2)+#13+
               'Все равно добавить запись ?',mtInformation,[mbYes,mbNo],0) = mrNo then exit;
                }
              if x=0 then x:=x0;

         //  if x=0 then EXIT;



           //  MessageDlg('Сумма аванса берется из карточки сотрудника,'+#13+'в случае если в карточке не установлено - берется 50% от оклада',mtInformation,[mbOk],0);

          xKod:=0;
          try
           xKod:=StrToInt(Trim(form1.kartD1C2.Value));
          except
           xKod:=0;
          end;

           {ПРОВЕРКА УВОЛЕННЫХ !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

         if rtf then
          begin

                    pVicet:=FVic1(RMes)+FKartVic2(RMes,skodv)+FIjdev2(RMes,skodv);  //право
                    datam.qtmp.close;
                    datam.qtmp.databasename:=form1.dbdir;
                    datam.qtmp.sql.clear;
                    datam.qtmp.sql.add('select sum(rvicet) from sdoxod where nls='+floattostr(form1.kartnls.value));
                    datam.qtmp.sql.add('and mes='+floattostr(RMes)+' and god='+floattostr(RGod));
                    datam.qtmp.prepare;
                    datam.qtmp.open;
                    rvicet:=datam.qtmp.fields[0].asfloat;
                    datam.qtmp.close;
                    pVicet:=pVicet-rvicet;
                    if pVicet<0 then pVicet:=0;

            datam.Query1.Close;
            datam.Query1.DatabaseName:=form1.DBDIR;
            datam.query1.sql.clear;
            datam.query1.SQL.add('select max(idp) from vyplzp');
            datam.Query1.Prepare;
            datam.Query1.Open;
            xIdp:=datam.Query1.Fields[0].AsFloat;
            datam.Query1.Close;

            xIdp:=xIdp+1;
            x:=0;
            vyplzp.Append;
            vyplzp.FieldByName('idp').asFloat:=xIdp;
            vyplzp.FieldByName('type').asFloat:=2;
            vyplzp.FieldByName('nls').asFloat:=form1.kartNls.Value;
            vyplzp.FieldByName('numdok').asFloat:=0;
            vyplzp.FieldByName('wm').asFloat:=RMEs;
            vyplzp.FieldByName('wg').asFloat:=RGod;
            vyplzp.FieldByName('summa').asFloat:=x;
            vyplzp.FieldByName('summa0').asFloat:=x;
            vyplzp.FieldByName('avdoxod').asFloat:=x;
             vyplzp.FieldByName('avvicet').asFloat:=DRound(pVicet,2);
            vyplzp.FieldByName('kod').asInteger:=xKod;
            vyplzp.FieldByName('proveden').asInteger:=0;
            vyplzp.post;
          end;
       end;

     form1.kart.Next;
    end;


 // form124.ShowModal;
 ZaprosV2(0);
 Query1.Locate('idp',xIdp,[loCaseInsensitive]);
 JvXPButton12Click(nil);


//   ZaprosV2(0);
 END; //аванс


 IF form123.TYPEVYPLAT=0 then     //выплата з/п
 BEGIN

   form1.kart.IndexName:='FAM';
   Form2.ShowModal;
   form1.kart.IndexName:='NLS';
   if form2.TOK=False then exit;

   datam.kart2.Locate('nls',form1.kartnls.value,[loCaseinsensitive]);

  x:=0;



 Query1.Close;
 Query1.DatabaseName:=form1.DBDIR;
 query1.sql.clear;
 query1.SQL.add('select max(idp) from vyplzp');
 Query1.Prepare;
 Query1.Open;
 xIdp:=Query1.Fields[0].AsFloat;
 Query1.Close;

//  MessageDlg('Сумма аванса берется из карточки сотрудника,'+#13+'в случае если в карточке не установлено - берется 50% от оклада',mtInformation,[mbOk],0);



   xKod:=0;
    try
     xKod:=StrToInt(Trim(form1.kartD1C.Value));
    except
     xKod:=0;
    end;

    {ПРОВЕРКА УВОЛЕННЫХ !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!}

         xIdp:=xIdp+1;
         vyplzp.Append;
         vyplzp.FieldByName('idp').asFloat:=xIdp;
         vyplzp.FieldByName('type').asFloat:=0;
         vyplzp.FieldByName('nls').asFloat:=form1.kartNls.Value;
         vyplzp.FieldByName('numdok').asFloat:=0;
         vyplzp.FieldByName('wm').asFloat:=RMEs;
         vyplzp.FieldByName('wg').asFloat:=RGod;
         vyplzp.FieldByName('summa').asFloat:=x;
         vyplzp.FieldByName('summa0').asFloat:=x;
         vyplzp.FieldByName('kod').asInteger:=xKod;
         vyplzp.FieldByName('proveden').asInteger:=0;
         vyplzp.post;

    form124.ShowModal;
 ZaprosV2(0);
 Query1.Locate('idp',xIdp,[loCaseInsensitive]);
 JvXPButton12Click(nil);

//   ZaprosV2(0);
 END; //аванс





end;

procedure TForm123.JvXPButton295Click(Sender: TObject);
var E:OleVariant;
    i,k:integer;
    x:Real;
    fNameXLS:String;
begin

 fNameXLS:=GetNameXLSn('Реестр_к_Выплате','rv')+'.xls';

 if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\к_выплате.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
   begin
    MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
    exit;
   end;



 E:=CreateOleObject('Excel.Application');
 E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXLS);
 E.Visible:=True;
 try
  E.ActiveWindow.DisplayGridlines:=False;
 except
 end;
 E.Application.WindowState:=2;


 if form123.TYPEVYPLAT=2 then
      E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:=form1.config2name.value+#10+'Суммы аванса к выплате, расчетный период '+namemes[rmes]+' '+floattostr(rgod)+'г.'
        else
          E.ActiveWorkBook.Sheets.Item[1].Range['A1'].Value:=form1.config2name.value+#10+'Суммы з/п к выплате, расчетный период '+namemes[rmes]+' '+floattostr(rgod)+'г.'
           ;


 Query1.first;
 i:=0;x:=0;
 while not query1.eof do
  begin
    if query1.fieldbyname('proveden').asInteger<>1 then
     begin
      i:=i+1;
      x:=x+Query1.FieldByName('summa').asFloat;
      E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(i+2)].Value:=i;
      E.ActiveWorkBook.Sheets.Item[1].Range['B'+inttostr(i+2)].Value:=Query1.fieldByName('Fam').asString+' '+Copy(Query1.fieldByName('Im').asString,1,1)+'.'+
             Copy(Query1.fieldByName('Ot').asString,1,1)+'.';
      E.ActiveWorkBook.Sheets.Item[1].Range['C'+inttostr(i+2)].Value:=query1.fieldbyname('summa').asfloat;
     end;
   query1.next;
  end;

   try
    for k:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A2','C'+IntToStr(i+2)].Borders[k].LineStyle:=1;
   except
   end;


      i:=i+1;
      E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(i+2)].Value:='';
      E.ActiveWorkBook.Sheets.Item[1].Range['B'+inttostr(i+2)].Value:='итого к выплате';
      E.ActiveWorkBook.Sheets.Item[1].Range['C'+inttostr(i+2)].Value:=x;


 i:=i+3;
 E.ActiveWorkBook.Sheets.Item[1].Range['A'+inttostr(i+2)].Value:='Сведения сформированы по состоянию на'+' '+Formatdatetime('dd.mm.yyyy',DAte())+'  '+TimeToStr(Time());

 try
  E.ActiveWorkbook.Sheets.Item[1].PageSetup.Zoom:=False;
  E.ActiveWorkbook.Sheets.Item[1].PageSetup.FitToPagesWide := 1;
  E.ActiveWorkbook.Sheets.Item[1].PageSetup.FitToPagesTall := 100;
 except
 end;

 E.Application.WindowState:=-4137;
 E:=UnAssigned;


end;

procedure TForm123.N010620141Click(Sender: TObject);
var i:Integer;
    x,x0:Real;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
    xNum:Integer;
    npp:Integer;
    xNls:Real;
    E:OleVariant;
    fNameXls:String;
begin
  i:=0;
  x0:=0;
  try
   xNum:=StrToInt(Trim(Edit2.text));
  except
   xNum:=1;
  end;
  npp:=0;

    rtf:=false;
    x:=Query1.FieldByName('SUMMA').asFloat ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=1 then rtf:=True;
     end;

    if not rtf then
     begin
      MessageDlg('Невозможно формирование для данного места выплаты'+#13+'По данной операции в настройках Справочника удержаний  не выставлен признак, что документом является Расходный кассовый ордер',mtError,[mbOk],0);
      exit;
     end;

 form93.Query1.DatabaseName:=form1.DBDIR;

       npp:=npp+1;
       form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;

       vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
       if vyplzpNUMDOK.Value=0 then
        begin
         vyplzp.Edit;
         vyplzp.fieldByName('numdok').asFloat:=xNum;
        { if vyplzpDATDOK.Value<EncodeDAte(2000,1,1) then vyplzp.fieldByName('datdok').asDateTime:=xD;
        } vyplzp.Post;
         xNum:=xNum+1;
        end;
       form124.ShowModal;

       if not form124.TOk then exit;

      fNameXLS:=GetNameXLSn('ko2','ko2')+'.xls'  ;
      if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\ko2014.xls',
               GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
      begin
       MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
       exit;
      end;

      E:=CreateOleObject('Excel.Application');
      E.Visible:=True;
      E.Application.WindowState:=2;
      E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXLS);

      try
       E.ActiveWindow.DisplayGridlines:=False;
      except
      end;


       E.ActiveWorkBook.Sheets.Item[1].Range['A2'].value:=form1.config2NAME.Value;
//       E.ActiveWorkBook.Sheets.Item[1].Range['A33']:=form1.configRUKOVDOLJN.Value;

       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['Y33'].value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';

       mainLib.GetFIO(form1.configGLBUH.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['Y31'].value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';

       E.ActiveWorkBook.Sheets.Item[1].Range['F16'].value:=SumPropis(x);

       if not form124.CheckBox2.Checked then E.ActiveWorkBook.Sheets.Item[1].Range['AI6'].value:=FormatdateTime('dd.mm.yyyy',vyplzpDATDOK.VAlue);
       if not form124.CheckBox1.Checked then E.ActiveWorkBook.Sheets.Item[1].Range['AD6'].value:=FloatToStr(vyplzpNUMDOK.Value);

       E.ActiveWorkBook.Sheets.Item[1].Range['U28'].value:=form1.kartNAMEDOC.Value+','+
              form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartPASS.Value+
              ', выдан '+form1.kartDATVYD.Text+' '+form1.kartVYDAN.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['K21'].value:=vyplzpNOTE.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['G12'].value:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['V9'].value:=FloatToStrF(x,ffNumber,12,2);
       E.Visible:=True;
       E.Application.WindowState:=-4137;
       E:=UnAssigned;


 Edit2.TExt:=IntToStr(xNum);



  xNls:=Query1.fieldByName('nls').asFloat;
  ZaprosV2(0);
  Query1.Locate('nls',xNls,[loCaseInsensitive]);



end;

procedure TForm123.JvXPButton222Click(Sender: TObject);
begin
 form701:=Tform701.Create(nil);
 form701.ShowModal;
 form701.Free;
end;

procedure TForm123.JvXPButton233Click(Sender: TObject);
var xNls,xIdp:Real;
    s:string;
    x:real;
    xMes:Word;
begin

 if Query1.FieldByNAme('proveden').asFloat=1 then
  begin
    MessageDlg('Исправление невозможно, уже проведено !'+#13+
           'Снимите сначала признак проведения (отменить проведение)'+#13+
            'затем только возможная корректировка',mtWarning,[mbOk],0);
    exit;
  end;
 xNls:=Query1.FieldByNAme('nls').asFloat;
 if not form1.kart.Locate('nls',xNls,[loCaseInsensitive]) then exit;

 xIdp:=Query1.fieldByName('idp').asFloat;
 if not vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then exit;


 form806:=Tform806.Create(nil);
 form806.Caption:='Месяц, в котором начислены Доход и НДФЛ';
 form806.Label1.Caption:=form1.kartfam.value+' '+form1.kartim.value+' '+form1.kartot.value;

 form806.ShowModal;
 if form806.Tok then
  begin
   xMes:=form806.tMes;
  end
   else
   begin
    form806.free;
    exit;
   end ;
 form806.free;



 if Query1.RecordCount<=0 then exit;

 Form_58.Button1Click(nil);   //заполнение snalog по начислениям

 form123.vdoxod.DatabaseName:=form1.DBDIR;
 if form123.vdoxod.Active then form123.vdoxod.Active:=false;

 if form123.vdoxod.Exists then form123.vdoxod.DeleteTable;

 form123.vdoxod.CreateTable;

 form123.vdoxod.Active:=true;







 form1.kart.Locate('nls',xNls,[loCaseInsensitive]) ;
 ZapolnMas;
 s:='';
 datam.Query1.close;
 datam.query1.SQL.clear;
 datam.Query1.SQL.Add('select * from glnew where nls='+floattostr(xNls));
 datam.Query1.SQL.Add('and wm='+floattostr(xMes));
 datam.Query1.SQL.Add('and wg='+floattostr(RGod));
 datam.Query1.Prepare;
 datam.Query1.Open;
 if datam.Query1.RecordCount=1 then
  begin
   vdoxod.append;
  // Showmessage(floattostr(xmes)+#13+floattostr(DOklad[xMes]));
   vdoxod.fieldbyname('kod').asFloat:=0;
   vdoxod.fieldbyname('name').asString:='Оклад/тариф';
   vdoxod.fieldbyname('summa').asFloat:=DOklad[xMes]+Dround(Doklad[xMes]*form1.configRK.Value/100,2);
   vdoxod.fieldbyname('snalog').asFloat:=datam.Query1.FieldByNAme('snalog').asFloat;
   vdoxod.post;
  end;

 datam.Query1.close;
 datam.query1.SQL.clear;
 datam.Query1.SQL.Add('select * from obrt1new where nls='+floattostr(xNls));
 datam.Query1.SQL.Add('and wm='+floattostr(xMEs));
 datam.Query1.SQL.Add('and wg='+floattostr(RGod));
 datam.Query1.Prepare;
 datam.Query1.Open;
 datam.Query1.first;
 while not datam.Query1.Eof do
  begin
   form1.NACISL.Locate('kod',datam.Query1.FieldbyName('kod').asInteger,[loCaseInsensitive]);
   x:=datam.Query1.FieldbyName('kr').asFloat;
   if form1.nacisl.FieldbyNAme('rk').asBoolean then x:=x+Dround(x*form1.configRK.Value/100,2);
   vdoxod.append;
   vdoxod.fieldbyname('kod').asFloat:=form1.NACISLKOd.Value;
   vdoxod.fieldbyname('name').asString:=form1.NACISLNAME.Value;
   vdoxod.fieldbyname('summa').asFloat:=x;
   vdoxod.fieldbyname('snalog').asFloat:=datam.Query1.FieldByNAme('snalog').asFloat;
   vdoxod.post;
   datam.Query1.Next;
  end;

 form805:=Tform805.Create(nil);
 form805.tMes:=xMes;
 form805.tGod:=RGod;
 form805.tNls:=xNls;
 form805.ShowModal;

 if form805.TOk then
  begin
   vyplzp.Edit;
   vyplzp.fieldbyname('summa0').asFloat:=vdoxod.FieldByName('oi').asFloat-vdoxod.FieldByName('oi2').asFloat;;
   vyplzp.fieldbyname('god').asFloat:=RGod;
   vyplzp.fieldbyname('mes').asFloat:=xMes;
   vyplzp.fieldbyname('kodnac').asFloat:=vdoxod.FieldByName('kod').asFloat;
   vyplzp.fieldbyname('sdoxod').asFloat:=vdoxod.FieldByName('oi').asFloat;
   vyplzp.fieldbyname('snalog').asFloat:=vdoxod.FieldByName('oi2').asFloat;
   vyplzp.post;
  end;

 form805.Free;

 form123.vdoxod.Active:=false;

 ZaprosV2(0);
 Query1.Locate('idp',xIdp,[loCaseInsensitive]);
 JvXPButton12Click(nil);

end;

procedure TForm123.CheckBox1Click(Sender: TObject);
begin
 Zaprosv2(0);
 JvXPButton12Click(nil);
end;

procedure TForm123.JvXPButton299Click(Sender: TObject);
begin
 ShellExecute(handle,'open',PChar(form1.DBCurr+'\vyplat.doc'),nil,nil,SW_SHOW);
end;

procedure TForm123.Button1Click(Sender: TObject);
var xIdp,xNls:Real;
    idel:integer;
begin

//**********************

 Query1.first;
 idel:=0;
 while not Query1.Eof do
  begin
    if Query1.FieldByName('proveden').asFloat=1 then
      begin
         xIdp:=query1.fieldbyname('idp').asFloat;
         xNls:=query1.fieldbyname('nls').asFloat;
         form1.kart.locate('nls',xNls,[loCaseInsensitive]);
          if not form1.obrt2.Locate('nls;numdok',VarArrayOf([xNls,xIdp]),[loCAseInsensitive]) then
             begin
              if MessageDlg(query1.fieldbyname('fam').asString+' '+query1.fieldbyname('im').asString+' '+query1.fieldbyname('ot').asString+#13+
                 'Обнаружена недействующая ссылка на выплату (выплата в карточке сотрудника отсутствует)'+#13+
                  'Сумма='+floattostr(query1.fieldbyname('summa').asFloat)+#13+
                    'Удалить ссылку на выплату ?',mtError,[mbYes,mbNo],0) = mrYes
                   then
                     begin
                      if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then vyplzp.delete;
                      idel:=idel+1;
                     end;
             end
               else
              begin
               if DRound(form1.obrt2KR.Value-query1.fieldbyname('summa').asFloat,2)<>0 then
                 MessageDlg(query1.fieldbyname('fam').asString+' '+query1.fieldbyname('im').asString+' '+query1.fieldbyname('ot').asString+#13+
                 'Несоответствие ссылки на выплату в карточке сотрудника'+#13+'Возможно была ручная корректировка после выплаты'+#13+
                  'Сумма по реестру='+floattostrf(query1.fieldbyname('summa').asFloat,ffNumber,12,2)+#13+
                    'Выплата в карточке сотрудника='+floattostrf(form1.obrt2KR.Value,ffNumber,12,2),mtError,[mbOk],0);
               if form1.obrt2DATPROV.Value<>query1.fieldbyname('datdok').asDateTime then
                MessageDlg(query1.fieldbyname('fam').asString+' '+query1.fieldbyname('im').asString+' '+query1.fieldbyname('ot').asString+#13+
                 'Несоответствие ссылки на выплату в карточке сотрудника'+#13+'Возможно была ручная корректировка после выплаты'+#13+
                  'Дата по реестру='+FormatDatetime('dd.mm.yyyy',query1.fieldbyname('datdok').asdateTime)+#13+
                    'Дата в карточке сотрудника='+FormatDateTime('dd.mm.yyyy',form1.obrt2DATPROV.Value),mtError,[mbOk],0) ;
              end;
      end;
   Query1.Next;
  end;

 if idel<>0 then
   begin
    ZaprosV2(0);
    JvXPButton12Click(nil);
   end;
  Query1.first;

end;

procedure TForm123.RadioGroup2Click(Sender: TObject);
begin
 zaprosV2(0);
  JvXPButton12Click(nil);
end;

procedure TForm123.JvXPButton258Click(Sender: TObject);
begin
 Form1.JvXPBar4Items3Click(nil);
end;

procedure TForm123.JvXPButton_222Click(Sender: TObject);
var s:String;
begin
  s:=form1.DBCurr+'\platpor.exe';

  if mainlib.GetProcessByEXE('platpor.exe') = 0 then
    ShellExecute(0, 'open', PChar(s), PChar(form1.DBDIR), nil, SW_SHOWNORMAL)
  else
    ShowMessage('Модуль Платежные поручения уже запущен');

end;

procedure TForm123.JvXPButton298Click(Sender: TObject);
var xIdp:Real;
    oldnumdok,newnumdok:Real;
    xNote,xNote1,xNote2,s1:String;
    i:integer;
    rtf:boolean;
begin
 newnumdok:=999;
 if Query1.RecordCount<=0 then exit;
 xIdp:=Query1.fieldByName('idp').asFloat;
 if Query1.Fieldbyname('proveden').asInteger<>1 then
  begin
   MessageDlg('Выплата должна быть проведена для изменения номера платежной ведомости и комментария',mtWarning,[mbOk],0);
   exit;
  end;
 vyplzp.Locate('IDP',xIdp,[loCaseInsensitive])  ;
 xNote:=vyplzp.Fieldbyname('note').asString;
 xNote1:='';
 xNote2:='';  rtf:=false;
 for i:=1 to Length(xNote) do
  begin
   s1:=copy(xNote,i,1);
   if s1=',' then rtf:=true;
   if not rtf then xNote1:=xNote1+s1 else xNote2:=xNote2+s1;
  end;

 xNote1:=FAc(xNote1);
 xNote2:=Trim(Copy(xNote2,2,150)); // , отрезать
 Form11125:=Tform11125.Create(nil);
 form11125.Edit1.Text:=xNote1;
 form11125.Edit3.Text:=xNote1;
 form11125.Edit2.Text:=xNote2;
 form11125.Edit4.Text:=xNote2;
 form11125.ShowModal;
 xNote1:=form11125.Edit3.Text;
 xNote2:=form11125.Edit4.Text;
 rtf:=form11125.TOk;
 form11125.Free;
 if not rtf then exit;

 Query1.First;
 while not Query1.Eof do
  begin
   xIdp:=Query1.fieldByName('idp').asFloat;
   if (Query1.Fieldbyname('proveden').asInteger=1) and (vyplzp.Locate('IDP',xIdp,[loCaseInsensitive])) then
    begin
   // ShowMessage(xNote+#13+vyplzpNote.Value);
    form1.kart.locate('nls',vyplzpNls.Value,[loCaseInsensitive]);
     if (form1.Obrt2.Locate('NUMDOK',vyplzpIDP.Value,[loCAseInsensitive])) and (trim(vyplzpNote.value)=Trim(xNote)) then
      begin
       form1.obrt2.edit;
      // form1.obrt2.FieldByNAme('NPVED').asFloat:=StrToFloat(xNote1);
       form1.obrt2.FieldByNAme('NPVED2').asString:=xNote1;
       form1.obrt2.post;
       vyplzp.edit;
       vyplzp.fieldbyname('note').asString:='Вед.№'+xNote1+', '+xNote2;
       vyplzp.post;
      end;
    end;
   Query1.next;
  end;
 ZaprosV2(0);
 MessageDlg('Выполнено !',mtInformation,[mbOk],0);

end;

procedure TForm123.N21Click(Sender: TObject);
var i:Integer;
    x,x0:Real;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
    xNum:Integer;
    npp:Integer;
    xNls:Real;
    E:OleVariant;
    fNameXls:String;
    xKodU:Real;
    xDatDok:TDAte;
    xNote,xNote2:String;
begin
   i:=0;
  x0:=0;

  xNum:=Trunc(Query1.Fieldbyname('numdok').asFloat);

  if Trunc(xNum)=0 then
   begin
     try
      xNum:=StrToInt(Trim(Edit2.text));
      form97:=Tform97.Create(nil);
      form97.caption:='Ввод номера';
      form97.rxLAbel2.Visible:=false;
      form97.Dateedit1.visible:=false;
      form97.rxLAbel1.CAption:='Номер документа';
      form97.rxCAlcEdit1.Value:=xNum;
      form97.Showmodal;
      if form97.Tok=1 then
        begin
         xNum:=Trunc(form97.rxCAlcEdit1.Value);
         form97.Free;
        end
          else
         begin
          form97.free;
          exit;
         end;

     except
      xNum:=1;
     end;
    xNum:=xNum+1;
     Edit2.TExt:=IntToStr(xNum);
   end;



  npp:=0;

    rtf:=false;
    vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
    xKodU:=Query1.fieldByName('kod').asFloat;
    xDatDok:=Query1.fieldbyname('datdok').asDAtetime;
    xNote:=Trim(Query1.fieldbyname('note').asString);
    xNote2:=vyplzpNOTE.Value;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=1 then rtf:=True;
     end;

    if not rtf then
     begin
      MessageDlg('Невозможно формирование для данного места выплаты'+#13+'По данной операции в настройках Справочника удержаний  не выставлен признак, что документом является Расходный кассовый ордер',mtError,[mbOk],0);
      exit;
     end;

 form93.Query1.DatabaseName:=form1.DBDIR;
 xNls:=Query1.fieldByName('nls').asFloat;



       npp:=npp+1;
  //     form1.kart.Locate('NLS',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]) ;

  Query1.First;
  x:=0;
  while not query1.eof do
   begin
    if (query1.fieldbyname('proveden').asInteger=1) and (xKodU=Query1.fieldByName('kod').asFloat)
            and (xDatDok=Query1.fieldbyname('datdok').asDAtetime) and (xNote=Trim(Query1.fieldbyname('note').asString)) then
     begin
      vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
       if vyplzpNUMDOK.Value=0 then
        begin
         vyplzp.Edit;
         vyplzp.fieldByName('numdok').asFloat:=xNum;
         vyplzp.Post;
        end;
        if vyplzp.fieldbyname('numdok').asFloat=xNum then  x:=x+vyplzpsumma.Value;
       end;
    query1.next;
   end;


     {  form124.ShowModal;

       if not form124.TOk then exit;

    }
      fNameXLS:=GetNameXLSn('ko2','ko2')+'.xls'  ;
      if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\ko2.xls',
               GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
      begin
       MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
       exit;
      end;

      E:=CreateOleObject('Excel.Application');
      E.Visible:=True;
      E.Application.WindowState:=2;
      E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXLS);

      try
       E.ActiveWindow.DisplayGridlines:=False;
      except
      end;


       E.ActiveWorkBook.Sheets.Item[1].Range['A5'].value:=form1.config2NAME.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['AI6'].value:=form1.config2OKPO.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['O34'].value:=form1.configRUKOVDOLJN.Value;
       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['AD34'].value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';
       mainLib.GetFIO(form1.configGLBUH.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['S38'].value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';
       E.ActiveWorkBook.Sheets.Item[1].Range['S58'].value:=form1.configKASSIR.Value;
       E.ActiveWorkBook.Sheets.Item[1].Range['F25'].value:=SumPropis(x);

       E.ActiveWorkBook.Sheets.Item[1].Range['AI12'].value:=FormatdateTime('dd.mm.yyyy',xDATDOK);
       E.ActiveWorkBook.Sheets.Item[1].Range['AD12'].value:=FloatToStr(xNum);

     {
       E.ActiveWorkBook.Sheets.Item[1].Range['B53']:=form1.kartNAMEDOC.Value+','+
              form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartPASS.Value+
              ', выдан '+form1.kartDATVYD.Text+' '+form1.kartVYDAN.Value;
      }

       E.ActiveWorkBook.Sheets.Item[1].Range['J22'].value:=xNote2;
     //  E.ActiveWorkBook.Sheets.Item[1].Range['G18']:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['AD16'].value:=FloatToStrF(x,ffNumber,12,2);
       E.Visible:=True;
       E.Application.WindowState:=-4137;
       E:=UnAssigned;


  ZaprosV2(0);
  Query1.Locate('nls',xNls,[loCaseInsensitive]);



end;

procedure TForm123.N010620142Click(Sender: TObject);
var i:Integer;
    x,x0:Real;
    xFam,xIm,xOt:String;
    rtf:boolean;
    xDok:Integer;
    xNum:Integer;
    npp:Integer;
    xNls:Real;
    E:OleVariant;
    fNameXls:String;
    xKodU:Real;
    xDatDok:TDAte;
    xNote,xNote2:String;
begin
 i:=0;
  x0:=0;

  xNum:=Trunc(Query1.Fieldbyname('numdok').asFloat);
  if Trunc(xNum)=0 then
   begin
     try
      xNum:=StrToInt(Trim(Edit2.text));
      form97:=Tform97.Create(nil);
      form97.caption:='Ввод номера';
      form97.rxLAbel2.Visible:=false;
      form97.Dateedit1.visible:=false;
      form97.rxLAbel1.CAption:='Номер документа';
      form97.rxCAlcEdit1.Value:=xNum;
      form97.Showmodal;
      if form97.Tok=1 then
        begin
         xNum:=Trunc(form97.rxCAlcEdit1.Value);
         form97.Free;
        end
          else
         begin
          form97.free;
          exit;
         end;

     except
      xNum:=1;
     end;
    xNum:=xNum+1;
     Edit2.TExt:=IntToStr(xNum);
   end;


  npp:=0;

    rtf:=false;
    vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
    xKodU:=Query1.fieldByName('kod').asFloat;
    xDatDok:=Query1.fieldbyname('datdok').asDAtetime;
    xNote:=Trim(Query1.fieldbyname('note').asString);
    xNote2:=vyplzpNOTE.Value;

    x:=0 ;
    if form1.uderj.Locate('kod',Query1.fieldByName('kod').asFloat,[loCaseInsensitive]) then
     begin
      try
       xDok:=StrToInt(Trim(form1.uderjDBSPRAV.Value))     ;
      except
       xDok:=0;
      end;
      if xDok=1 then rtf:=True;
     end;

    if not rtf then
     begin
      MessageDlg('Невозможно формирование для данного места выплаты'+#13+'По данной операции в настройках Справочника удержаний  не выставлен признак, что документом является Расходный кассовый ордер',mtError,[mbOk],0);
      exit;
     end;

 xNls:=Query1.fieldByName('nls').asFloat;

 form93.Query1.DatabaseName:=form1.DBDIR;


 

  Query1.First;
  x:=0;
  while not query1.eof do
   begin
    if (query1.fieldbyname('proveden').asInteger=1) and (xKodU=Query1.fieldByName('kod').asFloat)
            and (xDatDok=Query1.fieldbyname('datdok').asDAtetime) and (xNote=Trim(Query1.fieldbyname('note').asString)) then
     begin
      vyplzp.Locate('IDP',Query1.fieldByName('idp').asFloat,[loCaseInsensitive]);
       if vyplzpNUMDOK.Value=0 then
        begin
         vyplzp.Edit;
         vyplzp.fieldByName('numdok').asFloat:=xNum;
         vyplzp.Post;
        end;
        if vyplzp.fieldbyname('numdok').asFloat=xNum then x:=x+vyplzpsumma.Value;
       end;
    query1.next;
   end;
   

      fNameXLS:=GetNameXLSn('ko2','ko2')+'.xls'  ;
      if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\ko2014.xls',
               GetCurrentDir()+'\TMP_XLS\'+fNameXLS) then
      begin
       MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
       exit;
      end;

      E:=CreateOleObject('Excel.Application');
      E.Visible:=True;
      E.Application.WindowState:=2;
      E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXLS);

      try
       E.ActiveWindow.DisplayGridlines:=False;
      except
      end;


       E.ActiveWorkBook.Sheets.Item[1].Range['A2'].Value:=form1.config2NAME.Value;
//       E.ActiveWorkBook.Sheets.Item[1].Range['A33']:=form1.configRUKOVDOLJN.Value;

       mainLib.GetFIO(form1.configRUKOVOD.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['Y33'].Value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';

       mainLib.GetFIO(form1.configGLBUH.Value,xFam,xIm,xOt);
       E.ActiveWorkBook.Sheets.Item[1].Range['Y31'].Value:=xFam+' '+Copy(xIm,1,1)+'.'+Copy(xOt,1,1)+'.';

       E.ActiveWorkBook.Sheets.Item[1].Range['F16'].Value:=SumPropis(x);

       E.ActiveWorkBook.Sheets.Item[1].Range['AI6'].Value:=FormatdateTime('dd.mm.yyyy',xDATDOK);;
      E.ActiveWorkBook.Sheets.Item[1].Range['AD6'].Value:=FloatToStr(xNum);

     {
       E.ActiveWorkBook.Sheets.Item[1].Range['U28']:=form1.kartNAMEDOC.Value+','+
              form1.kartSER1.Value+' '+form1.kartSER2.Value+' '+form1.kartPASS.Value+
              ', выдан '+form1.kartDATVYD.Text+' '+form1.kartVYDAN.Value;
       }

       E.ActiveWorkBook.Sheets.Item[1].Range['K21'].Value:=xNote2;
   //    E.ActiveWorkBook.Sheets.Item[1].Range['G12']:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;

       E.ActiveWorkBook.Sheets.Item[1].Range['V9'].Value:=FloatToStrF(x,ffNumber,12,2);
       E.Visible:=True;
       E.Application.WindowState:=-4137;
       E:=UnAssigned;


 

  ZaprosV2(0);
  Query1.Locate('nls',xNls,[loCaseInsensitive]);

end;

procedure TForm123.JvXPButton9Click(Sender: TObject);
begin
 form1160:=Tform1160.Create(nil);
 form1160.ShowModal;
 form1160.Free;
end;

procedure TForm123.JvXPButton555Click(Sender: TObject);
var
  E:OleVariant;
  NPVed, fNameXLS,s:String;
  i:integer;
begin
  if Query1.RecordCount<=0 then
  begin
   MessageDlg('Отсутствует проведенние выплат',mtInformation,[mbOk],0);
   exit;
  end;

   if Query1.FieldByName('proveden').asInteger<>1 then
    begin
       MessageDlg('Данная выплата не проведена',mtInformation,[mbOk],0);
       exit;                                                                      
    end;

   nPVed:='';
   if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]) then nPVed:=vyplzpNote.Value
    else
      begin
      // MessageDlg('Не найдена выплата в карточке сотрудника по данной строке',mtInformation,[mbOk],0);
      // exit;
      end;


  if MessageDlg('Сформировать '+npVed,mtInformation,[mbYes,mbNo],0) = mrNo then exit;

   fNameXls:=mainlib.GetNameXLSn('РеестрБанк','rb')+'.xls';
 if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\reebank.xls',form1.DBCurr+'\TMP_XLS\'+fNameXLS) then
   begin
    MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
    exit;
   end;

 E:=CreateOleObject('Excel.Application');
 E.WorkBooks.Open(form1.DBCurr+'\TMP_XLS\'+fNameXLS);

 E.Visible:=True;
 E.Application.WindowState:=2;

 s:=form1.config2NAME.VAlue+', ИНН '+form1.config2INN.VAlue;
 if trim(form1.config2KPP.VAlue)<>'' then s:=s+' ОГРН '+form1.config2OGRN.VAlue else s:=s+' ОГРНИП '+form1.config2OGRN.VAlue;
 E.ActiveWorkBook.Sheets.Item[1].Range['B2'].Value:=s;

 E.ActiveWorkBook.Sheets.Item[1].Range['B4'].Value:='В Банк: '+form1.config2BANKNAME.Value+' БИК '+form1.config2BIK.Value+' к/сч '+form1.config2KSC.Value;

 E.ActiveWorkBook.Sheets.Item[1].Range['I144'].Value:=form1.configGLBUH.Value;
 E.ActiveWorkBook.Sheets.Item[1].Range['I143'].Value:=form1.configRUKOVOD.Value;
 E.ActiveWorkBook.Sheets.Item[1].Range['B143'].Value:=form1.configRUKOVDOLJN.Value;


 E.ActiveWorkBook.Sheets.Item[1].Range['B3'].Value:='Реестр №        от '+FormatDateTime('dd.mm.yyyy',vyplzp.FieldByName('DATDOK').asDateTime);
 E.ActiveWorkBook.Sheets.Item[1].Range['F7'].Value:= nPved;


 Query1.First;
 i:=0;
 while not Query1.Eof do
  begin
    vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]);
    if (Query1.FieldByName('proveden').asInteger=1) and (nPved=vyplzpNote.Value) then
     begin

      form1.kart.Locate('nls',Query1.fieldByName('nls').asFloat,[loCaseInsensitive]);
      i:=i+1;
      E.ActiveWorkBook.Sheets.Item[1].Range['B'+Inttostr(i+11)].Value:=i;
      E.ActiveWorkBook.Sheets.Item[1].Range['C'+Inttostr(i+11)].Value:=form1.kartFAM.Value+' '+form1.kartIM.Value+' '+form1.kartOT.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['J'+Inttostr(i+11)].Value:=Query1.FieldByNAme('summa').asFloat   ;
      E.ActiveWorkBook.Sheets.Item[1].Range['H'+Inttostr(i+11)].Value:=form1.kartLSC.Value;
      E.ActiveWorkBook.Sheets.Item[1].Range['I'+Inttostr(i+11)].Value:=form1.kartBANKNAME.Value+' БИК '+form1.kartBANKBIK.Value+' к/сч '+form1.kartBANKKSC.Value;
     end;
   Query1.Next;
  end;

 try
  if not form1.TCheckOO then
     E.ActiveWorkBook.Sheets.Item[1].Range['A'+IntToStr(i+12),'M141'].EntireRow.Delete(EmptyParam)
      else
        E.ActiveWorkBook.Sheets.Item[1].Rows[Floattostr(i+12)+':141'].Hidden:=True;;
 except
 end;

 E.Application.WindowState:=-4137;
 E:=UnAssigned;

   
end;

function TForm123.Prov2Avans(xNls:Real):Integer;     //проверяет два и более аванса наличие   0 - НЕТ, 1 - ДА, нельзя провести, 2- ДА, можно провести
var rtf:Boolean;
    n1,n2:integer;
    xNacisl,sNacisl:Real;
    x:Real;
    nItog:Integer;
begin
 rtf:=false;
  datam.qtmpstaj.close;
  datam.qtmpstaj.DatabaseName:=form1.DBDIR;
  datam.qtmpstaj.sql.clear;
  datam.qtmpstaj.sql.add('select * from sdoxod where nls='+floattostr(xNls));
  datam.qtmpstaj.sql.add('and mes='+floattostr(RMes));
  datam.qtmpstaj.sql.add('and god='+floattostr(RGod));
  datam.qtmpstaj.Prepare;
  datam.qtmpstaj.Open;
  rtf:=false; n1:=0; n2:=0;
  datam.qtmpstaj.first;
  while not datam.qtmpstaj.Eof do
   begin
    if datam.qtmpstaj.fieldbyname('tprav2').asFloat=1 then n1:=n1+1;
    if datam.qtmpstaj.fieldbyname('tavans').asFloat<>1 then n2:=n2+1;
    datam.qtmpstaj.Next;
   end;
 datam.qtmpstaj.close;

 // ShowMessage(floattostr(xNls)+#13+floattostr(n1)+#13+floattostr(n2)) ;

 nItog:=0;

//  if (n1>=2) and (n2=0) then rtf:=true;          //число авансов >=2 число выплат = 0 тогда

   if n1>0 then rtf:=true;

 //

  datam.qtmpstaj.close;
  datam.qtmpstaj.DatabaseName:=form1.DBDIR;
  datam.qtmpstaj.sql.clear;
  datam.qtmpstaj.sql.add('select * from glnew where nls='+floattostr(xNls));
  datam.qtmpstaj.sql.add('and wm='+floattostr(RMes));
  datam.qtmpstaj.sql.add('and wg='+floattostr(RGod));
  datam.qtmpstaj.Prepare;
  datam.qtmpstaj.Open;
  datam.qtmpstaj.first;
  xNacisl:=0;
  if datam.qtmpstaj.RecordCount>0 then
   begin
    if datam.qtmpstaj.FieldByNAme('DAYRAB').asFloat<>0 then xNAcisl:=mainlib.DRound(datam.qtmpstaj.FieldByName('OKLAD').asFloat*DelenieCas(datam.qtmpstaj.FieldByNAme('DAYOTR').asFloat,
                 datam.qtmpstaj.FieldByNAme('DAYRAB').asFloat,datam.qtmpstaj.FieldByName('DAYCAS').asFloat),form1.DRZn) ;
    xNacisl:=xNacisl+DRound(xNacisl*form1.configRK.Value/100,2)-datam.qtmpstaj.FieldByName('snalog').asFloat;
   end;

  datam.qtmpstaj.close;
  datam.qtmpstaj.DatabaseName:=form1.DBDIR;
  datam.qtmpstaj.sql.clear;
  datam.qtmpstaj.sql.add('select o.*, n.rk,n.koddox from obrt1new o, nacisl n where o.kod=n.kod and o.nls='+floattostr(xNls));
  datam.qtmpstaj.sql.add('and o.wm='+floattostr(RMes));
  datam.qtmpstaj.sql.add('and o.wg='+floattostr(RGod));
  datam.qtmpstaj.Prepare;
  datam.qtmpstaj.Open;
  datam.qtmpstaj.first;
  while not datam.qtmpstaj.eof do
   begin
    { if datam.qtmpstaj.fieldbyname('koddox').asString='2000' then
      begin
    }
       x:=datam.qtmpstaj.fieldbyname('kr').asFloat;
       if datam.qtmpstaj.fieldbyname('rk').asBoolean then x:=x+DRound(x*form1.configrk.Value/100,2) ;
       xNacisl:=xNacisl+x-datam.qtmpstaj.FieldByName('snalog').asFloat;
    {  end;
      }
    datam.qtmpstaj.next;
   end;


  datam.qtmpstaj.close;
  datam.qtmpstaj.DatabaseName:=form1.DBDIR;
  datam.qtmpstaj.sql.clear;
  datam.qtmpstaj.sql.add('select * from sdoxod where nls='+floattostr(xNls));
  datam.qtmpstaj.sql.add('and mes='+floattostr(RMes));
  datam.qtmpstaj.sql.add('and god='+floattostr(RGod));
  datam.qtmpstaj.Prepare;
  datam.qtmpstaj.Open;
  datam.qtmpstaj.first;
  sNacisl:=0;
  while not datam.qtmpstaj.Eof do
   begin
     sNacisl:=sNacisl+datam.qtmpstaj.fieldbyname('sdoxod').asFloat-datam.qtmpstaj.FieldByName('nalog').asFloat; //выплачено всего по sdoxod
    datam.qtmpstaj.next;
   end;

   if rtf then
     begin
      if xNacisl<sNAcisl then nItog:=1 else nItog:=2;
     end;
     
  // ShowMessage(floattostr(xNacisl)+#13+floattostr(sNacisl));



 form123.tOstatok:=xNAcisl-sNAcisl;   //регистр не выплаченный остаток !!!



 //


 Prov2Avans:=nItog;

end;


procedure TForm123.JvXPButton315Click(Sender: TObject);
begin

 

 form2080:=Tform2080.Create(nil);
 form2080.ShowModal;
 form2080.Free;




   if TYPEVYPLAT=2 then
    begin
     ZaprosNOAvans;
     exit;
    end;
    
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);



end;

procedure TForm123.CheckBox199Click(Sender: TObject);
begin
  JvXPButton2Click(nil);
end;

procedure TForm123.JvXPButton16Click(Sender: TObject);
begin
 PopupMenu4.Popup(JvXPButton16.Left+JvXPButton16.Width+form123.Left,JvXPButton16.Top+JvXPButton16.Height+form123.Top);

end;

procedure TForm123.N6Click(Sender: TObject);
var npp,oldRMes,st:integer;
    gm1,gm2:integer;
    gok:boolean;
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


 oldRMes:=RMEs;
 RMes:=gm1;
  form102.RxLabel1.Caption:='Предварительная проверка'   ;
  form102.ProgressBar1.Position:=0;
  form102.Show;
  form102.Refresh;
  npp:=0;
  form1.kart.first;
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
  form102.close;
 RMes:=oldRMes;

 form_58:=tform_58.create(self);
 form_58.RepVypl(gm1,RGod);
 form_58.Free;

end;

procedure TForm123.JvXPButton201Click(Sender: TObject);
begin
  PopupMenu5.Popup(JvXPButton201.Left+JvXPButton201.Width+form123.left,JvXPButton201.Top+JvXPButton201.Height+form123.top)  ;
end;

procedure TForm123.N12Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
begin
   if MessageDlg('Отменить применение вычетов к окончательной выплате заработной платы'+#13+
        'Применено будет к записям, которые еще не проведены',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 oldRMEs:=RMEs;

    query1.first;
    while not query1.eof do
     begin
      if query1.fieldbyname('proveden').asfloat<>1 then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs to oldRMes+1 do
          begin
           RMes:=i;
           form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs+1,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value=1) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=0;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=0;
             datam.pvbase.post;

              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=0;
                vyplzp.post;
               end;

             for i:=oldRMEs to oldRMes+1 do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;

       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);

end;

procedure TForm123.N11Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
begin
     if MessageDlg('Применить вычеты будущего периода к окончательной выплате заработной платы текущего месяца'+#13+
                   'На усмотрение пользователя регулирует как применять вычеты - к выплате аванса будущего месяца либо к выплате з/п за текущий'+#13+
                    'Применено будет к записям, которые еще не проведены',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

    oldRMEs:=RMEs;
    query1.first;
    while not query1.eof do
     begin
      if query1.fieldbyname('proveden').asfloat<>1 then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs to oldRMes+1 do
          begin
           RMes:=i;
           if not datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs+1,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value<>1) and (DRound(datam.pvbaseoivicet.Value,2)>0) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=1;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
             datam.pvbase.post;
              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
                vyplzp.post;
               end;

             for i:=oldRMEs to oldRMes+1 do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;
       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);

end;

procedure TForm123.N9Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
begin
    if RMEs=1 then exit;

   if MessageDlg('Применить вычеты текущего периода к окончательной выплате заработной платы за предыдущий месяц'+#13+
                   'На усмотрение пользователя регулирует как применять вычеты - к выплате аванса текущего месяца либо к выплате з/п за предыдущий'+#13+
                    'Применено будет к записям, которые еще не проведены',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

    oldRMEs:=RMEs;
    query1.first;
    while not query1.eof do
     begin
      if query1.fieldbyname('proveden').asfloat<>1 then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs-1 to oldRMes do
          begin
           RMes:=i;
           if not datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value<>1) and (DRound(datam.pvbaseoivicet.Value,2)>0) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=1;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
             datam.pvbase.post;
              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
                vyplzp.post;
               end;

             for i:=oldRMEs-1 to oldRMes do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;
       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);

end;

procedure TForm123.N13Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
begin

  if RMes=1 then exit;

  if MessageDlg('Отменить применение вычетов к окончательной выплате заработной платы'+#13+
        'Применено будет к записям, которые еще не проведены',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 oldRMEs:=RMEs;

    query1.first;
    while not query1.eof do
     begin
      if query1.fieldbyname('proveden').asfloat<>1 then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs-1 to oldRMes do
          begin
           RMes:=i;
           form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value=1) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=0;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=0;
             datam.pvbase.post;

              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=0;
                vyplzp.post;
               end;

             for i:=oldRMEs-1 to oldRMes do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;

       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);

end;

procedure TForm123.N14Click(Sender: TObject);
begin
  form_58:=tform_58.create(self);
  Form_58.FormRepDatNew(0,0);
  form_58.free;
end;

procedure TForm123.N15Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
    TIdp:Real;
begin

   if query1.RecordCount<=0 then exit;
   Tidp:=query1.fieldbyname('idp').asfloat;

   if query1.fieldbyname('proveden').asfloat=1 then
    begin
     MessageDlg('Текущая запись проведена, операция невозможна',mtinformation,[mbOk],0);
     exit;
    end;

   if MessageDlg('Применить вычеты будущего периода к окончательной выплате заработной платы текущего месяца'+#13+
                   'На усмотрение пользователя регулирует как применять вычеты - к выплате аванса будущего месяца либо к выплате з/п за текущий'+#13+
                    'Применено будет к текущей записи (если не проведена)',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

    oldRMEs:=RMEs;
    query1.first;
    while not query1.eof do
     begin
      if (query1.fieldbyname('proveden').asfloat<>1) and (Tidp=query1.fieldbyname('idp').asfloat) then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs to oldRMes+1 do
          begin
           RMes:=i;
           if not datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs+1,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value<>1) and (DRound(datam.pvbaseoivicet.Value,2)>0) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=1;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
             datam.pvbase.post;
              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
                vyplzp.post;
               end;

             for i:=oldRMEs to oldRMes+1 do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;
       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);
       Query1.Locate('idp',TIdp,[loCaseInsensitive]);
       MessageDlg('Завершено',mtInformation,[mbOk],0);
end;

procedure TForm123.N16Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp:Real;
    TIdp:Real;
begin

 if query1.RecordCount<=0 then exit;
   Tidp:=query1.fieldbyname('idp').asfloat;

   if query1.fieldbyname('proveden').asfloat=1 then
    begin
     MessageDlg('Текущая запись проведена, операция невозможна',mtinformation,[mbOk],0);
     exit;
    end;

  if MessageDlg('Отменить применение вычетов к окончательной выплате заработной платы'+#13+
        'Применено будет к текущей записи если не проведена',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 oldRMEs:=RMEs;

    query1.first;
    while not query1.eof do
     begin
      if (query1.fieldbyname('proveden').asfloat<>1) and (TIdp=query1.fieldbyname('idp').asfloat) then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs to oldRMes+1 do
          begin
           RMes:=i;
           form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs+1,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value=1) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=0;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=0;
             datam.pvbase.post;

              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=0;
                vyplzp.post;
               end;

             for i:=oldRMEs to oldRMes+1 do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;

       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);
       Query1.Locate('idp',TIdp,[loCaseInsensitive]);
        MessageDlg('Завершено',mtInformation,[mbOk],0);
end;

procedure TForm123.N17Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp,TIdp:Real;
begin

  if query1.RecordCount<=0 then exit;
   Tidp:=query1.fieldbyname('idp').asfloat;

    if query1.fieldbyname('proveden').asfloat=1 then
    begin
     MessageDlg('Текущая запись проведена, операция невозможна',mtinformation,[mbOk],0);
     exit;
    end;

  if RMEs=1 then exit;

   if MessageDlg('Применить вычеты текущего периода к окончательной выплате заработной платы за предыдущий месяц'+#13+
                   'На усмотрение пользователя регулирует как применять вычеты - к выплате аванса текущего месяца либо к выплате з/п за предыдущий'+#13+
                    'Применено будет к текущей записи, если еще не проведена',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

    oldRMEs:=RMEs;
    query1.first;
    while not query1.eof do
     begin
      if (query1.fieldbyname('proveden').asfloat<>1) and (Tidp=query1.fieldbyname('idp').asfloat) then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs-1 to oldRMes do
          begin
           RMes:=i;
           if not datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,RMes,RGod]),[locaseinsensitive]) then form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value<>1) and (DRound(datam.pvbaseoivicet.Value,2)>0) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=1;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
             datam.pvbase.post;
              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=DRound(datam.pvbase.fieldbyname('oivicet').asfloat,2);
                vyplzp.post;
               end;

             for i:=oldRMEs-1 to oldRMes do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;
       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);
        Query1.Locate('idp',TIdp,[loCaseInsensitive]);
          MessageDlg('Завершено',mtInformation,[mbOk],0);
end;

procedure TForm123.N18Click(Sender: TObject);
var i,_RMEs,oldRMEs:integer;
    xIdp,TIdp:Real;
begin

 if query1.RecordCount<=0 then exit;
  Tidp:=query1.fieldbyname('idp').asfloat;

   if query1.fieldbyname('proveden').asfloat=1 then
    begin
     MessageDlg('Текущая запись проведена, операция невозможна',mtinformation,[mbOk],0);
     exit;
    end;

  if RMes=1 then exit;

  if MessageDlg('Отменить применение вычетов к окончательной выплате заработной платы'+#13+
        'Применено будет к текущей записи если еще не проведена',mtInformation,[mbYes,mbNo],0) = mrNo then exit;

 oldRMEs:=RMEs;

    query1.first;
    while not query1.eof do
     begin
      if (query1.fieldbyname('proveden').asfloat<>1) and (Tidp=query1.fieldbyname('idp').asfloat) then
        begin
         form1.kart.locate('nls',query1.fieldbyname('nls').asfloat,[locaseinsensitive]);
         for i:=oldRMEs-1 to oldRMes do
          begin
           RMes:=i;
           form1.RaspredVicet;
          end;
         if datam.pvbase.locate('nls;mes;god',VarArrayOf([form1.kartnls.value,oldRMEs,RGod]),[locaseinsensitive]) then
          begin
           if (datam.pvbasePERENOS.Value=1) then
            begin
             datam.pvbase.edit;
             datam.pvbase.fieldbyname('perenos').asfloat:=0;
             datam.pvbase.fieldbyname('fixvicet').asfloat:=0;
             datam.pvbase.post;

              xIdp:=Query1.fieldByName('idp').asFloat;
              if vyplzp.Locate('IDP',xIdp,[loCaseInsensitive]) then
               begin
                vyplzp.edit;
                vyplzp.fieldbyname('avvicet').asfloat:=0;
                vyplzp.post;
               end;

             for i:=oldRMEs-1 to oldRMes do
              begin
               RMes:=i;
               form_58.FNdfl6raspred(true);
               form1.RaspredVicet;
              end;
            end;
          end;
        end;
      query1.next;
     end;

       RMEs:=oldRMEs;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);
         Query1.Locate('idp',TIdp,[loCaseInsensitive]);
           MessageDlg('Завершено',mtInformation,[mbOk],0);
end;

procedure TForm123.N19Click(Sender: TObject);
begin
 form3315:=tform3315.create(nil);
 form3315.ShowModal;
 form3315.free;
end;

procedure TForm123.JvXPButton404Click(Sender: TObject);
var sNAme:String;
    NewValueArray: OLEVariant;
    NSTROK:Integer;
    i,npp:Integer;
    nFields,k:Integer;
    E:Variant;
    fNameXls:String;
    x:real;
    dx1,dx2,dx3:real;
    nDok:integer;
begin
  fNameXls:=GetNameXlsn('Экспорт','exp')+'.xls';

  if not mainlib.FCopyFile(GetCurrentDir()+'\FORM_EXCEL\expvypl.xls',GetCurrentDir()+'\TMP_XLS\'+fNameXls) then
   begin
    MessageDlg('Форма уже открыта, закройте ее',mtError,[mbOk],0);
    exit;
   end;

 E:=CreateOleObject('Excel.Application');
 E.WorkBooks.Open(GetCurrentDir()+'\TMP_XLS\'+fNameXls);

 E.Visible:=True;
 E.Application.WindowState:=2;

 nFields:=11;

  NewValueArray := VarArrayCreate([1, 1, 1, nFields], varVariant);

  query1.First;
  Nstrok:=2;
  dx1:=0; dx2:=0; dx3:=0;
  while not query1.Eof do
   begin
    Nstrok:=Nstrok+1;
    for i:=1 to nFields do NewValueArray[1,i]:='';

      if Query1.fieldByName('PROVEDEN').asInteger=1 then NewValueArray[1,1]:='[+]' else NewValueArray[1,1]:='-' ;

                if (Query1.FieldByName('proveden').asFloat=1) or (DRound(Query1.FieldByName('AVANS').asFloat,2)=0) then s1:=''
                  else
                    begin
                     s1:=FloatToStrF(Query1.FieldByName('AVANS').asFloat,ffNumber,12,2);
                     NewValueArray[1,4]:=Query1.FieldByName('AVANS').asFloat;
                    end;

      NewValueArray[1,3]:=Query1.FieldByName('SUMMA').asFloat;
 //     NewValueArray[1,4]:=Query1.FieldByName('AVANS').asFloat;
      NewValueArray[1,9]:=Query1.FieldByName('AVdoxod').asFloat;
      NewValueArray[1,10]:=Query1.FieldByName('AVnalog').asFloat;
      NewValueArray[1,11]:=Query1.FieldByName('AVvicet').asFloat;
      dx1:=dx1+Query1.FieldByName('summa').asFloat;
      dx2:=dx2+Query1.FieldByName('avdoxod').asFloat;
      dx3:=dx3+Query1.FieldByName('avnalog').asFloat;

      

          x:=Query1.FieldByName('SUMMA').asFloat-Query1.FieldByName('AVANS').asFloat;
          if (Query1.FieldByName('proveden').asFloat=1) or (DRound(Query1.FieldByName('AVANS').asFloat,2)=0) then
               NewValueArray[1,5]:=Query1.FieldByName('SUMMA').asFloat
                       else NewValueArray[1,5]:=x;

   
   
  //    if Query1.FieldByName('DAT').asDateTime>EncodeDate(2000,1,1) then NewValueArra[1,7]:=Query1.FieldByName('DAT').asDateTime;



     NewValueArray[1,2]:=Query1.fieldByName('Fam').asString+' '+Copy(Query1.fieldByName('Im').asString,1,1)+'.'+Copy(Query1.fieldByName('Ot').asString,1,1)+'.';


    if form1.uderj.Locate('KOD',Query1.fieldByName('KOD').asInteger,[loCaseInsensitive]) then
      s1:=AnsiLowerCase(form1.uderjNAME.Value) else s1:='<не определено>';
    NewValueArray[1,6]:=s1;

    if vyplzp.Locate('IDP',Query1.fieldByName('IDP').asInteger,[loCaseInsensitive]) then
      s1:=vyplzpNote.Value else s1:='';
    NewValueArray[1,8]:=s1;



    if form1.uderj.Locate('KOD',Query1.fieldByName('KOD').asInteger,[loCaseInsensitive]) then
     begin
      try
       nDok:=StrToInt(Trim(form1.uderjDBSPRAV.asString));
      except
       nDok:=0;
      end;
      if Query1.FieldByName('datdok').asDateTime>EncodeDate(2000,1,1) then
                            s1:=DateToStr(Query1.FieldByName('datdok').asDateTime)
                             else s1:='-';
      if Query1.FieldByName('numdok').asInteger<>0 then
       begin
        s1:='№'+IntToStr(Query1.FieldByName('numdok').asInteger)+' от '+DateToStr(Query1.FieldByName('datdok').asDateTime);
        if nDok=1 then s1:='р/о №'+IntToStr(Query1.FieldByName('numdok').asInteger)+' '+DateToStr(Query1.FieldByName('datdok').asDateTime);
        if nDok=2 then s1:='п/п №'+IntToStr(Query1.FieldByName('numdok').asInteger)+' '+DateToStr(Query1.FieldByName('datdok').asDateTime);
       end;
     end
       else s1:='<>';
    NewValueArray[1,7]:=s1;

     if not form1.TCheckOO then E.ActiveWorkBook.Sheets.Item[1].Range['A'+InttoStr(NSTROK)+':'+GetAddr(nFields)+InttoStr(NSTROK)]:=newValueArray
      else
       begin
        excellib.PExcelVyvod(E,1,'A'+InttoStr(NSTROK),GetAddr(nFields)+InttoStr(NSTROK),newValueArray) ;
       end;

    query1.Next;
   end;

   if form123.TYPEVYPLAT=2 then //аванс
    begin
     E.ActiveWorkBook.Sheets.Item[1].Columns[4].Hidden:=True;
     E.ActiveWorkBook.Sheets.Item[1].Columns[5].Hidden:=True;
    end
     else
    begin
     E.ActiveWorkBook.Sheets.Item[1].Columns[9].Hidden:=True;
     E.ActiveWorkBook.Sheets.Item[1].Columns[10].Hidden:=True;
     E.ActiveWorkBook.Sheets.Item[1].Columns[11].Hidden:=True;
    end ;


       try
         for i:=7 to 12 do  E.ActiveWorkbook.Sheets.Item[1].Range['A'+IntToStr(3),GetAddr(nFields)+IntToStr(Nstrok)].Borders[i].LineStyle:=1;
       except
       end;

   try
  //  for i:=1 to 30 do E.ActiveWorkbook.Sheets.Item[1].Columns[i].AutoFit;
   except
   end;


   NStrok:=NStrok+1;
   for i:=1 to nFields do NewValueArray[1,i]:='';
    NewValueArray[1,3]:=dx1;
    NewValueArray[1,9]:=dx2;
    NewValueArray[1,10]:=dx3;


     if not form1.TCheckOO then E.ActiveWorkBook.Sheets.Item[1].Range['A'+InttoStr(NSTROK)+':'+GetAddr(nFields)+InttoStr(NSTROK)]:=newValueArray
      else
       begin
        excellib.PExcelVyvod(E,1,'A'+InttoStr(NSTROK),GetAddr(nFields)+InttoStr(NSTROK),newValueArray) ;
       end;


   E.Visible:=True;
   E.Application.WindowState:=-4137;
   E:=UnAssigned;

end;

procedure TForm123.N7Click(Sender: TObject);
var i,j:integer;
begin
 form2073:=tform2073.Create(nil);

 form2073.TVYB:=TVYBOR; //галочки все или ничего

 for i:=0 to 100 do form2073.DKOD[i]:=false;

 ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]

 for i:=0 to 100 do form2073.DKOD[i]:=UWKOD[i];


 form2073.JvCheckListBox1.Items.Clear;
 j:=-1;
 for i:=0 to 100 do
  begin
   if form2073.DKOD[i] then
     begin
      j:=j+1;
      if i=0 then form2073.JvCheckListBox1.Items.Add('Оклад/тариф');
      if i>0 then
        begin
         form1.NACISL.Locate('kod',i,[loCaseInsensitive]);
         form2073.JvCheckListBox1.Items.Add(form1.NACISLNAME.Value);
        end;
       form2073.DKOD2[j]:=i;
       if form2073.TVYB then form2073.JvCheckListBox1.Checked[j]:=true
                               else form2073.JvCheckListBox1.Checked[j]:=false;
     end;
  end;



 form2073.ShowModal;


 if form2073.TOk then
   begin


    for i:=0 to 100 do UWKOD[i]:=false;
    for i:=0 to form2073.JvCheckListBox1.Items.Count-1 do
    begin
     if form2073.JvCheckListBox1.Checked[i] then
       begin
        UWKOD[form2073.DKOD2[i]]:=true;
      end;
    end;

    ZaprosV1(1); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
    xTYPE:=0; {для ZaprosV2}
    ZaprosV2(0);
    JvXPButton12Click(nil);
   end
    else
     begin
       T_OWKOD:=false;
       ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
       xTYPE:=0; {для ZaprosV2}
       ZaprosV2(0);
       JvXPButton12Click(nil);

     end ;


 form2073.Free;

  //RadioGroup2.ItemIndex:=0;

 if (T_OWKOD) and (TYPEVYPLAT<>2) then
       begin
         DBGrid1.Columns[3].Visible:=false;
         DBGrid1.Columns[4].Visible:=false;
         //RadioGroup2.ItemIndex:=2;
       end ;

 if (not T_OWKOD) and (TYPEVYPLAT<>2) then
       begin
         DBGrid1.Columns[3].Visible:=true;
         DBGrid1.Columns[4].Visible:=true;
       end ;





end;

procedure TForm123.JvXPButton288Click(Sender: TObject);
begin
  PopupMenu6.Popup(JvXPButton288.Left+JvXPButton288.Width+form123.left,JvXPButton288.Top+JvXPButton288.Height+form123.top)  ;
end;

procedure TForm123.N22Click(Sender: TObject);
begin

 if TYPEVYPLAT=0 then
  begin
   _TPKOD:=-1;
   ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
   xTYPE:=0; {для ZaprosV2}
   ZaprosV2(0);
   JvXPButton12Click(nil);
   MessageDlg('Фильтр по подразделению снят',mtInformation,[mbOk],0);
  end ;

  if TYPEVYPLAT=2 then
  begin
   _TPKOD:=-1;
   ZaprosV3;
   xTYPE:=2; {аванс, для ZaprosV2}
   ZaprosV2(0);
   JvXPButton12Click(nil);
   MessageDlg('Фильтр по подразделению снят',mtInformation,[mbOk],0);
  end;
  
end;

procedure TForm123.N20Click(Sender: TObject);
var sName:String;
begin
 sName:='';
 form_4:=TForm_4.Create(nil);
 form_4.JvXPButton1.Visible:=False;
 form_4.JvXPButton2.Visible:=True;
 form_4.Vfilial:='';
 form_4.ShowModal;
 form_4.JvXPButton1.Visible:=True;
 form_4.JvXPButton2.Visible:=False;
 if form_4.Tok then
   begin
    _TPKOD:=form1.filialPKOD.Value;
    sName:=form1.filialname.value;
   end;
 form_4.Free;

  if TYPEVYPLAT=0 then
  begin
   ZaprosV1(0); //0 - заполнение UWKOD[i], 1 - исполнение UWKOD[i]
   xTYPE:=0; {для ZaprosV2}
   ZaprosV2(0);
   JvXPButton12Click(nil);
   if _TPKOD>=0 then MessageDlg('Установлен фильтр по подразделению'+#13+sName,mtInformation,[mbOk],0);
  end;

 if TYPEVYPLAT=2 then
  begin
   ZaprosV3;
   xTYPE:=2; {аванс, для ZaprosV2}
   ZaprosV2(0);
   JvXPButton12Click(nil);
   if _TPKOD>=0 then MessageDlg('Установлен фильтр по подразделению'+#13+sName,mtInformation,[mbOk],0);
  end;



end;

procedure TForm123.JvXPButton20Click(Sender: TObject);
begin
 ControlMrot;
   JvXPButton6Click(nil);
end;

end.
