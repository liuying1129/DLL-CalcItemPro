library CalcItemPro;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  SysUtils,
  Classes,
  ADODB;

{$R *.res}

//�ҵ����ʽ��С����λ�������ֵ.��56.5*100+23.01��ֵΪ2
function MaxDotLen(const ACalaExp:PChar):integer;stdcall;external 'LYFunction.dll';
function CalParserValue(const CalExpress:Pchar;var ReturnValue:single):boolean;stdcall;external 'CalParser.dll';

function InputToArry(const CaculExpress: string):TStrings;//�����㹫ʽ[2]+[1]/[3]���뵽�ַ����б���,��('2','1','3')
var
  sCaculExpress:string;
begin
  result:=TStringList.Create;
  sCaculExpress:=CaculExpress;
  while (pos('[',sCaculExpress)<>0)and(pos(']',sCaculExpress)<>0) do
  begin
    result.Add(copy(sCaculExpress,pos('[',sCaculExpress)+1,pos(']',sCaculExpress)-pos('[',sCaculExpress)-1));
    delete(sCaculExpress,1,pos(']',sCaculExpress));
  end;
end;

procedure addOrEditCalcItem(const Aadoconnstr:Pchar;const ComboItemID:Pchar;const checkunid: integer);stdcall;
//��������Ŀ���ӻ�༭������������
  function CompareArray(Refer,b: TStrings):boolean;
  //�ַ����б�b�е�Ԫ��('2','1','3')���ַ����б�Refer���Ƿ񶼴���
  var
    i,j:integer;
    bs:boolean;
  begin
    bs:=true;
    IF (B.Count=0)or(Refer.Count=0) THEN BS:=FALSE;
    for i:=0 to b.Count-1 do
    begin
      if not bs then begin result:=false;exit;end;
      for  j:=0  to Refer.Count-1 do
      begin
        if b[i]=Refer[j] then begin bs:=true;break;end else
        bs:=false;
      end;
    end;
    if bs then result:=true else result:=false;
  end;
var
  adotemp2,adotemp3,adotemp4:tadoquery;
  adoconn:Tadoconnection;
  ISSURE,strsql:string;
  aFF,sAA:TStrings;//sAA������ż������е�'��Ŀ����'�ַ����б�
begin
  adoconn:=Tadoconnection.Create(nil);
  adoconn.ConnectionString:=strpas(Aadoconnstr);
  adoconn.LoginPrompt:=false;
  
  //�����˼�����Ŀ�����sAA�� start
  adotemp3:=tadoquery.Create(nil);
  adotemp3.Connection:=adoconn;
  adotemp3.Close;
  adotemp3.SQL.Clear;
  adotemp3.SQL.Text:='SELECT issure,'+
                     ' itemid '+
                     ' FROM chk_valu '+
                     ' WHERE '+
                     ' (pkunid = :P_pkunid) '+
                     ' and pkcombin_id=:P_pkcombin_id ';
  adotemp3.Parameters.ParamByName('P_pkunid').Value:=checkunid;
  adotemp3.Parameters.ParamByName('P_pkcombin_id').Value:=trim(strpas(ComboItemID));
  adotemp3.Open;//���˼�������.ע�������Ŀ����Ҫ������һ��������Ŀ�ļ���(���/���ֵ���򵰰�)
  issure:=trim(adotemp3.fieldbyname('issure').AsString);
  if(issure<>'0')and(issure<>'1')then issure:='1'; //�����Ӧ�ò����ܳ���
  sAA:=TStringList.Create;
  while not adotemp3.Eof do
  begin
    sAA.Add(adotemp3.fieldbyname('itemid').AsString);
    adotemp3.Next;
  end;
  adotemp3.Free;
  //�����˼�����Ŀ�����sAA�� end

  adotemp2:=tadoquery.Create(nil);
  adotemp2.Connection:=adoconn;
  adotemp2.Close;
  adotemp2.SQL.Clear;
  adotemp2.SQL.Text:=' select * from clinicchkitem where ltrim(rtrim(isnull(caculexpress,'''')))<>'''' order by itemid ';
  adotemp2.Open;   //���еļ�����Ŀ
  while not adotemp2.Eof do
  begin
    aFF:=InputToArry(adotemp2.fieldbyname('caculexpress').AsString);

    //�����˼�����Ŀ�����sAA�С���ǰ��������ģ�Ϊ�����ѡ�����Ŀʱ�������⣬�ᵽ��ǰ��
    
    if CompareArray(sAA,aFF) then
    begin
      adotemp4:=tadoquery.Create(nil);
      adotemp4.Connection:=adoconn;
      adotemp4.Close;
      adotemp4.SQL.Clear;
      strsql :=' select issure from chk_valu where pkunid=:P_pkunid and pkcombin_id=:P_pkcombin_id and itemid=:P_itemid ';
      adotemp4.SQL.Text:=strsql;
      ADOtemp4.Parameters.ParamByName('P_pkcombin_id').Value :=trim(strpas(ComboItemID));
      ADOtemp4.Parameters.ParamByName('P_pkunid').Value := checkunid;
      ADOtemp4.Parameters.ParamByName('p_itemid').Value := adotemp2.FieldByName('itemid').AsString;
      ADOtemp4.Open;
      if adotemp4.RecordCount>0 then
      begin
        while not adotemp4.Eof do
        begin
          adotemp4.Edit;
          adotemp4.FieldByName('issure').AsString:=issure;
          adotemp4.Post;
          adotemp4.Next;
        end;
        adotemp4.Free;
        adotemp2.Next;
        continue;
      end;
      adotemp4.Free;

      adotemp4:=tadoquery.Create(nil);
      adotemp4.Connection:=adoconn;
      adotemp4.Close;
      adotemp4.SQL.Clear;
      strsql := 'insert into chk_valu(pkunid,issure,pkcombin_id,itemid ' +
              ' ) select' +
              ' :pkunid,:p_ISSURE,:P_combinitem,itemid ' +
              ' from clinicchkitem where itemid=:P_itemid ';

      adotemp4.SQL.Text:=strsql;
      ADOtemp4.Parameters.ParamByName('P_combinitem').Value := trim(strpas(ComboItemID));
      ADOtemp4.Parameters.ParamByName('pkunid').Value := checkunid;
      ADOtemp4.Parameters.ParamByName('p_itemid').Value := adotemp2.FieldByName('itemid').AsString;
      ADOtemp4.Parameters.ParamByName('p_ISSURE').Value :=issure;
      ADOtemp4.ExecSQL;
      adotemp4.Free;

      sAA.Add(adotemp2.FieldByName('itemid').AsString);
    end;
    aFF.Free;
    adotemp2.Next;
  end;
  adotemp2.Free;

  sAA.Free;
  
  adoconn.Free;
end;

procedure addOrEditCalcValu(const Aadoconnstr:Pchar;const checkunid: integer;const AifInterface:boolean;const ATransItemidString:pchar);stdcall;
//�������������ӻ�༭������������
//�����������ĿATransItemidString��ʽ:[1101][1102]
const
  ssql1=' select clinicchkitem.caculexpress,chk_valu.pkcombin_id as combinitem,chk_valu.itemvalue,chk_valu.itemid '+
        ' from clinicchkitem,chk_valu '+
        ' where ltrim(rtrim(isnull(clinicchkitem.caculexpress,'''')))<>'''' and '+
        ' clinicchkitem.itemid=chk_valu.itemid and chk_valu.pkunid=:P_pkunid ';
//  ssql2=' and ltrim(rtrim(isnull(clinicchkitem.caculexpress,''''))) not like ''%t[''+chk_valu.itemid+''t]%'' escape ''t'' ';//��ʾ�����������������Ŀ.�ӿڳ����޴���������
  ssql3=' order by clinicchkitem.itemid ';
var
  adoquerytemp:tadoquery;  //��ǰ�ļ�������
  ADOQuerytemp4:tadoquery;
  calc_express:string;
  i,iMaxDotLen:integer;
  querystr:string;
  itemid:string;
  itemvalueJ:string;
  l_ReturnValue:single;
  ls:TStrings;
  adoconn:Tadoconnection;
begin
  adoconn:=Tadoconnection.Create(nil);
  adoconn.ConnectionString:=strpas(Aadoconnstr);
  adoconn.LoginPrompt:=false;
  
  adoquerytemp:=tadoquery.Create(nil); //ָ�����˵ļ�����Ŀ
  adoquerytemp.Connection:=adoconn;
  adoquerytemp.Close;
  adoquerytemp.SQL.Clear;
  adoquerytemp.SQL.Text:=ssql1+ssql3;
  adoquerytemp.Parameters.ParamByName('p_pkunid').Value:=checkunid;
  adoquerytemp.Open;
  if adoquerytemp.RecordCount=0 then begin adoquerytemp.Free; exit; end;

  adoquerytemp.First;
  while not adoquerytemp.Eof do
  begin
    calc_express:=trim(adoquerytemp.fieldbyname('caculexpress').AsString);
    itemid:=trim(adoquerytemp.fieldbyname('itemid').AsString);
    if (pos('['+itemid+']',calc_express)>0)and(not AifInterface) then begin adoquerytemp.Next;continue; end;//(���������Ŀ)and(�ǽӿ�)
    if (pos('['+itemid+']',calc_express)>0)and(AifInterface)and(pos('['+itemid+']',strpas(ATransItemidString))=0) then begin adoquerytemp.Next;continue; end;//(���������Ŀ)and(�ӿ�)and(����Ŀû������)
    ls:=InputToArry(calc_express);
    for  i:=0  to ls.Count-1 do
    begin
      //============�ҵ�ָ��Ψһ��ŵ���ĿֵitemvalueJ==================//
      querystr:='select itemvalue '+
                ' from chk_valu '+
                ' where pkunid=:P_pkunid '+
                ' and pkcombin_id=:p_pkcombin_id '+
                ' and itemid=:itemid ';
      ADOQuerytemp4:=tadoquery.Create(nil);
      ADOQuerytemp4.Connection:=adoconn;
      ADOQuerytemp4.Close;
      ADOQuerytemp4.SQL.Clear;
      ADOQuerytemp4.SQL.Text:=querystr;
      ADOQuerytemp4.Parameters.ParamByName('P_pkunid').Value:=checkunid;
      ADOQuerytemp4.Parameters.ParamByName('p_pkcombin_id').Value:=
                        adoquerytemp.fieldbyname('combinitem').Value;
      ADOQuerytemp4.Parameters.ParamByName('itemid').Value:=ls[i];
      ADOQuerytemp4.Open;

      itemvalueJ:=ADOQuerytemp4.fieldbyname('itemvalue').AsString;
      ADOQuerytemp4.Close;
      ADOQuerytemp4.Free;
      //==================================================================//
      calc_express:=
        StringReplace(calc_express,'['+ls[i]+']',itemvalueJ,[rfReplaceAll,rfIgnoreCase]);
    end;

    ls.Free;

    iMaxDotLen:=MaxDotLen(pchar(calc_express));
    if CalParserValue(Pchar(calc_express),l_ReturnValue) then
      calc_express:=format('%.'+inttostr(iMaxDotLen)+'f',[l_ReturnValue])
    else begin
      if pos('['+itemid+']',calc_express)>0 then//�ӿڴ��������Լ�����Ŀ,���޷�����,����ʾԭֵ.��Ѫ�򴫹�����***
        calc_express:=itemvalueJ
      else calc_express:='';
    end;

    adoquerytemp.Edit;
    adoquerytemp.FieldByName('itemvalue').AsString:=calc_express;
    adoquerytemp.Post;

    adoquerytemp.Next;
  end; //end while
  adoquerytemp.Free;
  adoconn.Free;
end;

exports
  addOrEditCalcItem,
  addOrEditCalcValu;
  
begin
end.
