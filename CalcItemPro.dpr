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

//找到表达式中小数点位数的最大值.如56.5*100+23.01的值为2
function MaxDotLen(const ACalaExp:PChar):integer;stdcall;external 'LYFunction.dll';
function CalParserValue(const CalExpress:Pchar;var ReturnValue:single):boolean;stdcall;external 'CalParser.dll';

function InputToArry(const CaculExpress: string):TStrings;//将计算公式[2]+[1]/[3]导入到字符串列表中,如('2','1','3')
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
//将计算项目增加或编辑到检验结果表中
  function CompareArray(Refer,b: TStrings):boolean;
  //字符串列表b中的元素('2','1','3')在字符串列表Refer中是否都存在
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
  aFF,sAA:TStrings;//sAA用来存放检验结果中的'项目代码'字符串列表
begin
  adoconn:=Tadoconnection.Create(nil);
  adoconn.ConnectionString:=strpas(Aadoconnstr);
  adoconn.LoginPrompt:=false;
  
  //将病人检验项目表放入sAA中 start
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
  adotemp3.Open;//病人检验结果表.注意计算项目可能要参与另一个计算项目的计算(如白/球比值、球蛋白)
  issure:=trim(adotemp3.fieldbyname('issure').AsString);
  if(issure<>'0')and(issure<>'1')then issure:='1'; //该情况应该不可能出现
  sAA:=TStringList.Create;
  while not adotemp3.Eof do
  begin
    sAA.Add(adotemp3.fieldbyname('itemid').AsString);
    adotemp3.Next;
  end;
  adotemp3.Free;
  //将病人检验项目表放入sAA中 end

  adotemp2:=tadoquery.Create(nil);
  adotemp2.Connection:=adoconn;
  adotemp2.Close;
  adotemp2.SQL.Clear;
  adotemp2.SQL.Text:=' select * from clinicchkitem where ltrim(rtrim(isnull(caculexpress,'''')))<>'''' order by itemid ';
  adotemp2.Open;   //所有的计算项目
  while not adotemp2.Eof do
  begin
    aFF:=InputToArry(adotemp2.fieldbyname('caculexpress').AsString);

    //将病人检验项目表放入sAA中。以前放在这里的，为解决勾选组合项目时慢的问题，提到了前面
    
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
//将计算数据增加或编辑到检验结果表中
//传输过来的项目ATransItemidString格式:[1101][1102]
const
  ssql1=' select clinicchkitem.caculexpress,chk_valu.pkcombin_id as combinitem,chk_valu.itemvalue,chk_valu.itemid '+
        ' from clinicchkitem,chk_valu '+
        ' where ltrim(rtrim(isnull(clinicchkitem.caculexpress,'''')))<>'''' and '+
        ' clinicchkitem.itemid=chk_valu.itemid and chk_valu.pkunid=:P_pkunid ';
//  ssql2=' and ltrim(rtrim(isnull(clinicchkitem.caculexpress,''''))) not like ''%t[''+chk_valu.itemid+''t]%'' escape ''t'' ';//表示不计算对自身计算的项目.接口程序无此条件限制
  ssql3=' order by clinicchkitem.itemid ';
var
  adoquerytemp:tadoquery;  //当前的检验结果集
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
  
  adoquerytemp:=tadoquery.Create(nil); //指定病人的计算项目
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
    if (pos('['+itemid+']',calc_express)>0)and(not AifInterface) then begin adoquerytemp.Next;continue; end;//(自身计算项目)and(非接口)
    if (pos('['+itemid+']',calc_express)>0)and(AifInterface)and(pos('['+itemid+']',strpas(ATransItemidString))=0) then begin adoquerytemp.Next;continue; end;//(自身计算项目)and(接口)and(该项目没传过来)
    ls:=InputToArry(calc_express);
    for  i:=0  to ls.Count-1 do
    begin
      //============找到指定唯一编号的项目值itemvalueJ==================//
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
      if pos('['+itemid+']',calc_express)>0 then//接口传过来的自计算项目,如无法计算,则显示原值.如血球传过来的***
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
