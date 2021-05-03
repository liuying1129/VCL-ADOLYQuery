{**********************************************************}
{                                                          }
{  通用查询组件:TADOLYQuery Component Version 06.09.27     }
{  作者：刘鹰                                              }
{                                                          }
{                                                          }
{  新功能：                                                }
{                                                          }
{                                                          }
{  功能:                                                   }
{                                                          }
{                                                          }
{  调用方式：                                              }
{  lyquery1.Connection:=ADOConnection1;                    }
{  lyquery1.SelectString:=                                 }
{          'select patientname,age_real,unid,check_date from chk_con';
{  if lyquery1.Execute then                                }
{  begin                                                   }
{    showmessage(lyquery1.ResultSelect);                   }
{  end;                                                    }
{                                                          }
{                                                          }
{  他是一个免费软件,如果你修改了他,希望我能有幸看到你的杰作}
{                                                          }
{  我的 Email: Liuying1129@163.com                         }
{                                                          }
{  版权所有,欲用于商业用途,请与我联系!!!                   }
{                                                          }
{  Bug:                                                    }
{  1.若有group by子句，要求group与by之间只能有一个空格     }
{  2.若有order by子句，要求order与by之间只能有一个空格     }
{  3.不支持where子句中的select from子查询                  }
{  4.仅支持字段的as别名方式                                }
{**********************************************************}

unit UADOLYQuery;

interface

uses
  Classes, Forms,Inifiles,SysUtils{StringReplace}, 
  Buttons, ADODB,Controls, ExtCtrls,DB,StrUtils,
  ComCtrls{TDateTimePicker}, StdCtrls,Windows{SetWindowLong};

type TArFieldType = array of TFieldType;
     TDataBaseType = (dbtMSSQL,dbtOracle,dbtAccess);

type
  TfrmADOLYQuery = class(TForm)
    Panel1: TPanel;
    BitBtnCommQry: TBitBtn;
    BitBtnCommQryClose: TBitBtn;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    ScrollBox1: TScrollBox;
    btn_dec: TSpeedButton;
    btn_add: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure BitBtnCommQryClick(Sender: TObject);
    procedure btn_addClick(Sender: TObject);
    procedure btn_decClick(Sender: TObject);
    procedure BitBtnCommQryCloseClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
  private
    { Private declarations }
    slFieldNames:TStrings;//字段显示标签列表
    slFieldsList:TStrings;//字段列表
    ArFieldType:TArFieldType;//字段数据类型
    function GetFieldsList(const SelectString:string;Const FieldNameList:TStrings):TStrings;//取得字段列表
    function GetFieldNames(const SelectString:string):TStrings;//取得字段显示标签列表
    function GetFieldType(const SelectString:string):TArFieldType;
    function GetSqlTxt:string;
    procedure CreateAQueryControl(ifNumOne:boolean);
    procedure CreateAValueControl(AParentControl:TWinControl;AFieldType:TFieldType;AFieldName:string);
    procedure cmbFieldNameOnChange(sender:TObject);
  public
    { Public declarations }
    pConnection:TADOConnection;
    pSelectString:STRING;
    pResult:boolean;
    pResultSelect:string;
    pDataBaseType:TDataBaseType;
  end;

type
  TADOLYQuery = class(TComponent)
  private
    { Private declarations }
    FConnection:TADOConnection;
    FSelectString:STRING;
    FResultSelect:string;
    ffrmLYQuery:TfrmADOLYQuery;
    FDataBaseType:TDataBaseType;
    procedure FSetConnection(value:TADOConnection);
    procedure FSetSelectString(value:string);
    procedure FSetDataBaseType(value:TDataBaseType);
  protected
    { Protected declarations }
  public
    { Public declarations }
    constructor create(aowner:tcomponent);override;
    destructor destroy;override;
    function Execute:boolean;
    property ResultSelect:string read FResultSelect;
  published
    { Published declarations }
    property Connection:TADOConnection read FConnection write FSetConnection;
    property SelectString:string read FSelectString write FSetSelectString;
    property DataBaseType:TDataBaseType read FDataBaseType write FSetDataBaseType;
  end;

procedure Register;

implementation

{$R *.dfm}

procedure Register;
begin
  RegisterComponents('Eagle_Ly', [TADOLYQuery]);
end;

function LastPos(const subStr,sourStr:string):integer;
//取得subStr在sourStr中最后一次出现的位置
var
  sub,sour:string;
begin
  if Pos(subStr,sourStr)=0 then
  begin
    Result:=0;
    exit;
  end;
  sub:=ReverseString(subStr);
  sour:=ReverseString(sourStr);
  Result:=length(sourStr)-Pos(sub,sour)+1-length(subStr)+1;
end; 

function GetSelectPart(const ASelectString:string;var AOrderBy:string;var AGroupBy:string):string;
//取得select部分的语句.如select * from table where f1>0 group by f2 order by f3
//返回select * from table where f1>0 and
//该函数的作用是:使select语句可继续增加where条件
var
  tmpSelString:string;
  GroupByPos,OrderByPos:integer;
  frompos,wherepos:INTEGER;
begin
  tmpSelString:=ASelectString;
  
  frompos:=LastPos(' from ',pchar(LowerCase(ASelectString)));//决定了不支持where子句中的select from子查询

  OrderByPos:=Lastpos(' order by ',pchar(LowerCase(ASelectString)));//这就要求order与by之间只能有一个空格
  if (OrderByPos<>0) and (OrderByPos>frompos)then  //取得order by子句
  begin
    AOrderBy:=ASelectString;
    delete(AOrderBy,1,OrderByPos-1);
    tmpSelString:=copy(ASelectString,1,OrderByPos-1);
  end;

  GroupByPos:=Lastpos(' group by ',pchar(LowerCase(ASelectString)));//这就要求group与by之间只能有一个空格
  if (GroupByPos<>0) and (GroupByPos>frompos) then  //取得group by子句
  begin
    AGroupBy:=tmpSelString;
    delete(AGroupBy,1,GroupByPos-1);
    tmpSelString:=copy(ASelectString,1,GroupByPos-1);
  end;

  wherepos:=LastPos(' where ',pchar(LowerCase(ASelectString)));
  if (wherepos<>0) and (wherepos>frompos) then //取得where子句
    result:=tmpSelString+' and '
  else result:=tmpSelString+' where ';
end;

function TfrmADOLYQuery.GetFieldNames(const SelectString:string):TStrings;//取得字段显示标签列表
var
  tmpSelString:string;
  adotemp11:tadoquery;
  a,b:string;//无实际用处
begin
  Result := TStringList.Create;

  tmpSelString:=GetSelectPart(SelectString,a,b);
  tmpSelString:=tmpSelString+' 1=0 ';
  
  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=pConnection;
  adotemp11.Close;
  adotemp11.SQL.Clear;
  adotemp11.SQL.Text:=tmpSelString;
  try
     adotemp11.Open;//只有打开的情况下才能检测到字段
  except
     raise Exception.Create('请检查您的SQL语句!');
     adotemp11.Free;
     exit;
  end;
  adotemp11.Fields.GetFieldNames(Result);
  adotemp11.Free;
end;

function TfrmADOLYQuery.GetFieldType(
  const SelectString: string): TArFieldType;
var
  tmpSelString:string;
  adotemp11:tadoquery;
  i:integer;
  a,b:string;//无实际用处
begin
  tmpSelString:=GetSelectPart(SelectString,a,b);
  tmpSelString:=tmpSelString+' 1=0 ';

  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=pConnection;
  adotemp11.Close;
  adotemp11.SQL.Clear;
  adotemp11.SQL.Text:=tmpSelString;
  try
     adotemp11.Open;//只有打开的情况下才能检测到字段
  except
     raise Exception.Create('请检查您的SQL语句!');
     adotemp11.Free;
     Result:=nil;
     exit;
  end;
  setlength(Result,adotemp11.FieldCount);
  for i :=0  to adotemp11.FieldCount-1 do
  begin
    Result[i]:=adotemp11.Fields[i].datatype;
  end;
  adotemp11.Free;
end;

function TfrmADOLYQuery.GetFieldsList(const SelectString:string;const FieldNameList:TStrings):TStrings;//取得字段列表
var
  sqlstr1,tmpSelString:string;
  j,k,iLen:integer;
  adotemp11:tadoquery;
  a,b:string;//无实际用处

  //给字段加上表名的变量
  sFieldName,sFieldProperty:String;
  JhPos:INTEGER;
  //====================

  i:integer;
  sList:TStrings;
begin
  Result := TStringList.Create;

  sqlstr1:=stringreplace(SelectString,#10,'',[rfReplaceAll,rfIgnoreCase]);
  sqlstr1:=stringreplace(sqlstr1,#13,'',[rfReplaceAll,rfIgnoreCase]);
  while Pos('  ',sqlstr1)>0 do
    sqlstr1:=stringreplace(sqlstr1,'  ',' ',[rfReplaceAll,rfIgnoreCase]);

  for j :=1  to FieldNameList.Count do
  begin
      k:=pos(lowercase(' as '+FieldNameList[j-1]+' '),lowercase(sqlstr1));//' as 中国 from '的情况
      if k>0 then
      begin
        iLen:=length(' as '+FieldNameList[j-1]);
        delete(sqlstr1,k,iLen);
        Continue;
      end;
      k:=pos(lowercase(' as '+FieldNameList[j-1]+','),lowercase(sqlstr1));//' as 中国,... from '的情况
      if k>0 then
      begin
        iLen:=length(' as '+FieldNameList[j-1]);
        delete(sqlstr1,k,iLen);
        Continue;
      end;
  end;

  tmpSelString:=GetSelectPart(sqlstr1,a,b);
  tmpSelString:=tmpSelString+' 1=0 ';

  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=pConnection;
  adotemp11.Close;
  adotemp11.SQL.Clear;
  adotemp11.SQL.Text:=tmpSelString;
  try
    adotemp11.Open;//只有打开的情况下才能检测到字段
  except
    raise Exception.Create('生成字段列表时出错!');
    adotemp11.Free;
    exit;
  end;
  adotemp11.Fields.GetFieldNames(result);
  adotemp11.Free;

  //字段属性 add by liuying 20100811
  //1z2y3x:该字符后的串为字段属性串
  //1w2v3u:各属性的分隔
  //1T2S3R:属性名与属性值的分隔.属性COMBOBOXITEMS各ITEM间的分隔
  for j :=0  to result.Count-1 do
  begin
    sFieldName:=slFieldNames[j];
    JhPos:=Pos('1z2y3x',LowerCase(sFieldName));
    if JhPos<=0 then continue;
    
    sFieldProperty:=trim(copy(sFieldName,JhPos+6,maxint));
    sFieldProperty:=StringReplace(sFieldProperty,'1w2v3u',#$2,[rfReplaceAll,rfIgnoreCase]);
    sList:=TStringList.Create;
    ExtractStrings([#$2],[],PChar(sFieldProperty),sList);
    for i :=0 to sList.Count-1 do
    begin
      if UPPERCASE(leftstr(sList[i],15))='FIELDNAME1T2S3R' THEN//字段名.如果字段重复,GetFieldNames会在重复的字段名后加'_1'等
      BEGIN
        result[j]:=copy(sList[i],16,maxint);
      END;
    end;
    for i :=0 to sList.Count-1 do
    begin
      //if pos('1T2S3R',uppercase(sList[i]))<=0 THEN//兼容老版本,即1z2y3x后直接跟表名的版本
      //BEGIN
      //  result[j]:=sList[i]+'.'+result[j];
      //END;
      if UPPERCASE(leftstr(sList[i],15))='TABLENAME1T2S3R' THEN//表名
      BEGIN
        result[j]:=copy(sList[i],16,maxint)+'.'+result[j];
      END;
    end;
    sList.Free;
  end;
  //============================
  
  //处理Access的datetime类型字段//add by LY 20090821
  if pDataBaseType=dbtAccess then
  begin
    for j :=0  to result.Count-1 do
    begin
      if(ArFieldType[j]<>ftDate)and(ArFieldType[j]<>ftTime)and(ArFieldType[j]<>ftDateTime)then continue;

      if pos('FIELDTYPE1T2S3RFTTIME',uppercase(slFieldNames[j]))>0 then result[j]:='format('+result[j]+',''hh:mm:ss'')'
        else result[j]:='format('+result[j]+',''YYYY-MM-DD'')';
    end;
  end;
  //===========================

  //处理oracle的datetime类型字段
  if pDataBaseType=dbtOracle then
  begin
    for j :=0  to result.Count-1 do
    begin
      if(ArFieldType[j]=ftDate)or(ArFieldType[j]=ftTime)or(ArFieldType[j]=ftDateTime)then
      begin
        if pos('FIELDTYPE1T2S3RFTTIME',uppercase(slFieldNames[j]))>0 then result[j]:='to_char('+result[j]+',''HH24:mi:ss'')'
          else result[j]:='to_char('+result[j]+',''YYYY-MM-DD'')';
      end;
      if(ArFieldType[j]=ftString)or(ArFieldType[j]=ftWideString)then//add by liuying 20100825
      begin
        result[j]:='nvl('+result[j]+',''!@#$%'')';
      end;
    end;
  end;
  //===========================

  //处理SQL Server的datetime类型字段
  if pDataBaseType=dbtMSSQL then
  begin
    for j :=0  to result.Count-1 do
    begin
      if(ArFieldType[j]=ftDate)or(ArFieldType[j]=ftTime)or(ArFieldType[j]=ftDateTime)then
      begin
        if pos('FIELDTYPE1T2S3RFTTIME',uppercase(slFieldNames[j]))>0 then result[j]:='CONVERT(CHAR(8),'+result[j]+',108)'
          else result[j]:='CONVERT(CHAR(10),'+result[j]+',121)';
      end;
      if(ArFieldType[j]=ftString)or(ArFieldType[j]=ftWideString)then//add by liuying 20100825
      begin
        result[j]:='isnull('+result[j]+','''')';
      end;
    end;
  end;
  //===========================
end;

procedure TfrmADOLYQuery.cmbFieldNameOnChange(sender: TObject);
begin
  CreateAValueControl(TCombobox(sender).Parent,ArFieldType[TCombobox(sender).ItemIndex],slFieldNames[TCombobox(sender).ItemIndex]);
end;

procedure TfrmADOLYQuery.CreateAQueryControl(ifNumOne: boolean);
var
  Panel:TPanel;
  cmbFieldName,cmbAnd,cmbAmount:TCombobox;
  i,j:integer;
begin
  Panel:=TPanel.Create(self);
  Panel.Parent:=ScrollBox1;
  Panel.Top:=ScrollBox1.ControlCount*Panel.Width+10;
  Panel.Align:=alTop;

  if not ifNumOne then
  begin
    cmbAnd:=TCombobox.Create(self);
    cmbAnd.Parent:=Panel;
    cmbAnd.Width:=59;cmbAnd.Left:=8;cmbAnd.Top:=8;
    cmbAnd.Tag:=1;
    cmbAnd.Items.Add('并且');
    cmbAnd.Items.Add('或者');
    cmbAnd.ItemIndex:=0;
  end;

  cmbFieldName:=TCombobox.Create(self);
  cmbFieldName.Parent:=Panel;
  cmbFieldName.Width:=109;cmbFieldName.Left:=77;cmbFieldName.Top:=8;
  cmbFieldName.Tag:=2;
  cmbFieldName.Items:=slFieldNames;
  cmbFieldName.OnChange:=cmbFieldNameOnChange;
  for i := 0 to cmbFieldName.Items.Count-1 do//add by liuying 20100821
  begin
    j:=pos('1z2y3x',lowercase(cmbFieldName.Items.Strings[i]));
    if j>0 then cmbFieldName.Items.Strings[i]:=copy(cmbFieldName.Items.Strings[i],1,j-1);
  end;
  cmbFieldName.ItemIndex:=0;

  cmbAmount:=TCombobox.Create(self);
  cmbAmount.Parent:=Panel;
  cmbAmount.Width:=74;cmbAmount.Left:=196;cmbAmount.Top:=8;
  cmbAmount.Tag:=3;
  cmbAmount.Items.Add('等于');
  cmbAmount.Items.Add('大于');
  cmbAmount.Items.Add('小于');
  cmbAmount.Items.Add('不等于');
  cmbAmount.Items.Add('大于等于');
  cmbAmount.Items.Add('小于等于');
  cmbAmount.Items.Add('包含');
  cmbAmount.Items.Add('不包含');
  cmbAmount.ItemIndex:=0;

  CreateAValueControl(Panel,ArFieldType[cmbFieldName.ItemIndex],slFieldNames[cmbFieldName.ItemIndex]);
end;

procedure TfrmADOLYQuery.CreateAValueControl(AParentControl: TWinControl;
  AFieldType: TFieldType;AFieldName:string);
var
  edtValue:TEdit;
  cbbValue:TComboBox;
  dtpValue:TDateTimePicker;
  
  i,k:integer;
  JhPos:integer;
  sItems:string;
  sList:TStrings;
  adotemp11:tadoquery;
begin
  //先删除相应的结果框
  for  i:=0  to AParentControl.ControlCount-1 do
    if AParentControl.Controls[I].Tag=4 then//tag:4--结果框,1--并且或者框,2--字段框,3--等于框
      AParentControl.Controls[i].Free;
  //-----------------
  
  case AFieldType of
        ftDate,ftTime,ftDateTime:
        begin
          dtpValue:=TDateTimePicker.Create(self);
          dtpValue.Parent:=AParentControl;
          dtpValue.Left:=280;dtpValue.Width:=120;dtpValue.Top:=8;
          dtpValue.Tag:=4;
          if pos('FIELDTYPE1T2S3RFTTIME',uppercase(AFieldName))>0 then dtpValue.Kind:=dtkTime;//add by liuying 20100821
        end;
        ftSmallint, ftInteger, ftWord,ftAutoInc:
        begin
          edtValue:=TEdit.Create(self);
          edtValue.Parent:=AParentControl;
          edtValue.Left:=280;edtValue.Width:=120;edtValue.Top:=8;
          edtValue.Tag:=4;
          SetWindowLong(edtValue.Handle, GWL_STYLE,GetWindowLong(edtValue.Handle, GWL_STYLE) or ES_NUMBER);
        end else
        begin
          JhPos:=Pos('COMBOBOXITEMS1T2S3R',uppercase(AFieldName));
          if JhPos>0 then//add by liuying 20100821
          begin
            cbbValue:=TComboBox.Create(self);
            cbbValue.Parent:=AParentControl;
            cbbValue.Left:=280;cbbValue.Width:=120;cbbValue.Top:=8;
            cbbValue.Tag:=4;

            sItems:=trim(copy(AFieldName,JhPos+19,maxint));
            k:=pos('1w2v3u',sItems);
            if k>0 then sItems:=copy(sItems,1,k-1);
            sItems:=StringReplace(sItems,'1T2S3R',#$2,[rfReplaceAll,rfIgnoreCase]);
            sList:=TStringList.Create;
            ExtractStrings([#$2],[],PChar(sItems),sList);
            if sList.Count=1 then//视图名
            begin
              adotemp11:=tadoquery.Create(nil);
              adotemp11.Connection:=pConnection;
              adotemp11.Close;
              adotemp11.SQL.Clear;
              adotemp11.SQL.Text:='select * from '+sList[0];
              try
                adotemp11.Open;//只有打开的情况下才能检测到字段
              except
                raise Exception.Create('打开视图'+sList[0]+'出错!');
                adotemp11.Free;
                sList.Free;
                exit;
              end;
              while not adotemp11.Eof do
              begin
                cbbValue.Items.Add(adotemp11.Fields[0].value);
                adotemp11.Next;
              end;
              adotemp11.Free;
            end else for i :=0 to sList.Count-1 do cbbValue.Items.Add(sList[i]);
            sList.Free;
          end ELSE
          begin
            edtValue:=TEdit.Create(self);
            edtValue.Parent:=AParentControl;
            edtValue.Left:=280;edtValue.Width:=120;edtValue.Top:=8;
            edtValue.Tag:=4;
          end;
        end;
  end;
end;

procedure TfrmADOLYQuery.FormShow(Sender: TObject);
var
  TLYQueryini:Tinifile;
  ifSave:boolean;
  QueryConditionCount,i,j:integer;
  ss:string;
begin
  pResult:=false;

  ArFieldType:=GetFieldType(pSelectString);
  slFieldNames:=GetFieldNames(pSelectString);
  slFieldsList:=GetFieldsList(pSelectString,slFieldNames);

  //==========加载保存的查询条件==============================================//
  if (csDesigning in ComponentState) then exit;

  TLYQueryini:=tinifile.create('.\TAdoLYQuery.ini');
  ifSave:=TLYQueryini.ReadBool('interface','ifSave',false);
  if not ifSave then begin TLYQueryini.Free;exit;end;
  CheckBox1.Checked:=ifSave;
  QueryConditionCount:=TLYQueryini.ReadInteger('interface','QueryConditionCount',0);

  for i :=1  to QueryConditionCount do
  begin
    if ScrollBox1.ControlCount=0 then CreateAQueryControl(true)
      else CreateAQueryControl(false);
  end;

  for i :=0  to ScrollBox1.ControlCount-1 do
  begin
    for j :=0  to TPanel(ScrollBox1.Controls[i]).ControlCount-1 do//每个panel
    begin
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=1 then
        TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex:=TLYQueryini.ReadInteger('interface','logicExp'+inttostr(i),0);
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=2 then
      begin
        TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex:=TLYQueryini.ReadInteger('interface','sField'+inttostr(i),0);
        CreateAValueControl(TPanel(ScrollBox1.Controls[i]),ArFieldType[TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex],slFieldNames[TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex]);
      end;
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=3 then
        TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex:=TLYQueryini.ReadInteger('interface','mathExp'+inttostr(i),0);
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=4 then
      begin
        ss:=TLYQueryini.ReadString('interface','sValue'+inttostr(i),'');
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TEdit then
          TEdit(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text:=ss;
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TComboBox then//add by liuying 20100821
          TComboBox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text:=ss;
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TDateTimePicker then
          TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).DateTime:=StrToDateTimeDef(ss,date);
      end;
    end;
  end;

  TLYQueryini.Free;
  //==========================================================================//}
end;

function TfrmADOLYQuery.GetSqlTxt: string;
var
  logicExp,sField,mathExp,sValue:String;
  GroupBy,OrderBy:string;
  i,j:INTEGER;
  FieldType:TFieldType;

  iPos,frompos:integer;
  tempStr:string;
begin
  result:=GetSelectPart(pSelectString,OrderBy,GroupBy);

  //add by liuying 20100821.删除查询出来的标签中的属性值
  frompos:=LastPos(' from ',pchar(LowerCase(result)));
  iPos:=pos('1z2y3x',lowercase(result));
  while iPos>0 do
  begin
    tempStr:=copy(result,iPos,maxint);
    if (pos(',',tempStr)>0) and (pos(',',tempStr)<frompos) then delete(result,iPos,pos(',',tempStr)-1)
      else if (pos(' ',tempStr)>0) and (pos(' ',tempStr)<frompos) then delete(result,iPos,pos(' ',tempStr)-1)
        else delete(result,iPos,maxint);//当语句写错时才会出现此情况.写此句是为了避免死循环

    iPos:=pos('1z2y3x',lowercase(result));
  end;
  //====================================================

  FieldType:=ftString;
  for i :=0  to ScrollBox1.ControlCount-1 do//全部为TPanel
  begin
      for j :=0  to TPanel(ScrollBox1.Controls[i]).ControlCount-1 do//每个panel
      begin
        if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=1 then
        begin
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='并且' then logicExp:='and';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='或者' then logicExp:='or';
        end;
        if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=2 then//字段
        begin
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex=-1 then begin raise Exception.Create('请设置正确的查询条件!');exit;end;
          sField:=slFieldsList[TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex];
          FieldType:=ArFieldType[TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex];
        end;
        if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=3 then
        begin
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='等于' then mathExp:='=';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='小于' then mathExp:='<';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='大于' then mathExp:='>';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='不等于' then mathExp:='<>';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='大于等于' then mathExp:='>=';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='小于等于' then mathExp:='<=';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='包含' then mathExp:=' like ';
          if TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text='不包含' then mathExp:=' not like ';
        end;
        if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=4 then
        begin
          if TPanel(ScrollBox1.Controls[i]).Controls[j] is TEdit then
            sValue:=TEdit(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text;
          if TPanel(ScrollBox1.Controls[i]).Controls[j] is TComboBox then//add by liuying 20100821
            sValue:=TComboBox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text;
          if(TPanel(ScrollBox1.Controls[i]).Controls[j] is TDateTimePicker)and(TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).Kind=dtkDate)then
            sValue:=FormatDateTime('YYYY-MM-DD',TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).Date);
          if(TPanel(ScrollBox1.Controls[i]).Controls[j] is TDateTimePicker)and(TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).Kind=dtkTime)then
            sValue:=FormatDateTime('hh:nn:ss',TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).Time);
          if((FieldType=ftString)or(FieldType=ftWideString))and(sValue='')and(pDataBaseType=dbtOracle)then//add by liuying 20100825//20130915 ArFieldType[j]->FieldType
            sValue:='!@#$%';
          if (mathExp=' like ')or(mathExp=' not like ') then sValue:='%'+sValue+'%';
          case FieldType of
                ftUnknown:sValue:=sValue;
                ftString,ftWideString,ftDate,ftTime,ftDateTime:sValue:=''''+sValue+'''';//增加ftDate,ftTime,ftDateTime by LY 20090821
                //ftDate,ftTime,ftDateTime:begin if pDataBaseType=dbtAccess THEN sValue:='#'+sValue+'#' else sValue:=''''+sValue+'''';end;//注释 by LY 20090821
                ftSmallint,ftInteger,ftWord,ftFloat,ftCurrency,ftAutoInc:sValue:=sValue;
          else sValue:=sValue;
          end;
        end;
      end;
      result:=result+' '+logicExp+' '+sField+mathExp+sValue ;
  end;
  result:=result+' '+GroupBy+' '+OrderBy;
end;

procedure TfrmADOLYQuery.BitBtnCommQryClick(Sender: TObject);
var
  ADOTEMP11:TADOQUERY;
  tmpSelString:string;
  a,b:string;//无实际用处
  i,j:integer;
  s1:string;
begin
  if (not CheckBox2.Checked)and(ScrollBox1.ControlCount>0) then
    pResultSelect:=GetSqlTxt
  else begin//增加全部查询功能
    pResultSelect:=pSelectString;
    //add by liuying 20110217.删除字段的属性值
    for i :=0  to slFieldNames.Count-1 do
    begin
      j:=Pos('1z2y3x',LowerCase(slFieldNames[i]));
      if j>0 then
      begin
        s1:=copy(slFieldNames[i],j,maxint);
        pResultSelect:=StringReplace(pResultSelect,s1,'',[rfReplaceAll,rfIgnoreCase]);
      end;
    end;
    //========================================
  end;

  ADOTEMP11:=tadoquery.Create(nil);
  ADOTEMP11.Connection:=pConnection;
  ADOTEMP11.Close;
  ADOTEMP11.SQL.Clear;

  tmpSelString:=GetSelectPart(pSelectString,a,b);
  tmpSelString:=tmpSelString+' 1=0 ';
  ADOTEMP11.SQL.Text:=tmpSelString;

  try
    ADOTEMP11.Open;
    pResult:=true;
  finally
    ADOTEMP11.Free;
  end;
  Close;
end;

procedure TfrmADOLYQuery.btn_addClick(Sender: TObject);
begin
  if ScrollBox1.ControlCount=0 then CreateAQueryControl(true)
    else CreateAQueryControl(false);
end;

procedure TfrmADOLYQuery.btn_decClick(Sender: TObject);
begin
  if ScrollBox1.ControlCount>0 then
    TPanel(ScrollBox1.Controls[ScrollBox1.ControlCount-1]).Free;
end;

procedure TfrmADOLYQuery.BitBtnCommQryCloseClick(Sender: TObject);
begin
  close;
end;

procedure TfrmADOLYQuery.FormDestroy(Sender: TObject);
var
  TLYQueryini:Tinifile;
  i,j:integer;
begin
  slFieldNames.Free;
  slFieldsList.Free;
  pConnection.Free;
  
  //==========保存查询条件====================================================//
  if (csDesigning in ComponentState) then exit;
  
  TLYQueryini:=tinifile.create('.\TAdoLYQuery.ini');
  TLYQueryini.WriteBool('interface','ifSave',CheckBox1.Checked);
  TLYQueryini.WriteInteger('interface','QueryConditionCount',ScrollBox1.ControlCount);
  for i :=0  to ScrollBox1.ControlCount-1 do
  begin
    for j :=0  to TPanel(ScrollBox1.Controls[i]).ControlCount-1 do//每个panel
    begin
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=1 then
        TLYQueryini.WriteInteger('interface','logicExp'+inttostr(i),TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex);
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=2 then
        TLYQueryini.WriteInteger('interface','sField'+inttostr(i),TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex);
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=3 then
        TLYQueryini.WriteInteger('interface','mathExp'+inttostr(i),TCombobox(TPanel(ScrollBox1.Controls[i]).Controls[j]).ItemIndex);
      if TPanel(ScrollBox1.Controls[i]).Controls[j].Tag=4 then
      begin
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TEdit then
          TLYQueryini.WriteString('interface','sValue'+inttostr(i),TEdit(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text);
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TComboBox then//add by liuying 20100821
          TLYQueryini.WriteString('interface','sValue'+inttostr(i),TComboBox(TPanel(ScrollBox1.Controls[i]).Controls[j]).Text);
        if TPanel(ScrollBox1.Controls[i]).Controls[j] is TDateTimePicker then
          TLYQueryini.WriteDateTime('interface','sValue'+inttostr(i),TDateTimePicker(TPanel(ScrollBox1.Controls[i]).Controls[j]).DateTime);
      end;
    end;
  end;
  TLYQueryini.Free;
  //==========================================================================//}
end;

procedure TfrmADOLYQuery.CheckBox2Click(Sender: TObject);
begin
  ScrollBox1.Enabled:=not tcheckbox(sender).Checked;//增加全部查询功能
end;

{ TADOLYQuery }

constructor TADOLYQuery.create(aowner: tcomponent);
begin
  inherited Create(AOwner);
  FDataBaseType:=dbtMSSQL;
end;

destructor TADOLYQuery.destroy;
begin
  inherited Destroy;
end;

function TADOLYQuery.Execute: boolean;
begin
  if fConnection=nil then
  begin
    raise Exception.Create('没有设置连接属性!');  
    result:=false;
    exit;
  end;
  
  ffrmLYQuery:=TfrmADOLYQuery.Create(nil);

  ffrmLYQuery.pSelectString:=fSelectString;
  ffrmLYQuery.pDataBaseType:=FDataBaseType;

  ffrmLYQuery.pConnection:=tAdoconnection.Create(nil);
  ffrmLYQuery.pConnection.ConnectionString:=fConnection.ConnectionString;
  ffrmLYQuery.pConnection.LoginPrompt:=false;

  ffrmLYQuery.ShowModal;
  fResultSelect:=ffrmLYQuery.pResultSelect;
  result:=ffrmLYQuery.pResult;
  ffrmLYQuery.Free;
end;

procedure TADOLYQuery.FSetConnection(value: TADOConnection);
begin
  if value=FConnection then exit;
  FConnection:=value;
end;

procedure TADOLYQuery.FSetDataBaseType(value: TDataBaseType);
begin
  if value=FDataBaseType then exit;
  FDataBaseType:=value;
end;

procedure TADOLYQuery.FSetSelectString(value: string);
begin
  if value=FSelectString then exit;
  FSelectString:=value;
end;

end.
