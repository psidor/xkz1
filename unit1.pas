unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, OleServer, ExcelXP, Grids,ComObj,GenPersonQrcodeStr,UnitGwjjbIdCard,StrUtils,
  ExtDlgs, ComCtrls;

type
  TForm1 = class(TForm)
    SpeedButton1: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Panel1: TPanel;
    OpenDialog1: TOpenDialog;
    ExcelApplication1: TExcelApplication;
    Label7: TLabel;
    SpeedButton2: TSpeedButton;
    ScrollBox1: TScrollBox;
    SpeedButton3: TSpeedButton;
    BitBtn1: TBitBtn;
    OpenPictureDialog1: TOpenPictureDialog;
    ProgressBar1: TProgressBar;
    Label5: TLabel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
   procedure DrawCards;
   procedure ChooseAphoto(Sender:TObject);
   Procedure DrawACardFront(i:integer);
   Procedure DrawACardBack(i:integer);
   procedure setPersonInfsFromExcel(excelFile:string);
  end;


Function getCardType(aXMBBM:string):integer;




const TableKEY:Array[1..17] of string =('身份证','项目部编码','文化程度',
                                        '民族','急电话','急联系人关系',
                                        '血型','体检结果','体检时间','考试成绩',
                                        '进场','机构代码','职位工种编码','姓名',
                                        '急联系人', '单位名称','项目部名称');



var
  Form1: TForm1;
  ApPersonInfs:pPersonInfs;
  personNber:integer;
  PIXC:TPersonInfxlsCol;
  CheckBoxS:array of TCheckBox;
  CardImgs:array of TImage;
  FacePhotoesFile:array of string;
  aIdCard:TgwjjbIdCardFront;
  aIdCardB:TgwjjbIdCardBack;
  curPersonSeqI:integer;
  OpenRightXls:Boolean;




implementation


{$R *.dfm}

procedure TForm1.SpeedButton1Click(Sender: TObject);

begin
if openDialog1.Execute then
begin
  aIdCard:=TgwjjbIdCardFront.Create(self);
  aIdCardB:=TgwjjbIdCardBack.Create(self);
  OpenRightXls:=true;
  setPersonInfsFromExcel(openDialog1.FileName);
  if  OpenRightXls then
  begin

  DrawCards;
  panel1.Visible:=false;
  SpeedButton3.Caption:='选择多个登记照'+#10+'（请将登记照命名为类似“张三.jpg”根据姓名自动匹配）';
   SpeedButton2.Caption:='保存图片送印刷'+#10+'（图片保存在桌面上的“二维码卡标”目录中）';

  SpeedButton3.Visible:=true;
  SpeedButton2.Visible:=true;
  end;
  end;
end;


procedure TForm1.setPersonInfsFromExcel(excelFile:string);
var
    ExcelApp:Variant;
    i,j,ir,ic:integer;
    excelCols,excelRows,fieldTitleRow:integer;
    xlFieldName:array of string;
    AFieldName:string;
begin
   initialPersonInfxlsCol(PIXC);
   try
   begin
   
   ExcelApp:= CreateOleObject ('Excel.Application') ;
   ExcelApp.WorkBooks.Open (excelFile) ;
   ExcelApp.WorkSheets[1].Activate;
    //showmessage(ExcelApp.Cells [1,1].Value);
    if not ExcelApp.ActiveSheet.ProtectContents then
    begin
    OpenRightXls:=false;
    showmessage('打开的不是系统导出的excel文件，'+#10+'或本程序已经过期,请重新下载新程序！');
    end
    else
    begin

   ExcelApp.ActiveSheet.Unprotect('8fd352063e5e9297');
   ExcelCols:=ExcelApp.ActiveSheet.UsedRange.Columns.Count;
   excelRows:=ExcelApp.ActiveSheet.UsedRange.Rows.Count;

   Label5.visible:=true;
   ProgressBar1.Visible:=True;
   ProgressBar1.Max:=excelRows;

   fieldTitleRow:=0;
   ir:=1;
   ic:=1;
   while  (fieldTitleRow=0) and (ir<=ExcelRows) do
    begin
      while  (fieldTitleRow=0) and (ic<=ExcelCols) do
          begin
          if  ExcelApp.Cells [ir,ic].Value='姓名' then
            fieldTitleRow:=ir
          else
            ic:=ic+1;
          end;
       ic:=1;
       ir:=ir+1;
    end;
    //匹配各列
    setLength(xlFieldName,ExcelCols+1);

    for i:=1 to ExcelCols do
    begin
       AFieldName:=ExcelApp.Cells[fieldTitleRow,i].value;
       //去除字符串中的回车
       //xlFieldName[i]:=trimStrReturn(xlFieldName[i]);
       AFieldName:=StringReplace(AFieldName, #10, '', [rfReplaceAll]);
       //showmessage(xlFieldName[i]);
       if pos('身份证',AFieldName)>0 then PIXC.clSfzHao:=i
       else if pos('项目部编码',AFieldName)>0 then PIXC.clXiangMuId:=i
       else if pos('文化',AFieldName)>0 then PIXC.clWenHuaChengdu:=i
       else if pos('民族',AFieldName)>0 then PIXC.clMingZu:=i
       else if pos('急电话',AFieldName)>0 then PIXC.clYingJiDianHua:=i
       else if pos('急联系人关系',AFieldName)>0 then PIXC.clYingJiLianXiRenGuanXi:=i
       else if pos('血型',AFieldName)>0 then PIXC.clXueXing:=i
       else if pos('体检结果',AFieldName)>0 then PIXC.clTiJianJieGuo:=i
       else if pos('体检时间',AFieldName)>0 then PIXC.clTiJianShiJian:=i
       else if pos('考试时间',AFieldName)>0 then PIXC.clKaoShiShiJian:=i
       else if (pos('安全考试成绩',AFieldName)>0) or (pos('安全考试情况',AFieldName)>0) then PIXC.clKaoShiChengJi:=i
       else if pos('进场',AFieldName)>0 then PIXC.clJingChangRiQi:=i
       else if pos('组织',AFieldName)>0 then PIXC.clZuZhiDaiMa:=i
       else if pos('工种编码',AFieldName)>0 then PIXC.clZhiWuGongZhongBianMa:=i
       else if pos('姓名',AFieldName)>0 then PIXC.clXinMing:=i
       else if pos('急联系人',AFieldName)>0 then PIXC.clYingJiLianXiRenxing:=i
       else if pos('单位名称',AFieldName)>0 then PIXC.clDanWeiMingCheng:=i
       else if pos('项目部名称',AFieldName)>0 then PIXC.clXmMingCheng:=i
       else if (pos('人员类别',AFieldName)>0) or (pos('显示岗位',AFieldName)>0) then PIXC.clRenYuanLeiBie:=i;
    end;

    new(ApPersonInfs);
    personNber:=0;
    setlength(ApPersonInfs^,ExcelRows-fieldTitleRow+1);
    for i:=fieldTitleRow+1 to  ExcelRows  do
    begin
      //initialPersonInf(PersonInfs[i-fieldTitleRow]);
        //初始化
      ApPersonInfs^[i-fieldTitleRow].AURL:='m?p=';
      ApPersonInfs^[i-fieldTitleRow].ASfzHao:='000000000000000000';
      ApPersonInfs^[i-fieldTitleRow].AXiangMuId:='SG-00000000000000-00';
      ApPersonInfs^[i-fieldTitleRow].AWenHuaChengdu:='0';
      ApPersonInfs^[i-fieldTitleRow].AMingZu:='汉族';
      ApPersonInfs^[i-fieldTitleRow].AYingJiDianHua:='000000000000';
      ApPersonInfs^[i-fieldTitleRow].AYingJiLianXiRenGuanXi:='不确定';
      ApPersonInfs^[i-fieldTitleRow].AXueXing:='不确定';
      ApPersonInfs^[i-fieldTitleRow].ATiJianJieGuo:='不确定';
      ApPersonInfs^[i-fieldTitleRow].ATiJianShiJian:='700101';
      ApPersonInfs^[i-fieldTitleRow].AKaoShiShiJian:='070101';
      ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi:='00';
       ApPersonInfs^[i-fieldTitleRow].AJingChangRiQi:='070101';
      ApPersonInfs^[i-fieldTitleRow].AZuZhiDaiMa:='000000000000000000';
      ApPersonInfs^[i-fieldTitleRow].AZhiWuGongZhongBianMa:='0000';
      ApPersonInfs^[i-fieldTitleRow].AXinMing:='张三';
      ApPersonInfs^[i-fieldTitleRow].AYingJiLianXiRenxing:='';
      ApPersonInfs^[i-fieldTitleRow].ADanWeiMingCheng:='';
      ApPersonInfs^[i-fieldTitleRow].AXmMingCheng:='';
      ApPersonInfs^[i-fieldTitleRow].APhotoJpgFile:='';

     //赋值
      if PIXC.clSfzHao>0 then ApPersonInfs^[i-fieldTitleRow].ASfzHao:=ExcelApp.Cells [i,PIXC.clSfzHao].Value;
      if PIXC.clXiangMuId>0 then ApPersonInfs^[i-fieldTitleRow].AXiangMuId:=ExcelApp.Cells [i,PIXC.clXiangMuId].Value;
      if PIXC.clWenHuaChengdu>0 then ApPersonInfs^[i-fieldTitleRow].AWenHuaChengdu:=ExcelApp.Cells [i,PIXC.clXiangMuId].Value;
      if PIXC.clMingZu>0 then ApPersonInfs^[i-fieldTitleRow].AMingZu:=ExcelApp.Cells [i,PIXC.clMingZu].Value;
      if PIXC.clYingJiDianHua>0 then ApPersonInfs^[i-fieldTitleRow].AYingJiDianHua:=ExcelApp.Cells [i,PIXC.clYingJiDianHua].Value;
      if PIXC.clYingJiLianXiRenGuanXi>0 then ApPersonInfs^[i-fieldTitleRow].AYingJiLianXiRenGuanXi:=ExcelApp.Cells [i,PIXC.clYingJiLianXiRenGuanXi].Value;
      if PIXC.clXueXing>0 then ApPersonInfs^[i-fieldTitleRow].AXueXing:=ExcelApp.Cells [i,PIXC.clXueXing].Value;
      if PIXC.clTiJianJieGuo>0 then ApPersonInfs^[i-fieldTitleRow].ATiJianJieGuo:=ExcelApp.Cells [i,PIXC.clTiJianJieGuo].Value;
      if PIXC.clTiJianShiJian>0 then ApPersonInfs^[i-fieldTitleRow].ATiJianShiJian:=ArrangeDateStr(ExcelApp.Cells [i,PIXC.clTiJianShiJian].Text);

      if PIXC.clKaoShiShiJian>0 then ApPersonInfs^[i-fieldTitleRow].AKaoShiShiJian:=ArrangeDateStr(ExcelApp.Cells [i,PIXC.clKaoShiShiJian].Text);

      if PIXC.clKaoShiChengJi>0 then ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi:=ExcelApp.Cells [i,PIXC.clKaoShiChengJi].Text;
      if ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi='未考试' then ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi:='00';
      if ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi='合格' then ApPersonInfs^[i-fieldTitleRow].AKaoShiChengJi:='80';



      if PIXC.clJingChangRiQi>0 then ApPersonInfs^[i-fieldTitleRow].AJingChangRiQi:=ArrangeDateStr(ExcelApp.Cells [i,PIXC.clJingChangRiQi].Text);
      //showmessage(strtodate(ApPersonInfs^[i-fieldTitleRow].AJingChangRiQi,'yymmdd'));

      if PIXC.clZuZhiDaiMa>0 then ApPersonInfs^[i-fieldTitleRow].AZuZhiDaiMa:=ExcelApp.Cells [i,PIXC.clZuZhiDaiMa].Value;

      if PIXC.clZhiWuGongZhongBianMa>0 then
      ApPersonInfs^[i-fieldTitleRow].AZhiWuGongZhongBianMa:=ExcelApp.Cells [i,PIXC.clZhiWuGongZhongBianMa].Value
      else
        if PIXC.clRenYuanLeiBie>0 then ApPersonInfs^[i-fieldTitleRow].AZhiWuGongZhongBianMa:=
            convertRenYuanLeiBieS2GongZhongBianMa(leftstr(ApPersonInfs^[i-fieldTitleRow].AXiangMuId,2), ExcelApp.Cells [i,PIXC.clRenYuanLeiBie].Value);

      if PIXC.clXinMing>0 then ApPersonInfs^[i-fieldTitleRow].AXinMing:=ExcelApp.Cells [i,PIXC.clXinMing].Value;
      if PIXC.clYingJiLianXiRenxing>0 then ApPersonInfs^[i-fieldTitleRow].AYingJiLianXiRenxing:=ExcelApp.Cells [i,PIXC.clYingJiLianXiRenxing].Value;
      if PIXC.clDanWeiMingCheng>0 then ApPersonInfs^[i-fieldTitleRow].ADanWeiMingCheng:=ExcelApp.Cells [i,PIXC.clDanWeiMingCheng].Value;
      if PIXC.clXmMingCheng>0 then ApPersonInfs^[i-fieldTitleRow].AXmMingCheng:=ExcelApp.Cells [i,PIXC.clXmMingCheng].Value;
       personNber:=personNber+1;
       ProgressBar1.position:=i;
    end;
    end;
   end;
    except
  showmessage('打开的不是系统导出的excel文件，'+#10+'或本程序已经过期,请重新下载新程序！');
  OpenRightXls:=false;
  end;




end;


procedure TForm1.DrawCards;
var i,s:integer;
    astr1:string;

begin
  s:=5;
  ScrollBox1.Top:=10;
  ScrollBox1.Left:=10;
  ScrollBox1.Width:=Form1.Width-20;
  ScrollBox1.Height:=Form1.Height-120;
  ScrollBox1.Visible:=True;

  setLength(CardImgs,personNber+1);
  setLength(FacePhotoesFile,personNber+1);

  for i:=1 to personNber do
    begin
      CardImgs[i]:=TImage.Create(self);
      CardImgs[i].Width:=85*s;
      CardImgs[i].Height:=53*s;
      CardImgs[i].Top:=((i-1) div 2)*53*s+((i-1) div 2)*10+10;
      if ((i mod 2) =0) then CardImgs[i].Left:=20+CardImgs[i].Width  else CardImgs[i].Left:=10;
      CardImgs[i].Stretch:=true;
      //draw card
         DrawACardFront(i);
        CardImgs[i].Picture.Assign(aIdCard.Picture.Bitmap);
        CardImgs[i].Hint:=inttostr(i);
        CardImgs[i].OnClick:=ChooseAphoto;

      //CardImgs[i].Picture.LoadFromFile('C:\Users\Administrator\Desktop\zsf.bmp');
      CardImgs[i].Parent:=ScrollBox1;
    end;





end;






procedure TForm1.BitBtn1Click(Sender: TObject);
begin
//保存卡照
if not DirectoryExists(GetDeskeptPath+'\二维码卡标') then
//if not DirectoryExists(Edit1.Text) then 判断目录是否存在
  try
    begin
      ForceDirectories(GetDeskeptPath+'\二维码卡标');
    //ForceDirectories(Edit1.Text);   创建目录
    end
  finally
    //raise Exception.Create('Cannot Create'+ GetDeskeptPath+'\二维码卡标');
  end;
end;


procedure TForm1.ChooseAphoto(Sender: TObject);
var i:integer;
begin
  //showmessage(TImage(Sender).Hint);
  OpenPictureDialog1.Options:=OpenPictureDialog1.Options-[ofAllowMultiSelect];
  i:=strtoint(TImage(Sender).Hint);
  if OpenPictureDialog1.Execute then
    begin
    ApPersonInfs^[i].APhotoJpgFile:=OpenPictureDialog1.Files.Strings[0];
    DrawACardFront(i);
    CardImgs[i].Picture.Assign(aIdCard.Picture.Bitmap);
    end;



end;


Procedure TForm1.DrawACardFront(i:integer);
var astr1:string;
begin
  astr1:=GenAPersonQrcodeStr(ApPersonInfs^[i].AURL,
                                 ApPersonInfs^[i].ASfzHao,
                                 ApPersonInfs^[i].AXiangMuId,
                                 ApPersonInfs^[i].AWenHuaChengdu,
                                 ApPersonInfs^[i].AMingZu,
                                 ApPersonInfs^[i].AYingJiDianHua,
                                 ApPersonInfs^[i].AYingJiLianXiRenGuanXi,
                                 ApPersonInfs^[i].AXueXing,
                                 ApPersonInfs^[i].ATiJianJieGuo,
                                 ApPersonInfs^[i].ATiJianShiJian,
                                 ApPersonInfs^[i].AKaoShiShiJian,
                                 ApPersonInfs^[i].AKaoShiChengJi,
                                 ApPersonInfs^[i].AJingChangRiQi,
                                 ApPersonInfs^[i].AZuZhiDaiMa,
                                 ApPersonInfs^[i].AZhiWuGongZhongBianMa,
                                 ApPersonInfs^[i].AXinMing,
                                 ApPersonInfs^[i].AYingJiLianXiRenxing,
                                 ApPersonInfs^[i].ADanWeiMingCheng,
                                 ApPersonInfs^[i].AXmMingCheng);
        aIdCard.QrcodeStr:=astr1;
        aIdCard.CompanyName:=ApPersonInfs^[i].ADanWeiMingCheng;
        aIdCard.ProjectName:=ApPersonInfs^[i].AXmMingCheng;
        aIdCard.CardType:= getCardType(ApPersonInfs^[i].AXiangMuId);
        aIdCard.Title:=getGangWeiFromBianMa(leftstr(ApPersonInfs^[i].AZhiWuGongZhongBianMa,4));
        aIdCard.PName:=ApPersonInfs^[i].AXinMing;
        aIdCard.PhotoFileName:=ApPersonInfs^[i].APhotoJpgFile;
        aIdCard.InfOk:=True;
        if(ApPersonInfs^[i].ASfzHao='000000000000000000') or
        (ApPersonInfs^[i].ASfzHao='') or
        ( ApPersonInfs^[i].AXiangMuId='') or
        ( ApPersonInfs^[i].AXiangMuId='SG-00000000000000-00') then
            aIdCard.InfOk:=False;


        aIdCard.drawCard;

end;


Function getCardType(aXMBBM:string):integer;
begin
    if leftstr(aXMBBM,2)='YZ' then result:=0
    else if leftstr(aXMBBM,2)='JL' then result:=1
    else if leftstr(aXMBBM,2)='SG' then result:=2
    else result:=3;

end;

procedure TForm1.SpeedButton3Click(Sender: TObject);
var i,j,k:integer;
begin
//选择多张照片自动匹配
  OpenPictureDialog1.Title:='选择多张登记照以名字自动匹配';
  OpenPictureDialog1.Options:=OpenPictureDialog1.Options+[ofAllowMultiSelect];
  if OpenPictureDialog1.Execute then
    begin
      k:=OpenPictureDialog1.Files.Count;
      for i:=1 to personNber do
        for j:=0 to k-1 do
          if pos(ApPersonInfs^[i].AXinMing,OpenPictureDialog1.Files.Strings[j])>0 then
            begin
              ApPersonInfs^[i].APhotoJpgFile:= OpenPictureDialog1.Files.Strings[j];
              DrawACardFront(i);
              CardImgs[i].Picture.Assign(aIdCard.Picture.Bitmap);
            end;

    end;
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
var i:integer;
begin
//保存卡照
  if not DirectoryExists(GetDeskeptPath+'\二维码卡标') then
    try
      ForceDirectories(GetDeskeptPath+'\二维码卡标');
    finally
      //raise Exception.Create('Cannot Create 二维码卡标目录');
    end;

    aIdCard.SavePath:=GetDeskeptPath+'\二维码卡标\';
    aIdCardB.SavePath:=GetDeskeptPath+'\二维码卡标\';
    for i:=1 to personNber do
    begin
      DrawACardFront(i);
      if aIdCard.infOk and (ApPersonInfs^[i].APhotoJpgFile<>'') then
        begin
         aIdCard.saveJpg300Dpi;
         DrawACardBack(i);
         aIdCardB.saveJpg300Dpi;
        end; 
    end;

end;



Procedure TForm1.DrawACardBack(i:integer);
begin
   //aIdCardB.BkColor:=clBlue;
   aIdCardB.BarcodeStr:=GenAPersonBarcodeStr(
                        ApPersonInfs^[i].ASfzHao);
   aIdCardB.CardType:= getCardType(ApPersonInfs^[i].AXiangMuId);
   aIdCardB.PName:=ApPersonInfs^[i].AXinMing;
   aIdCardB.drawCard;
end;

end.
