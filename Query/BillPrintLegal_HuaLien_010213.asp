<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style13 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style14 {font-family:"標楷體"; font-size: 14px;}
.style15 {font-family:"標楷體"; font-size: 20px;}
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,5,439,50">
</object>
<%
'on Error Resume Next
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

Sys_DriverHomeAddress="":Sys_DriverHomeZip=""

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))

if Not rs.eof then
	Sys_Jurgeday=split(gArrDT(trim(rs("Jurgeday"))),"-")
else
	Sys_Jurgeday=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_Owner="":Sys_OwnerZip="":Sys_OwnerAddress="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

if Not rsfound.eof then

	If not ifnull(Sys_DriverID) and instr(trim(rsbil("BillNo")),"PA")<=0 and Not ifnull(trim(rsfound("Driver"))) Then
		Sys_Owner=trim(rsfound("Driver"))

	else
		Sys_Owner=trim(rsfound("Owner"))

	end if
end if

strSql="select DriverHomeZip,DriverHomeAddress,OwnerNotifyAddress from BillbaseDCIReturn where Carno in (select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

set rs1=conn.execute(strSQL)

if not rs1.eof then

	If not ifnull(Sys_DriverID) and instr(trim(rsbil("BillNo")),"PA")<=0 and Not ifnull(trim(rsfound("DriverHomeAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	
	elseIf not ifnull(rs1("OwnerNotifyAddress")) Then
		if Not rs1.eof then Sys_OwnerAddress=trim(rs1("OwnerNotifyAddress"))

	elseIf Not ifnull(trim(rsfound("OwnerAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	
	else
		if Not rs1.eof then Sys_OwnerAddress=trim(rs1("DriverHomeAddress"))
		if Not rs1.eof then Sys_OwnerZip=trim(rs1("DriverHomeZip"))

	End If 

	If ifnull(Sys_DriverHomeAddress) Then

		strSQL="Update Billbase set DriverZip='"&trim(rs1("DriverHomeZip"))&"',DriverAddress='"&trim(rs1("DriverHomeAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and DriverAddress is null"
		conn.execute(strSQL)

	end If 

else
	If not ifnull(Sys_DriverID) and instr(trim(rsbil("BillNo")),"PA")<=0 and Not ifnull(trim(rsfound("DriverHomeAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	
	elseIf Not ifnull(trim(rsfound("OwnerAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

	else
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	End if
end if

rs1.close

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerZipName) Then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
End if

If not ifnull(Sys_OwnerAddress) Then
	strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
	conn.execute(strSQL)

	strSQL="Update billbasedcireturn set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where billno='"&trim(rsbil("Billno"))&"' and exchangetypeid='W'"
	conn.execute(strSQL)
end If 

if ifnull(Sys_OwnerAddress) then
	response.write "<font size=""10"">"
    response.write rsbil("BillNo")&"入案地址為空白"
    response.Write "<br>請至舉發單資料繀護系統確認！！"
    response.write "</font>"
	response.end
end If 

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
end if

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex="":Sys_IMAGEFILENAME=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB,RECORDMEMBERID from BillBase where sn="&trim(rsbil("BillSN"))
set rssex=conn.execute(strSql)
if trim(Sys_BillTypeID)="1" then
	if Not rssex.eof then
		if trim(rssex("DriverSex"))="1" then
			Sys_Sex="男"
		else
			Sys_Sex="女"
		end if
	end if
end if

if Not rssex.eof then Sys_RecordMemberID=trim(rssex("RECORDMEMBERID"))

if Not rssex.eof then
	Sys_IllegalDate=split(gArrDT(trim(rssex("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end if
if Not rssex.eof then
	Sys_IllegalDate_h=hour(trim(rssex("IllegalDate")))
else
	Sys_IllegalDate_h=""
end if
if Not rssex.eof then
	Sys_IllegalDate_m=minute(trim(rssex("IllegalDate")))
else
	Sys_IllegalDate_m=""
end if
if Not rssex.eof then
	Sys_DealLineDate=split(gArrDT(trim(rssex("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if
if Not rssex.eof then
	Sys_DriverBirth=split(gArrDT(trim(rssex("DriverBirth"))),"-")
else
	Sys_DriverBirth=split(gArrDT(trim("")),"-")
end if
if Not rssex.eof then Sys_IMAGEFILENAME=trim(rssex("IMAGEFILENAME"))
'if Not rssex.eof then Sys_IMAGEFILENAMEB=trim(rssex("IMAGEFILENAMEB"))
if Not rssex.eof then Sys_IMAGEPATHNAME=trim(rssex("IMAGEPATHNAME"))
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

If Sys_BillFillerMemberID = "3480" Then
	Stop_IllegalDate=split(gArrDT(trim(rssex("IllegalDate"))),"-")
	Stop_IllegalDate_h=hour(trim(rssex("IllegalDate")))
	Stop_IllegalDate_m=minute(trim(rssex("IllegalDate")))

	strSQL="select DealLineDate from BillBase where CarNo='"&Sys_CarNo&"' and IllegalDate=to_date('"&gOutDT(gInitDT(rssex("IllegalDate")))&" "&hour(trim(rssex("IllegalDate")))&":"&minute(trim(rssex("IllegalDate")))&":00','YYYY/MM/DD/HH24:MI/SS') and ImageFileNameB is not null"

	set rsstop=conn.execute(strSQL)
	if Not rsstop.eof then Sys_IllegalDate=split(gArrDT(DateAdd("d",1,rsstop("DealLineDate"))),"-")
	if Not rsstop.eof then Sys_IllegalDate_h="00"
	if Not rsstop.eof then Sys_IllegalDate_m="00"
	rsstop.close
End If 

Sys_ImageFileNameA="":Sys_ImageFileNameB=""
strSQL="select * from BillIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	Sys_ImageFileNameA=trim(rsimage("ImageFileNameA"))
	Sys_ImageFileNameB=trim(rsimage("ImageFileNameB"))
end if

strSql="select a.LoginID,a.ChName,c.Content,b.UnitName,decode(b.Unitid,'Z000','A000',b.Unitid) Unitid,b.UnitTypeID,decode(b.Unitid,'Z000','1',b.UnitLevelID) UnitLevelID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerLogInID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_FillerUnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

chkJobID="":strUit=""
if trim(Sys_UnitLevelID)="3" then
	chkJobID="314,316"
	strUit="UnitID='"&Sys_FillerUnitID&"'"
elseif trim(Sys_UnitLevelID)="2" then
	chkJobID="304,305,318"
	strUit="UnitID in(select UnitID from UnitInfo where UnitTypeid=(select UnitTypeID from UnitInfo where UnitID='"&Sys_FillerUnitID&"') and UnitLevelID=2)"
elseif trim(Sys_UnitLevelID)="1" then
	chkJobID="303"
	strUit="UnitID='"&Sys_FillerUnitID&"'"
end if
		
strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and "&strUit&" and JobID in ("&chkJobID&")) a,(select ID,Content from Code where TypeID=4) b where a.JobID=b.ID order by ID"
set rsjob=conn.execute(strSQL)

if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.CreditID from UnitInfo a,Memberdata b where a.UnitID=b.Unitid and b.MemberID="&trim(Sys_RecordMemberID)

set Unit=conn.execute(strSQL)
Sys_UnitID=Unit("UnitID")
Sys_RedUnitName=Unit("UnitName")
Sys_UnitTypeID=Unit("UnitTypeID")
Sys_UnitLevelID=Unit("UnitLevelID")
Sys_Credit_ID=Unit("CreditID")
Unit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=replace(unit("UnitName"),"停車管理處","交通隊")
if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
unit.close

'strSQL="select UnitName,Address,Tel from UnitInfo where UnitID='"&Sys_BillUnitID&"'"
'set Unit=conn.execute(strSQL)
'If Not unit.eof Then
'SysUnitLevel3=Unit("UnitName")
'SysAddressLevel3=Unit("Address")
'SysUnitTelLevel3=Unit("Tel")
'end if
'Unit.close

strSQL="select Value from ApConfigUre where ID=40"
set City=conn.execute(strSQL)
Sys_City=City("Value")
City.close
Sys_IllegalRule1=""
if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		'Sys_Level1=trim(rsRule1("Level1"))
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if
rssex.close
Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		'Sys_Level1=trim(rsRule1("Level1"))
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

Sys_DCISTATIONID="":Sys_STATIONNAME="":Sys_StationTel="":StationID=""

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

if ifnull(Sys_DCISTATIONID) then
	response.write "<font size=""10"">"
    response.write rsbil("BillNo")&"為異常案件強制入案"
    response.Write "<br>故監理系統未回傳資料"
    response.Write "<br>請至舉發單資料維護系統修改應到案處所！！"
    response.write "</font>"
	response.end
end If 

strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
Sys_MailNumber=""
set rs=conn.execute(strSql)

if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
rs.close

strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&trim(rsbil("BillSN"))&" and a.CarNo='"&trim(rsbil("CarNo"))&"'"
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("Content")
	rsfast.movenext
wend
rsfast.close
if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

StrBass="select a.A_Name from BillBaseDCIReturn a,(select ID,Content from DCICode where TypeID=4) b,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DciReturnCarColor=b.ID(+) and a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+) and a.EXCHANGETYPEID='A' and a.CarNo='"&trim(rsbil("CarNo"))&"'"

Sys_A_Name="":Sys_CarColor=""
set rsCarType=conn.execute(strBass)
if not rsCarType.eof then Sys_A_Name=trim(rsCarType("A_Name"))
rsCarType.close

StrBass="select DciReturnCarColor from BillBaseDCIReturn where EXCHANGETYPEID='W' and CarNo='"&trim(rsbil("CarNo"))&"' and BillNo='"&trim(rsbil("BillNo"))&"'"
Sys_CarColor="":Sys_CarColorID=""
set rsCarType=conn.execute(strBass)
if not rsCarType.eof then Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
rsCarType.close
if IfNull(Sys_DciReturnCarColor) then Sys_DciReturnCarColor=""
if len(Sys_DciReturnCarColor)>1 then Sys_DciReturnCarColor=left(Sys_DciReturnCarColor,1)&","&right(Sys_DciReturnCarColor,1)

Sys_CarColorID=split(Sys_DciReturnCarColor,",")
for y=0 to ubound(Sys_CarColorID)
	if trim(Sys_CarColor)<>"" then Sys_CarColor=Sys_CarColor&","
	if trim(Sys_CarColorID(y))<>"" and not isnull(Sys_CarColorID(y)) then
		strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
		set rscolor=conn.execute(strColor)
		if not rscolor.eof then
			Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
		end if
		rscolor.close
	end if
next

if trim(Sys_BillTypeID)="1" then
	If ifnull(Sys_MailNumber) Then Sys_MailNumber="0"
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,0,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	DelphiASPObj.CreateBarCode Sys_MailNumber&"97300717",128,35,260
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	If ifnull(Sys_MailNumber) Then Sys_MailNumber=0	
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"97300717","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	DelphiASPObj.CreateBarCode Sys_MailNumber&"97300717",128,60,150
end if

Sys_FirstBarCode=Sys_BillNo

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
If Sys_OwnerZip="001" then Sys_OwnerZip=""
rsbil.close

pageTop=0
pageLeft=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer41" style="position:absolute; left:290px; top:0px; width:620px; height:12px; z-index:36"><%'=gInitDT(Sys_MailDate)
	Response.Write "地址："&Sys_UnitAddress &"&nbsp;&nbsp;&nbsp;&nbsp;" & "<b>請送回 花蓮縣警察局"&sysunit&"</b>"
%></div>

<div id="Layer01" class="style3" style="position:absolute; left:135px; top:10px; z-index:3"><B><%
	'response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write funcCheckFont(Sys_Owner,16,1)&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
%></B>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:135px; top:55px; z-index:2"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:155px; top:45px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"

	%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:350px; top:50px; z-index:1"><%
	If Sys_BillFillerMemberID = "3480" Then Response.Write "TG"
	%>
</div>

<div id="Layer44" class="style3" style="position:absolute; left:525px; top:150px; z-index:5">
	<%="<img src=""../BarCodeImage/"&Sys_MailNumber&"97300717"&".jpg"" height=""35"">"%><br>
    　<b><%=Sys_MAILCHKNUMBER%></b>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:390px; top:270px; z-index:1"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:130px; top:275px; z-index:1"><b><%
	if trim(Sys_BillTypeID)="1" then
		response.write Sys_Driver
	elseif trim(Sys_BillTypeID)="2" then
		response.write funcCheckFont(Sys_Owner,16,1)
	end if%>　台啟</b>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:130px; top:295px; width:350px; z-index:1"><b><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			if trim(Sys_BillTypeID)="1" then
				response.write Orz_DriverHomeZip&"　"
				response.write replace(Orz_DriverHomeZipName&Orz_DriverHomeAddress,Orz_DriverHomeZipName&Orz_DriverHomeZipName,Orz_DriverHomeZipName)
			elseif trim(Sys_BillTypeID)="2" then
				response.write Sys_OwnerZip&" "
				response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
			end if
			response.write "<br><br>"
	%></b>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:130px; top:325px; width:330px; z-index:3">
<%
	Response.Write "花蓮縣警察局"&sysunit&"<br>"
	Response.Write Sys_UnitAddress&"<br>"
	Response.Write Sys_UnitTel
%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:455px; top:282px; z-index:5">
	　<b>第<%=Sys_MailNumber%>號</b><br>
	<%="<img src=""../BarCodeImage/"&Sys_MailNumber&"97300717"&".jpg"">"%><br>
    　<b><%=Sys_MAILCHKNUMBER%></b>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:60px; top:355px; z-index:1"><%
	'Response.Write Sys_CarNo&"　"&Sys_A_Name&"　"&Sys_CarColor&"　"&Sys_STATIONNAME
	'Response.Write Sys_CarNo&"　"&Sys_STATIONNAME
	'Response.Write "　"&Sys_STATIONNAME
%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:230px; top:340px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:550px; top:375px; z-index:1"><%
	Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>
<!-------------------------- 判斷 billillegalimage 有沒有這些圖檔 ------------------------------>
<!---------- todo 依據法條判斷, 如果是闖紅燈, 要把 a 檔的 xxxxx_a.jpg 換成 b檔的xxxxxx_b.jpg ---------------------------------------->
<%
	if Sys_FillerUnitID="A002" then
		Response.Write "<div id=""Layer09"" style=""position:absolute; left:40px; top:465px; z-index:10"">"
		response.write "<img src=""./img/StopCar_HuaLien.jpg"" width=""680"">"
		Response.Write "</DIV>"
		
		Response.Write "<div id=""Layer18"" style=""position:absolute; left:140px; top:510px; width:100px; height:14px; z-index:11"">"&Sys_CarNo&"</div>"

		Response.Write "<div id=""Layer20"" style=""position:absolute; left:360px; top:510px; width:300px; height:17px; z-index:11"">"&funcCheckFont(Sys_Owner,16,1)&"</div>"

		Response.Write "<div id=""Layer32"" style=""position:absolute; left:140px; top:535px; width:400px; height:49px; z-index:11"">"
		response.write left(trim(Sys_Rule1),2)&"條 "
		response.write Mid(trim(Sys_Rule1),3,1)&"項 "&Mid(trim(Sys_Rule1),4,2)&"款"
		Response.Write "</div>"

		Response.Write "<div id=""Layer20"" style=""position:absolute; left:360px; top:535px; width:300px; height:17px; z-index:11"">"&Sys_BillNo&"</div>"

		Response.Write "<div id=""Layer29"" style=""position:absolute; left:590px; top:535px; width:300px; height:11px; z-index:22"">"&Sys_DealLineDate(0)&"年"&Sys_DealLineDate(1)&"月"&Sys_DealLineDate(2)&"日</div>"

	else
		if trim(Sys_ImageFileNameA)<>"" then
			Response.Write "<div id=""Layer09"" style=""position:absolute; left:40px; top:465px; z-index:5"">"
			response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
			Response.Write "</DIV>"

		elseIf trim(Sys_Rule1)="5620001" and trim(Sys_ImageFileName)<>"" Then
			Response.Write "<div id=""Layer09"" style=""position:absolute; left:40px; top:465px; z-index:5"">"
			response.write "<img src=""../StopCarPicture/"&Sys_ImageFileName&""" width=""390"" height=""280"">"
			Response.Write "</DIV>"
		End If 
		
		if trim(Sys_ImageFileNameB)<>"" and (Sys_Rule1 <>"6020303" and Sys_Rule2<>"6020303") then
			Response.Write "<div id=""Layer10"" style=""position:absolute; left:430px; top:465px; z-index:1"">"
			response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
			Response.Write "</DIV>"
		end if
	end if

'response.write Sys_Rule1 & "_" 
'response.write Sys_ImageFileNameB
' ssmith 20091015 紅燈月線不顯示B圖
%>
<%if trim(Sys_ImageFileNameB)<>"" and (Sys_Rule1 <>"6020303" and Sys_Rule2<>"6020303") then%>
	<div id="Layer10" style="position:absolute; left:430px; top:470px; z-index:1"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>


<%'1030605李小姐說要拿掉
'<div id="Layer44" class="style1" style="position:absolute; left:20px; top:760px; width:320px; height:36px; z-index:5">
'	自103年3月31日起，不服舉發者，應於接獲本單30日內，向處罰機關(機應到案處所)陳述；
'	受處罰人於自動繳納後，若不服舉發事實者，仍得於繳納罰鍰30日內向處罰機關陳述。
'</div>
%>
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:50px; top:800px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:50px; top:835px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" or trim(Sys_DriverID)<>"" then%>
<div id="Layer3" style="position:absolute; left:170px; top:807px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:825px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<div id="Layer9" style="position:absolute; left:40px; top:855px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		'response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:850px; width:233px; height:32px; z-index:3"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer12" style="position:absolute; left:120px; top:910px; width:250px; height:11px; z-index:6"><span class="style7">逕行舉發　<%=funcCheckFont(Sys_A_Name,18,1)%><br><%if int(Sys_Rule1)<>4340003 and int(Sys_Rule1)<>5620001 and Sys_BillFillerMemberID <> "3480" then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>

<div id="Layer14" style="position:absolute; left:370px; top:910px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "<font color=""red"">*本單可至郵局或委託代收之超商繳納</font>"%></div>

<div id="Layer17" style="position:absolute; left:620px; top:930px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:125px; top:955px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:270px; top:955px; width:250px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:510px; top:955px; width:300px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" style="position:absolute; left:125px; top:980px; width:590px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)%></div>

<div id="Layer22" style="position:absolute; left:120px; top:1000px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:170px; top:1000px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:220px; top:1000px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:265px; top:1000px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:315px; top:1000px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>

<div id="Layer27" style="position:absolute; left:400px; top:1005px; width:610px; height:31px; z-index:20"><span class="style3"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1
		if int(Sys_Rule1)=5620001 then	Sys_IllegalRule1=Sys_IllegalRule1&"(掛號催繳通知補繳逾7日期限)"
		If trim(Sys_Rule4)<>"" Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	end if
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end if
%></span></div>
<div id="Layer28" style="position:absolute; left:115px; top:1025px; width:260px; height:15px; z-index:21"><span class="style3"><%
	Response.Write Sys_ILLEGALADDRESS
	If Sys_BillFillerMemberID = "3480" Then
		Response.Write ""
		Response.Write "("&Stop_IllegalDate(0)&"年"&Stop_IllegalDate(1)&"月"&Stop_IllegalDate(2)&"日"
		Response.Write right("00"&Stop_IllegalDate_h,2)&"時"&right("00"&Stop_IllegalDate_m,2)&"分)"
	end if
%></span></div>
<div id="Layer29" style="position:absolute; left:130px; top:1055px; width:50px; height:11px; z-index:22"><b><font color="red"><%=Sys_DealLineDate(0)%></font></b></div>
<div id="Layer30" style="position:absolute; left:210px; top:1055px; width:35px; height:13px; z-index:23"><b><font color="red"><%=Sys_DealLineDate(1)%></font></b></div>
<div id="Layer31" style="position:absolute; left:290px; top:1055px; width:32px; height:15px; z-index:24"><b><font color="red"><%=Sys_DealLineDate(2)%></font></b></div>

<div id="Layer32" style="position:absolute; left:400px; top:1070px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:395px; top:1100px; width:410px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" style="position:absolute; left:635px; top:1105px; width:100px; height:40px; z-index:28"><span class="style7"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>
<div id="Layer35" style="position:absolute; left:400px; top:1165px; width:130px; height:49px; z-index:29"><%
	if billprintuseimage=1 then
		response.write "<table border=""1"" width=""120"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td width=""100"" class=""style14"">花蓮縣警察局<br>"&sysunit&"</td></tr>"
		response.write "<tr><td width=""100"" class=""style14"">TEL"&Sys_UnitTel&"</td></tr>"
		response.write "</table>"

		'if trim(Sys_UnitFilename)<>"" then	response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""110"" height=""80"">"
		'response.write "<img src=""unit3.jpg"">"
	end if
%></div>
<div id="Layer36" style="position:absolute; left:590px; top:1200px; width:140px; height:43px; z-index:30"><%
	If not ifnull(Sys_JobName) Then

		response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=1>"

		response.write "<tr><td style=""border-color:#ff0000;border-width:2px;border-style:solid;"" width=""120"" align=""center""><span class=""style10"">交通違規專用章</span><br><span class=""style13"">"& Sys_JobName & "&nbsp;" &Sys_MainChName&"</span></td></tr>"

		response.write "</table>"
	End if 
	

%></div>
<div id="Layer37" style="position:absolute; left:590px; top:1160px; width:200px; height:46px; z-index:31"><%
	If Sys_BillFillerMemberID = "3480" Then
		response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=1>"
		response.write "<tr><td style=""border-color:#ff0000;border-width:2px;border-style:solid;"" width=""100"" height=""25"" align=""center"" nowrap><span class=""style13"">&nbsp;組長&nbsp;司俊華</span></td></tr>"
		response.write "</table>"

	elseIf Sys_UnitLevelID=1 Then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		elseif instr(Sys_BillNo,"TG1")>0 then
			response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=1>"
			response.write "<tr><td style=""border-color:#ff0000;border-width:2px;border-style:solid;"" width=""100"" height=""25"" align=""center"" nowrap><span class=""style13"">&nbsp;組長&nbsp;司俊華</span></td></tr>"
			response.write "</table>"
			response.write "<font size=1>"&Sys_Credit_ID&"</font>"
		else
			response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=1>"
			response.write "<tr><td style=""border-color:#ff0000;border-width:2px;border-style:solid;"" width=""100"" height=""25"" align=""center"" nowrap><span class=""style13"">&nbsp;警員&nbsp;黃玉琳</span></td></tr>"
			response.write "</table>"
			response.write "<font size=1>"&Sys_Credit_ID&"</font>"
		end if
	else
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		else
			response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=0 width=""90"">"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=25 align=""center"" nowrap><span class=""style13"">&nbsp;"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"
		end if
	end if
%></div>
<div id="Layer38" style="position:absolute; left:210px; top:1240px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:365px; top:1240px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:515px; top:1240px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:690px; top:1240px; width:80px; height:12px; z-index:36"><%%></div>
<div id="Layer43" style="position:absolute; left:620px; top:1260px; width:250px; height:12px; z-index:36"><B><font size=4><%=Sys_CarNo%></font></B></div>

</div>

<%
		if (i mod 10)=0 then response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	//window.print();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>