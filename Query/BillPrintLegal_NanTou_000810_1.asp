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
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 10px}
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style13 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style14 {font-family:"標楷體"; font-size: 16px;}
.style15 {font-family:"標楷體"; font-size: 8px; color:#ff0000; }
-->
</style>
</head>

<body>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
'on Error Resume Next
if trim(request("printStyle"))<>"" then
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000

Sys_RecLoginID=""
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

for i=0 to Ubound(PBillSN)
'for j=0 to 1
'if (i+j)>Ubound(PBillSN) then exit for
if int(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)

Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""

Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_RecordDate=trim(rs("RecordDate"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then
	Sys_Jurgeday=split(gArrDT(trim(rs("Jurgeday"))),"-")
else
	Sys_Jurgeday=split(gArrDT(trim("")),"-")
end if

rs.close

If ifnull(Sys_OwnerAddress) Then
	strSQL="select OwnerNotIfyAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
	
	set rsfi=conn.execute(strSql)

	if Not rsfi.eof then
		If Not ifnull(trim(rsfi("OwnerNotIfyAddress"))) Then

			notify_Zip="":notify_Addr=""

			if IsNumeric(left(trim(rsfi("OwnerNotIfyAddress")),3)) then

				notify_Zip=left(trim(rsfi("OwnerNotIfyAddress")),3)
			end If 

			notify_Addr=replace(trim(rsfi("OwnerNotIfyAddress")),notify_Zip,"")

			If instr(replace(trim(rsfi("OwnerNotIfyAddress")),"（","("),"(")<=0 then

				notify_Addr=notify_Addr&"(通)"
			end If 
			
			strSQL="update billbasedcireturn set OwnerZip='"&notify_Zip&"',OwnerAddress='"&notify_Addr&"' where exchangetypeid='W' and BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"

			conn.execute(strSQL)
		end if
	end If 
	rsfi.close
end if

strSql="select a.*,b.DriverHomeZip DriverZip,b.DriverHomeAddress DriverAddress from (select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W') a,(select CarNo,DriverHomeZip,DriverHomeAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A') b where a.carno=b.carno(+)"
set rsfound=conn.execute(strSql)

If ifnull(Sys_OwnerAddress) Then

	if Not rsfound.eof then Sys_Owner=rsfound("Owner")

	chkaddress=""
	If Not ifnull(trim(rsfound("OwnerAddress"))) Then
		If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") or instr(replace(rsfound("OwnerAddress"),"（","("),"(通") Then
			if Not rsfound.eof then Sys_OwnerAddress=rsfound("OwnerAddress")
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		End if

		If ifnull(Sys_OwnerAddress) Then
			chkaddress="(戶)"
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverAddress"))
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverZip"))
		End if

	else
		If ifnull(Sys_OwnerAddress) Then
			chkaddress="(戶)"
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverAddress"))
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverZip"))
		End if
	end if

	If ifnull(Sys_OwnerAddress) Then
		chkaddress=""
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if

	If not ifnull(Sys_OwnerAddress) Then
		strSQL="Update Billbase set Owner='"&rsfound("Owner")&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
		conn.execute(strSQL)
	end if
end if

if not ifnull(Sys_OwnerAddress) then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"  "," ")," ","　")
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerAddress) Then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,"臺","台")
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
End if

If not ifnull(Sys_OwnerAddress) Then
	strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
	conn.execute(strSQL)
end if

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
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID from BillBase where SN="&trim(rsbil("BillSN"))
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

if ifnull(Sys_RecLoginID) then
	strSQL="select loginid from memberdata where memberid="&Sys_RecordMemberID
	set recmem=conn.execute(strSQL)
		Sys_RecLoginID=recmem("LoginID")
	recmem.close
end if

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
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

Sys_ImageFileNameA="":Sys_ImageFileNameB=""

strSQL="select * from BillIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	Sys_ImageFileNameA=trim(rsimage("ImageFileNameA"))
	Sys_ImageFileNameB=trim(rsimage("ImageFileNameB"))
end if

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.Tel,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then title_unit=unit("UnitName")
if Not unit.eof then title_UnitAddress=trim(unit("Address"))
if Not unit.eof then title_UnitTel=trim(unit("Tel"))
unit.close

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

chkJobID="":strUit=""
if trim(Sys_UnitLevelID)="3" then
	chkJobID="314,316"
	strUit="UnitID='"&Sys_UnitID&"'"
elseif trim(Sys_UnitLevelID)="2" then
	chkJobID="304,305"
	strUit="UnitID in(select UnitID from UnitInfo where UnitTypeid='"&Sys_UnitID&"' and UnitLevelID=2)"
elseif trim(Sys_UnitLevelID)="1" then
	chkJobID="303"
	strUit="UnitID='"&Sys_UnitID&"'"
end if
		
	strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and "&strUit&" and JobID in ("&chkJobID&")) a,(select ID,Content from Code where TypeID=4) b where a.JobID=b.ID order by ID"
	set rsjob=conn.execute(strSQL)

	if Not rsjob.eof then
		Sys_jobName=rsjob("Content")
		Sys_MainChName=rsjob("ChName")
	end if
	rsjob.close

strSQL="select UnitName,Tel from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
Unit.close

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

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=right("000000"&trim(rs("MailNumber")),14)
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
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	If not rscolor.eof Then	Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	rscolor.close
next

strSQL="Update BillMailHistory set FirstBarCode='"&Sys_Rule1&"-"&Sys_BillNo&"-"&Sys_CarNo&"',MailChkNumber='"&Sys_MailNumber&"' where BillSN="&PBillSN(i)

conn.execute(strSQL)

If ifnull(Sys_MailNumber) Then Sys_MailNumber=0

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
end if
strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close

pageTop=0
pageLeft=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer41" class="style3" style="position:absolute; left:105px; top:5px; width:200px; height:12px; z-index:36"><font color="red"><%=title_Unit%></font></div>
<div id="Layer41" class="style2" style="position:absolute; left:355px; top:5px; width:200px; height:12px; z-index:36"><%=title_UnitAddress%></div>
<div id="Layer41" style="position:absolute; left:485px; top:0px; width:200px; height:12px; z-index:36"><%=Sys_RecLoginID&"&nbsp;"&sys_Date(0)&sys_Date(1)&sys_Date(2)%></div>

<div id="Layer05" class="style3" style="position:absolute; left:570px; top:0px; z-index:5">
	<%If Sys_MailNumber<>0 Then Response.Write "　<b>第"&Sys_MailNumber&"號</b>"%>
</div>

<div id="Layer01" class="style3" style="position:absolute; left:125px; top:20px; z-index:3"><B><%
	'response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write funcCheckFont(Sys_Owner,16,1)&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
%></B>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:125px; top:60px; z-index:2"><%
	if trim(Sys_CarNo)<>"" and not isnull(Sys_CarNo) then
		response.write left(Sys_CarNo,4)
		response.write left("*************",len(Sys_CarNo)-4)
	end if 
	%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:145px; top:50px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:390px; top:275px; z-index:1"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:125px; top:290px; z-index:1"><b><%
	if trim(Sys_BillTypeID)="1" then
		response.write Sys_Driver
	elseif trim(Sys_BillTypeID)="2" then
		response.write funcCheckFont(Sys_Owner,16,1)
	end if%>　台啟</b>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:105px; top:310px; width:330px; z-index:1"><b><%
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

<!---------------------------------- 寄件人------------------------------------>
<div id="Layer72" class="style3" style="position:absolute; left:123px; top:328px; width:300px; height:36px; z-index:1"><%
	Response.Write thenPasserCity&title_Unit&"<br>"
		
	
%></div>

<div id="Layer73" class="style3" style="position:absolute; left:123px; top:348px; width:300px; height:36px; z-index:5"><%
	
		Response.Write title_UnitAddress
		Response.Write "<br>"
		Response.Write title_UnitTel
	
%></div>


<div id="Layer05" class="style3" style="position:absolute; left:470px; top:287px; z-index:5">
	　<%
		If Sys_MailNumber<>0 Then
			Response.Write "<b>第"&Sys_MailNumber&"號</b><br>"
			Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg"">"
				Response.Write "<br>"
		End if
		
		response.write "　　"&left(Sys_CarNo,4)
		response.write left("*************",len(Sys_CarNo)-4)
	

		response.write "　"&Sys_A_Name&"　"&Sys_CarColor
	%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:123px; top:390px; z-index:1"><%
	
	Response.Write "　"&Sys_STATIONNAME
%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:270px; top:345px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:360px; top:370px; z-index:1"><%
	'Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>
<!-------------------------- 判斷 billillegalimage 有沒有這些圖檔 ------------------------------>
<!---------- todo 依據法條判斷, 如果是闖紅燈, 要把 a 檔的 xxxxx_a.jpg 換成 b檔的xxxxxx_b.jpg ---------------------------------------->
<%if trim(Sys_ImageFileNameA)<>"" then%>
	<div id="Layer09" style="position:absolute; left:40px; top:485px; z-index:5"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	%></DIV>
<%
elseIf trim(Sys_Rule1)="5620001" and trim(Sys_ImageFileName)<>"" Then%>
	<div id="Layer09" style="position:absolute; left:40px; top:485px; z-index:5"><%
		response.write "<img src=""../StopCarPicture/"&Sys_ImageFileName&""" width=""390"" height=""280"">"
	%></DIV>
<%End if
'response.write Sys_Rule1 & "_" 
'response.write Sys_ImageFileNameB
' ssmith 20091015 紅燈月線不顯示B圖 陳淑雲於20150305取消這需求 and (Sys_Rule1 <>"6020303" and Sys_Rule2<>"6020303")
%>
<%if trim(Sys_ImageFileNameB)<>"" then%>
	<div id="Layer10" style="position:absolute; left:430px; top:485px; z-index:1"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:50px; top:800px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:50px; top:835px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:170px; top:800px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:820px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<div id="Layer9" style="position:absolute; left:40px; top:870px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:860px; width:233px; height:32px; z-index:3"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer12" style="position:absolute; left:120px; top:925px; width:250px; height:11px; z-index:6"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>

<div id="Layer13" style="position:absolute; left:260px; top:920px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:395px; top:920px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "<font color=""red"">*本單可至郵局或委託代收之超商繳納</font>"%></div>

<div id="Layer15" style="position:absolute; left:260px; top:930px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:425px; top:930px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:620px; top:930px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:125px; top:975px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:325px; top:975px; width:250px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:580px; top:975px; width:300px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" style="position:absolute; left:190px; top:1003px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)%></div>
<!---------------------------------- 違規時間------------------------------------>
<div id="Layer22" style="position:absolute; left:120px; top:1032px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:170px; top:1032px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:220px; top:1032px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:270px; top:1032px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:320px; top:1032px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:440px; top:1040px; width:610px; height:31px; z-index:20"><span class="style3"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里，經檢定合格儀器測照，時速"&Sys_IllegalSpeed&"公里<br>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "(100以上)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "(80以上未滿100)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "(60以上未滿80)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "(40以上未滿60)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "(20以上未滿40)。"
			else
				response.write "(未滿20公里)。"
			end if
		end If 
		if left(trim(Sys_Rule1),2)="40" then Response.Write "<br>(前方設有測速照相告示牌)"

		if trim(Sys_Rule1)="3310102" then Response.Write "<br>(前方設有測速照相告示牌)"
		if trim(Sys_Rule1)="3310104" then Response.Write "<br>(前方設有測速照相告示牌)"
		if trim(Sys_Rule1)="3310106" then Response.Write "<br>(前方設有測速照相告示牌)"

		if trim(Sys_Rule1)="4310201" then Response.Write "<br>(前方設有測速照相告示牌)"
		if trim(Sys_Rule1)="4310202" then Response.Write "<br>(前方設有測速照相告示牌)"
		if trim(Sys_Rule1)="4310203" then Response.Write "<br>(前方設有測速照相告示牌)"
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
			'response.write "<BR>"&Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"	
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
	If trim(Sys_UnitID)="046A" Then response.write " (經科學儀器採證)"	
%></span></div>
<div id="Layer28" style="position:absolute; left:115px; top:1056px; width:220px; height:15px; z-index:21"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:120px; top:1080px; width:50px; height:11px; z-index:22"><b><font color="red"><%=Sys_DealLineDate(0)%></font></b></div>
<div id="Layer30" style="position:absolute; left:200px; top:1080px; width:35px; height:13px; z-index:23"><b><font color="red"><%=Sys_DealLineDate(1)%></font></b></div>
<div id="Layer31" style="position:absolute; left:280px; top:1080px; width:32px; height:15px; z-index:24"><b><font color="red"><%=Sys_DealLineDate(2)%></font></b></div>
<div id="Layer32" class="style3" style="position:absolute; left:537px; top:1123px; width:400px; height:49px; z-index:29"><%
'//for 陳小姐 IP 
'if Request.ServerVariables("REMOTE_ADDR")="10.120.28.38" then
'	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　"
'	if len(trim(Sys_Rule1))>7 then response.write "　　　"&right(trim(Sys_Rule1),1)
'	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)	
'	response.write "　　　　　　　　"&Sys_Level1
'
'else
	response.write left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　罰鍰"&Sys_Level1 &"元"
'end if
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　罰鍰"&Sys_Level2&"元"
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:428px; top:1160px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" style="position:absolute; left:670px; top:1162px; width:70px; height:40px; z-index:28"><span class="style2"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>
<div id="Layer35" style="position:absolute; left:450px; top:1210px; width:130px; height:49px; z-index:29"><%
	response.write "<span class=""style10"">南投縣政府警察局"&Sys_UnitName&"</font></span><br><span class=""style10"">"&Sys_UnitTEL&"</span>"
%></div>
<div id="Layer36" style="position:absolute; left:590px; top:1205px; width:140px; height:43px; z-index:30"><%
	if trim(Sys_UnitLevelID)="3" then Sys_JobName="所長"
	
	response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=1>"

'smith 20110628 暫時修改 由副隊長 代 right='0' 

	'response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" align=""center""><span class=""style10"">違規舉發專用章</span><br><span class=""style13"">"& Sys_JobName & "&nbsp;" &Sys_MainChName&"</span></td></tr>"

	response.write "</table>"
%></div>
<div id="Layer37" style="position:absolute; left:630px; top:1215px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
	else
		response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=0 width=""90"">"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=25 align=""center"" nowrap><span class=""style13"">&nbsp;"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
		response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
	end if%></div>

<div id="Layer38" style="position:absolute; left:120px; top:1258px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:188px; top:1258px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:251px; top:1258px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:690px; top:1278px; width:80px; height:12px; z-index:36"><%=Sys_RecLoginID%></div>
<div id="Layer43" style="position:absolute; left:300px; top:1278px; width:250px; height:12px; z-index:36"><%=Sys_DCIRETURNCARTYPE%></div>
</div>

</div>

<%
		if (i mod 100)=0 then response.flush
next
end if
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
                         //left4.5  right 7.08 up 6.8. down 8.08
	printWindow(true,5.08,4.1,5.08,5.08);
</script>