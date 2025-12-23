<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 10px;}
.style2 {font-family:"標楷體"; font-size: 12px;}
.style3 {font-family:"標楷體"; font-size: 16px;}
.style4 {font-family:"標楷體"; font-size: 22px;}
.style5 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style6 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style7 {font-family:"標楷體"; font-size: 20px;}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsxie8.cab#Version=6,5,439,50">
</object>
<%
on Error Resume Next
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
for i=0 to Ubound(PBillSN)
if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
strBil="select distinct BillSN,BillNo,CarNo,DCIerrorCarData,BatchNumber from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_DCIerrorCarData=""
if Not rsbil.eof then Sys_DCIerrorCarData=trim(rsbil("DCIerrorCarData"))
if Not rsbil.eof then Sys_BatchNumber=trim(rsbil("BatchNumber"))
strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillUnitID=trim(rs("BillUnitID"))
if Not rs.eof then Sys_RecordMemberID=trim(rs("RecordMemberID"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

DriverStatus=""

strSql="select a.*,b.DriverHomeZip DriverZip,b.DriverHomeAddress DriverAddress from (select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W') a,(select CarNo,DriverHomeZip,DriverHomeAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A') and ExchangetypeID='A') b where a.carno=b.carno(+)"
set rsfound=conn.execute(strSql)

If ifnull(Sys_OwnerAddress) Then

	Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

	if Not rsfound.eof then Sys_Owner=rsfound("Owner")
	chkaddress=""
	If Not ifnull(trim(rsfound("OwnerAddress"))) Then
		If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") Then
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

strSql="select BillNo from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='N'"
set rstwo=conn.execute(strSql)
if not rstwo.eof then
	Sys_OwnerAddressDeliver=trim(rsfound("DriverAddress"))
	Sys_OwnerZipDeliver=trim(rsfound("DriverZip"))

	If ifnull(Sys_OwnerAddressDeliver) Then
		Sys_OwnerAddressDeliver=Sys_OwnerAddress
		Sys_OwnerZipDeliver=Sys_OwnerZip
	End if
else

	Sys_OwnerAddressDeliver=Sys_OwnerAddress
	Sys_OwnerZipDeliver=Sys_OwnerZip

end if

rstwo.close

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerAddress) Then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
end if

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then
		Sys_Owner=trim(replace(Sys_Owner," ","*"))
		errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
	end if
end if

Sys_Owner=trim(replace(Sys_Owner," ","*"))

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZipDeliver&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipNameDeliver=trim(rszip("ZipName"))
rszip.close

rsfound.close
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where SN="&PBillSN(i)
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

chkIllegalDate=false:tmpIllegalDate1="":tmpIllegalDate2=""
if Not rssex.eof then
	Sys_IllegalDate=split(gArrDT(trim(rssex("IllegalDate"))),"-")

	tmpIllegalDate1=formatDateTime(DateAdd("d",-1,rssex("IllegalDate")),vbshortDate)
	tmpIllegalDate2=formatDateTime(DateAdd("d",1,rssex("IllegalDate")),vbshortDate)
	
	strSQL="select NVL(count(*),0) cnt from BillUserTake where CarNo='"&Sys_CarNo&"' and IllegalDate between "&funGetDate(tmpIllegalDate1,0)&" and "&funGetDate(tmpIllegalDate2,0)
	set rscnt=conn.execute(strSQL)
	If cdbl(rscnt("cnt")) > 0 Then chkIllegalDate=true
	rscnt.close

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
if Not rssex.eof then Sys_IMAGEFILENAMEB=trim(rssex("IMAGEFILENAMEB"))
if Not rssex.eof then Sys_IMAGEPATHNAME=trim(rssex("IMAGEPATHNAME"))
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSQL="select UnitID,UnitName,UnitTypeID,UnitLevelID from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
Sys_UnitID=Unit("UnitID")
Sys_RedUnitName=Unit("UnitName")
Sys_UnitTypeID=Unit("UnitTypeID")
Sys_UnitLevelID=Unit("UnitTypeID")
Unit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysAddress=Unit("Address")
SysUnitTel=Unit("Tel")
Unit.close

strSQL="select UnitName,Address,Tel from UnitInfo where UnitID='"&Sys_BillUnitID&"'"
set Unit=conn.execute(strSQL)
SysUnitLevel3=Unit("UnitName")
SysAddressLevel3=Unit("Address")
SysUnitTelLevel3=Unit("Tel")
Unit.close

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
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
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

strSql="select LPAD(MailNumber,6,'0') MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))

rs.close
strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&replace(trim(rsbil("BillSN")),"","0")&" and a.CarNo='"&trim(rsbil("CarNo"))&"'"
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

Sys_DciReturnCarColor=""

StrBass="select DciReturnCarColor from BillBaseDCIReturn where EXCHANGETYPEID='W' and CarNo='"&trim(rsbil("CarNo"))&"' and BillNo='"&trim(rsbil("BillNo"))&"'"
Sys_CarColor="":Sys_CarColorID=""
set rsCarType=conn.execute(strBass)
if not rsCarType.eof then 
	If not ifnull(trim(rsCarType("DciReturnCarColor"))) Then
		For h = 1 to len(trim(rsCarType("DciReturnCarColor")))
			If h>1 Then Sys_DciReturnCarColor=Sys_DciReturnCarColor&","
			Sys_DciReturnCarColor=Sys_DciReturnCarColor&mid(trim(rsCarType("DciReturnCarColor")),h,1)
		Next
		
	End if	
end if
rsCarType.close

Sys_CarColorID=split(Sys_DciReturnCarColor,",")
Sys_CarColor=""
for y=0 to ubound(Sys_CarColorID)
	if trim(Sys_CarColor)<>"" then Sys_CarColor=Sys_CarColor&","
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	If not rscolor.eof Then Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	rscolor.close
next

If ifnull(Sys_MailNumber) Then Sys_MailNumber=0

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,800,263,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,800,263,36

	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,28,160,0

'	response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",1"
'	response.end
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
pagesum=0
if i>0 then pagesumAdd=-5
'pagesum=i*2074+530
'If i=0 Then pagesum=530
'if i>3 then pagesum=pagesum-5
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;">

<div id="Layer50" class="style3" style="position:absolute; left:10px; top:<%=0%>px; width:400px; height:10px; z-index:5"><%=Sys_BatchNumber&"　"&SysUnit&"　"&"("&cdbl(i+1)&"/"&cdbl(Ubound(PBillSN)+1)&")"%></div>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:10px; top:<%=pagesum+20%>px; width:10px; height:20px; z-index:5">v</div>
<%else%>
<div id="Layer2" style="position:absolute; left:10px; top:<%=pagesum+45%>px; width:10px; height:20px; z-index:5">v</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:130px; top:<%=pagesum+35%>px; width:202px; height:36px; z-index:5">v</div>
<%else%>
	<div id="Layer4" style="position:absolute; left:130px; top:<%=pagesum+40%>px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<!--<div id="Layer5" style="position:absolute; left:185px; top:<%=pagesum+45%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:625px; top:<%=pagesum+25%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:625px; top:<%=pagesum+35%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:625px; top:<%=pagesum+45%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:10px; top:<%=pagesum+70%>px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>
<div id="Layer9" class="style7" style="position:absolute; left:250px; top:<%=pagesum+85%>px; width:202px; height:36px; z-index:5"><%=DriverStatus%></div>
<!--<div id="Layer42" style="position:absolute; left:210px; top:<%=pagesum+70%>px; width:202px; height:36px; z-index:5"><%="<font size=1>"&SysUnit&"<br>("&SysUnitTel&")</font>"%></div>-->
<div id="Layer10" style="position:absolute; left:460px; top:<%=pagesum+70%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:485px; top:<%=(i*1550+110)%>px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer12" class="style3" style="position:absolute; left:60px; top:<%=pagesum+130%>px; width:150px; height:11px; z-index:3"><%
	response.write "逕行舉發　"&funcCheckFont(Sys_A_Name,16,1)&"<br>"
	'response.write "逕行舉發　<br>"
	if left(trim(Sys_Rule1),2)<>"562" then response.write "<span class=""style2"">依據採證照片</span>"
	response.write "　"&Sys_CarColor
%></div>
<div id="Layer13" class="style3" style="position:absolute; left:220px; top:<%=pagesum+130%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" class="style2" style="position:absolute; left:330px; top:<%=pagesum+130%>px; width:500px; height:10px; z-index:4"><%if showBarCode then response.write "<font size=2>*本單可至郵局或期限內至統一、全家、ok、萊爾富等超商繳納</font>"%></div>
<div id="Layer15" class="style3" style="position:absolute; left:210px; top:<%=pagesum+160%>px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" class="style3" style="position:absolute; left:365px; top:<%=pagesum+160%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" class="style3" style="position:absolute; left:560px; top:<%=pagesum+160%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" class="style3" style="position:absolute; left:60px; top:<%=pagesum+175%>px; width:100px; height:14px; z-index:11"><B><%=Sys_CarNo%></B></div>
<div id="Layer19" class="style3" style="position:absolute; left:225px; top:<%=pagesum+175%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style3" style="position:absolute; left:395px; top:<%=pagesum+175%>px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" class="style3" style="position:absolute; left:60px; top:<%=pagesum+200%>px; width:800px; height:13px; z-index:14"><%
	Response.Write Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,14,1)&chkaddress
	If chkIllegalDate Then Response.Write "　(車主自取)"
	
%></div>

<div id="Layer22" class="style3" style="position:absolute; left:70px; top:<%=pagesum+225%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style3" style="position:absolute; left:120px; top:<%=pagesum+225%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style3" style="position:absolute; left:170px; top:<%=pagesum+225%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style3" style="position:absolute; left:220px; top:<%=pagesum+225%>px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" class="style3" style="position:absolute; left:270px; top:<%=pagesum+225%>px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style3" style="position:absolute; left:355px; top:<%=pagesum+220%>px; width:250px; height:31px; z-index:20"><%
	response.write "<font size=3>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"
'			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
'				response.write "<br>100以上"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
'				response.write "<br>80以上未滿100"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
'				response.write "<br>60以上未滿80"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
'				response.write "<br>40以上未滿60"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
'				response.write "<br>20以上未滿40"
'			else
'				response.write "<br>未滿20公里"
'			end if
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		response.write Sys_IllegalRule1

		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"
		
	end if
	if trim(Sys_Rule4)<>"" then response.write "("&Sys_Rule4&")"
	response.write "</font>"
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		response.write "<br>"&Sys_IllegalRule2
	end if
%></div>
<div id="Layer28" class="style3" style="position:absolute; left:60px; top:<%=pagesum+250%>px; width:280px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" class="style3" style="position:absolute; left:100px; top:<%=pagesum+290%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style3" style="position:absolute; left:180px; top:<%=pagesum+290%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style3" style="position:absolute; left:260px; top:<%=pagesum+290%>px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style3" style="position:absolute; left:370px; top:<%=pagesum+295%>px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" class="style3" style="position:absolute; left:355px; top:<%=pagesum+335%>px; width:95px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>
<div id="Layer33" style="position:absolute; left:445px; top:<%=pagesum+330%>px; width:400px; height:30px; z-index:28"><%if showBarCode then response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"%></div>

<div id="Layer35" class="style3" style="position:absolute; left:300px; top:<%=pagesum+445%>px; width:200px; height:49px; z-index:29"><%
	If not ifnull(Request("Sys_BillPrintUnitTel")) Then
		Response.Write "<br>TEL："&Request("Sys_BillPrintUnitTel")
	end if
%></div>
<!--<div id="Layer36" style="position:absolute; left:580px; top:<%=pagesum+420%>px; width:100px; height:43px; z-index:30">主管</div>-->
<div id="Layer37" class="style3" style="position:absolute; left:625px; top:<%=pagesum+450%>px; width:200px; height:46px; z-index:31"><%=Sys_ChName%></div>
<div id="Layer38" class="style3" style="position:absolute; left:230px; top:<%=pagesum+480%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" class="style3" style="position:absolute; left:390px; top:<%=pagesum+480%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" class="style3" style="position:absolute; left:540px; top:<%=pagesum+480%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>

<div style="position:absolute; left:20px; top:<%=pagesum+600%>px;">
<table width="645" border="0">
  <tr>
	<th align="left">&nbsp;</th>
    <th align="left" valign="top"><span class="style4"><%=SysUnit&replace(SysUnitLevel3,SysUnit,"")%><br>　<%=SysAddressLevel3%></span></th>
    <td align="left" height="130" valign="top"></td>
  </tr>
  <tr>
	<td align="left" valign="top"></td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th align="left" colspan="2"><span class="style4"><%=chstr(Sys_Driver)%>　台啟</span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left"><span class="style4"><%=funcCheckFont(Sys_Owner,20,1)%>　台啟</span></th>
	<%end if%>
	<td align="left" valign="top"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th colspan="2" align="left" valign="top" nowrap><span class="style4"><%=Sys_DriverHomeZip%><br>
    <%=Sys_DriverZipName&Sys_DriverHomeAddress%></span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left" valign="top" width="400"><span class="style4"><%=Sys_OwnerZip%><br>
    <%=Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,25,1)&chkaddress%></span></th>
	<%end if%>
    <td>&nbsp;</td>
  </tr>
</table>
</div>

<%If instr("BB,BC,BD",left(Sys_BillNo,2))>0 and Sys_MailNumber>0 Then%>

<div id="Layer61" class="style3" style="position:absolute; left:500px; top:<%=pagesum+850%>px; width:200px; height:32px; z-index:6">
大宗郵資已付掛號函件
</div>

<div id="Layer62" class="style3" style="position:absolute; left:540px; top:<%=pagesum+870%>px; width:200px; height:32px; z-index:6">
第<%=Sys_MailNumber%>號
</div>

<div id="Layer63" style="position:absolute; left:500px; top:<%=pagesum+890%>px; width:200px; height:32px; z-index:6">
<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>>
</div>

<div id="Layer64" class="style3" style="position:absolute; left:500px; top:<%=pagesum+940%>px; width:200px; height:32px; z-index:6">
<%=Sys_MAILCHKNUMBER%>
</div>
<%end if%>


<div id="Layer49" style="position:absolute; left:500px; top:<%=pagesum+660%>px; width:200px; height:32px; z-index:6"><%If chkIllegalDate Then Response.Write "<br>　　　　(車主自取)"%><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer49" class="style3" style="position:absolute; left:10px; top:<%=pagesum+1020%>px; height:32px; z-index:6"><%
If not ifnull(Request("Sys_BillPrintUnitTel")) Then
	Response.Write "申訴服務電話："&Request("Sys_BillPrintUnitTel")
	Response.Write "　上班受理時間：週一至週五&nbsp;上午8:00~12:00&nbsp;下午13:30~17:30"
end if
%>
</div>

<div id="Layer43" class="style7" style="position:absolute; left:290px; top:<%=pagesum+1080%>px; width:300px; height:12px; z-index:36"><%=Sys_RedUnitName%></div>
<div id="Layer44" class="style3" style="position:absolute; left:250px; top:<%=pagesum+1110%>px; width:800px; height:12px; z-index:36"><%
	if Sys_BillTypeID="1" then
		response.write "<font size=2>"&chstr(Sys_Driver)&"</font>"
		response.write "<font size=2>　　"&Sys_DriverHomeZip&"&nbsp;&nbsp;"&Sys_DriverZipName&Sys_DriverHomeAddress&"</font>"
	else
		response.write "<font size=2>"&funcCheckFont(Sys_Owner,10,1)&"</font>"
		response.write "<font size=2>　　"&Sys_OwnerZipDeliver&"&nbsp;&nbsp;"&Sys_OwnerZipNameDeliver&funcCheckFont(Sys_OwnerAddressDeliver,10,1)&chkaddress&"</font>"
	end if%></div>
<div id="Layer45" class="style3" style="position:absolute; left:370px; top:<%=pagesum+1130%>px; width:280px; height:12px; z-index:36"><%=Sys_BillNo%></div>

<div id="Layer46" style="position:absolute; left:450px; top:<%=pagesum+1130%>px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>></div>

<%If instr("BB,BC,BD",left(Sys_BillNo,2))>0 and Sys_MailNumber>0 Then%>

<div id="Layer46" style="position:absolute; left:440px; top:<%=pagesum+1550%>px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%> width="140" height="30"></div>

<div id="Layer47" class="style3" style="position:absolute; left:580px; top:<%=pagesum+1555%>px; width:500px; z-index:36">第<%=Sys_MailNumber%>號</div> 
<%end if%>

<div id="Layer47" class="style3" style="position:absolute; left:70px; top:<%=pagesum+1555%>px; width:500px; z-index:36"><%
If instr(SysUnitLevel3,"分隊") > 0 Then
	Response.Write SysAddressLevel3
else
	Response.Write SysAddress
end if
%></div> 
</div>
<%
	if (i mod 50)=0 then response.flush
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
	printWindow(true,0,5.08,0,5.08);
</script>