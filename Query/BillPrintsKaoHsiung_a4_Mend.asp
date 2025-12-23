<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印--A4 size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style2 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.pageprint {
  margin-left: 0mm;
  margin-right: 0mm;
  margin-top: 0mm;
  margin-bottom: 0mm;
}
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
on Error Resume Next
if trim(request("printStyle"))<>"" then

Function onderline(strown)
	'if len(trim(strown))<3 then
		strown=replace(strown,"  ","＿")
	'end if
	onderline=strown
End Function

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Not ifnull(request("billprintuseimage")) Then session("billprintuseimage")=request("billprintuseimage")
Sys_billprintuseimage=session("billprintuseimage")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select BillTypeID,Driver,Rule4,DriverID,DriverAddress,DriverZip,ILLEGALADDRESS,IllegalSpeed,RuleSpeed,INSURANCE,RuleVer,Note,BillFillDate from BillBase where SN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Rule4=""
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
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
'if Not rsfound.eof then Sys_Driver=trim(rsfound("Driver"))
'if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
'if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
'if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
'if Not rsFound.eof then
'	if trim(rsfound("OwnerCounty"))<>"" then
'		Sys_OwnerZipName=trim(rsfound("OwnerCounty"))
'	else
'		Sys_DriverHomeZip=trim(rsfound("DriverCounty"))
'	end if
'end if
strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

If Not ifnull(trim(rsfound("Driver"))) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
else
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if
strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If Sys_BillTypeID=2 Then
	If instr(Sys_Owner," ")>0 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
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
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
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
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFillerTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

if trim(Sys_UnitLevelID)="3" then
	chkJobID="303,314"
else
	chkJobID="303"
end if

strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"
'response.write strSQL
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

If not ifnull(request("hd_BillJobName")) Then
	Sys_jobName=request("hd_BillJobName")
	Sys_MainChName=request("hd_MainChName")
End if

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
elseif Sys_UnitLevelID=2 and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	sysDBunit=unit("UnitName")
	Sys_UnitTel=trim(unit("Tel"))
	Sys_UnitAddress=trim(unit("Address"))
end if
unit.close

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

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close
strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
Sys_MailNumber=0
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
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	rscolor.close
next
if ifnull(Sys_MailNumber) then Sys_MailNumber=0
Sys_MailDate=0
if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
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

showBarCode=true

if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
	if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
		if Sys_IllegalSpeed-Sys_RuleSpeed>=60 then showBarCode=false
	end if
elseif left(trim(Sys_Rule1),5)="12101" or left(trim(Sys_Rule1),5)="12102" or left(trim(Sys_Rule1),5)="12105" or left(trim(Sys_Rule1),5)="12106" or left(trim(Sys_Rule1),5)="12107" or left(trim(Sys_Rule1),5)="12109" or left(trim(Sys_Rule1),5)="12110" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="16105" or left(trim(Sys_Rule1),5)="25103" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="25103" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="272" or left(trim(Sys_Rule1),3)="294" or left(trim(Sys_Rule1),3)="295" or left(trim(Sys_Rule1),3)="303" or left(trim(Sys_Rule1),3)="314" or left(trim(Sys_Rule1),3)="362" or left(trim(Sys_Rule1),3)="363" or left(trim(Sys_Rule1),3)="621" or left(trim(Sys_Rule1),3)="624" or left(trim(Sys_Rule1),3)="625" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="13" or left(trim(Sys_Rule1),2)="18" or left(trim(Sys_Rule1),2)="20" or left(trim(Sys_Rule1),2)="21" or left(trim(Sys_Rule1),2)="23" or left(trim(Sys_Rule1),2)="24" or left(trim(Sys_Rule1),2)="26" or left(trim(Sys_Rule1),2)="35" or left(trim(Sys_Rule1),2)="37" or left(trim(Sys_Rule1),2)="43" or left(trim(Sys_Rule1),2)="54" or left(trim(Sys_Rule1),2)="61" then
	showBarCode=false
elseif int(Sys_Rule1)=29300012 or int(Sys_Rule1)=29300022 or int(Sys_Rule1)=3400003 or int(Sys_Rule1)=3400004 then
	showBarCode=false
end if
If showBarCode=true Then
	if left(trim(Sys_Rule2),2)="40" or (int(Sys_Rule2)>3310101 and int(Sys_Rule2)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			if Sys_IllegalSpeed-Sys_RuleSpeed>=60 then showBarCode=false
		end if
	elseif left(trim(Sys_Rule2),5)="12101" or left(trim(Sys_Rule2),5)="12102" or left(trim(Sys_Rule2),5)="12105" or left(trim(Sys_Rule2),5)="12106" or left(trim(Sys_Rule2),5)="12107" or left(trim(Sys_Rule2),5)="12109" or left(trim(Sys_Rule2),5)="12110" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="16105" or left(trim(Sys_Rule2),5)="25103" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="25103" then
		showBarCode=false
	elseif left(trim(Sys_Rule1),3)="272" or left(trim(Sys_Rule2),3)="294" or left(trim(Sys_Rule2),3)="295" or left(trim(Sys_Rule2),3)="303" or left(trim(Sys_Rule2),3)="314" or left(trim(Sys_Rule2),3)="362" or left(trim(Sys_Rule2),3)="363" or left(trim(Sys_Rule2),3)="621" or left(trim(Sys_Rule2),3)="624" or left(trim(Sys_Rule2),3)="625" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="13" or left(trim(Sys_Rule2),2)="18" or left(trim(Sys_Rule2),2)="20" or left(trim(Sys_Rule2),2)="21" or left(trim(Sys_Rule2),2)="23" or left(trim(Sys_Rule2),2)="24" or left(trim(Sys_Rule2),2)="26" or left(trim(Sys_Rule2),2)="35" or left(trim(Sys_Rule2),2)="37" or left(trim(Sys_Rule2),2)="43" or left(trim(Sys_Rule2),2)="54" or left(trim(Sys_Rule2),2)="61" then
		showBarCode=false
	elseif int(Sys_Rule2)=29300012 or int(Sys_Rule2)=29300022 or int(Sys_Rule2)=3400003 or int(Sys_Rule2)=3400004 then
		showBarCode=false
	end if
End if

%><div id="L78" style="position:relative;"><%
if showBarCode then
%>
	<div id="Layer1" style="position:absolute; font-size:16px; left:70px; top:0px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer2" style="position:absolute; font-size:16px; left:70px; top:20px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; font-size:16px; left:190px; top:5px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer4" style="position:absolute; font-size:16px; left:190px; top:15px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; font-size:16px; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; font-size:16px; left:675px; top:10px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; font-size:16px; left:675px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; font-size:16px; left:675px; top:40px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; font-size:16px; left:35px; top:50px; width:233px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; font-size:16px; left:540px; top:50px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" style="position:absolute; font-size:16px; left:515px; top:85px; width:300px; height:12px; z-index:7"><font size=1>　<font size=3><%=BillPageUnit%>交字第<%="　　　　　"%>號</font></div>
<div id="Layer12" style="position:absolute; font-size:16px; left:125px; top:110px; width:250px; height:30px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer13" style="position:absolute; font-size:16px; left:280px; top:110px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; font-size:16px; left:390px; top:110px; width:360px; height:10px; z-index:4"><font size=2><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納，惟手續費需另行支付"%></font></div>
<div id="Layer15" style="position:absolute; font-size:16px; left:285px; top:130px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; font-size:16px; left:455px; top:130px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; font-size:16px; left:650px; top:130px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; font-size:16px; left:125px; top:155px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; font-size:16px; left:285px; top:155px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; font-size:16px; left:470px; top:155px; width:251px; height:17px; z-index:13"><%=onderline(funcCheckFont(Sys_Owner,22,1))%></div>
<div id="Layer21" style="position:absolute; font-size:16px; left:125px; top:180px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&"　"&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; font-size:16px; left:125px; top:205px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer23" style="position:absolute; font-size:16px; left:175px; top:205px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer24" style="position:absolute; font-size:16px; left:215px; top:205px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer25" style="position:absolute; font-size:16px; left:255px; top:205px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer26" style="position:absolute; font-size:16px; left:300px; top:205px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>
<div id="Layer27" style="position:absolute; font-size:15px; left:425px; top:205px; width:620px; height:31px; z-index:20"><%
	'response.write "<font size=2>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "行車限速"&Sys_RuleSpeed&"公里、經測時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
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
		If trim(Sys_Rule4)<>"" Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
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
				
	'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
	'response.write "</font>"
%></div>
<div id="Layer28" style="position:absolute; font-size:16px; left:125px; top:235px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; font-size:16px; left:125px; top:265px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer30" style="position:absolute; font-size:16px; left:175px; top:265px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer31" style="position:absolute; font-size:16px; left:215px; top:265px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer32" style="position:absolute; font-size:16px; left:425px; top:265px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>道路交通管理處罰條例<br>第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; font-size:16px; left:430px; top:315px; width:85px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>

<div id="Layer33" style="position:absolute; font-size:16px; left:510px; top:310px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>

<div id="Layer35" style="position:absolute; font-size:16px; left:425px; top:370px; width:150px; height:49px; z-index:29"><%
	if instr(SysUnit,"仁武")=0 and instr(SysUnit,"林園")=0 then
		response.write "<font size=3>高雄縣政府警察局"&Sys_UnitName&"<br>"&Sys_UnitFillerTel&"</font>"
	end if
%></div>
<div id="Layer37" style="position:absolute; font-size:16px; left:625px; top:375px; width:200px; height:46px; z-index:31"><%
	If trim(Sys_billprintuseimage)="1" Then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		elseif instr(SysUnit,"分局")=0 then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"
		end if
	end if
%></div>
<div id="Layer36" style="position:absolute; font-size:16px; left:625px; top:410px; width:160px; height:43px; z-index:30"><%
	If trim(Sys_billprintuseimage)="1" Then
		if instr(SysUnit,"分局")=0 then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">舉發單專用章<br>"&Sys_JobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
			response.write "</table>"
		end if
	end if
%></div>
<div id="Layer38" style="position:absolute; font-size:16px; left:125px; top:465px; width:250px; height:10px; z-index:32">中　　華　　民　　國　　<%=sys_Date(0)%>　年　</div>
<div id="Layer39" style="position:absolute; font-size:16px; left:390px; top:465px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>　月　</div>
<div id="Layer40" style="position:absolute; font-size:16px; left:470px; top:465px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>　日</div>
<div id="Layer41" style="position:absolute; font-size:16px; left:580px; top:465px; width:120px; height:12px; z-index:36">填單</div>

<%if showBarCode then%>
	<div id="Layer42" style="position:absolute; font-size:16px; left:70px; top:530px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer43" style="position:absolute; font-size:16px; left:70px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>

<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; font-size:16px; left:190px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; font-size:16px; left:190px; top:545px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>

<!--<div id="Layer46" style="position:absolute; font-size:16px; left:180px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; font-size:16px; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; font-size:16px; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; font-size:16px; left:670px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; font-size:16px; left:35px; top:580px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; font-size:16px; left:540px; top:575px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer52" style="position:absolute; font-size:16px; left:515px; top:610px; width:250px; height:12px; z-index:7"><font size=1>　<font size=3><%=BillPageUnit%>交字第<%="　　　　　"%>號</font></div>
<div id="Layer53" style="position:absolute; font-size:16px; left:125px; top:640px; width:250px; height:30px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>

<div id="Layer54" style="position:absolute; font-size:16px; left:285px; top:635px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; font-size:16px; left:390px; top:635px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; font-size:16px; left:285px; top:660px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; font-size:16px; left:455px; top:660px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; font-size:16px; left:650px; top:660px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; font-size:16px; left:125px; top:685px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; font-size:16px; left:285px; top:685px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>

<div id="Layer61" style="position:absolute; font-size:20px; left:470px; top:685px; width:251px; height:17px; z-index:13"><%=onderline(funcCheckFont(Sys_Owner,22,1))%></div>

<div id="Layer62" style="position:absolute; font-size:20px; left:125px; top:705px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&"　"&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" style="position:absolute; font-size:16px; left:125px; top:730px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer64" style="position:absolute; font-size:16px; left:175px; top:730px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer65" style="position:absolute; font-size:16px; left:215px; top:730px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer66" style="position:absolute; font-size:16px; left:255px; top:730px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer67" style="position:absolute; font-size:16px; left:295px; top:730px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>

<div id="Layer68" style="position:absolute; font-size:15px; left:425px; top:730px; width:620px; height:31px; z-index:20"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "行車限速"&Sys_RuleSpeed&"公里、經測時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
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
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
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
'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>
<div id="Layer69" style="position:absolute; font-size:16px; left:125px; top:765px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; font-size:16px; left:125px; top:790px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer71" style="position:absolute; font-size:16px; left:175px; top:790px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer72" style="position:absolute; font-size:16px; left:215px; top:790px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>

<div id="Layer73" style="position:absolute; font-size:16px; left:425px; top:790px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>道路交通管理處罰條例<br>第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>

<div id="Layer75" style="position:absolute; font-size:16px; left:430px; top:840px; width:280px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME& "<br>" &Sys_StationTel%></font></div>

<div id="Layer74" style="position:absolute; font-size:16px; left:510px; top:835px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>

<div id="Layer76" style="position:absolute; font-size:16px; left:430px; top:905px; width:150px; height:49px; z-index:29"><%="<font size=3>高雄縣警察局交通隊<br>07-7452001</font>"%></div>
<div id="Layer77" style="position:absolute; font-size:16px; left:625px; top:910px; width:200px; height:46px; z-index:31"><%
	If trim(Sys_billprintuseimage)="1" Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
		response.write "</table>"
	end if
%></div>
<div id="Layer82" style="position:absolute; font-size:16px; left:625px; top:945px; width:160px; height:43px; z-index:30"><%
	If trim(Sys_billprintuseimage)="1" Then
		if instr(SysUnit,"分局")=0 then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">舉發單專用章<br>"&Sys_JobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
			response.write "</table>"
		end if
	end if
%></div>
<div id="Layer78" style="position:absolute; font-size:16px; left:125px; top:990px; width:250px; height:10px; z-index:32">中　　華　　民　　國　　<%=sys_Date(0)%>　年</div>
<div id="Layer79" style="position:absolute; font-size:16px; left:390px; top:990px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>　月　</div>
<div id="Layer80" style="position:absolute; font-size:16px; left:470px; top:990px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>　日</div>
<div id="Layer81" style="position:absolute; font-size:16px; left:565px; top:990px; width:120px; height:12px; z-index:36">填單</div>

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
	printWindow(true,5.08,5.08,5.08,5.08);
</script>