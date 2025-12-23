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
.style1 {font-family:"標楷體"; font-size: 10px;}
.style2 {font-family:"標楷體"; font-size: 20px; line-height:1;}
.style3 {font-family:"標楷體"; font-size: 10px; color:#ff0000; }
.style4 {font-family:"標楷體"; font-size: 12px;}
.style5 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,5,439,50">
</object>
<%
on Error Resume Next
if trim(request("printStyle"))<>"" then
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
logBillNo=""
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

If not ifnull(logBillNo) Then logBillNo=logBillNo&","
logBillNo=logBillNo&trim(rsbil("BillNo"))

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)

Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_CarSimpleID="":Sys_CarAddID="":Sys_ProjectID="":Sys_Jurgeday=""
Sys_Driver="":Sys_DriverID="":Sys_DriverHomeAddress="":Sys_DriverHomeZip=""
Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=replace(trim(rs("ILLEGALADDRESS")),"台南市","")
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
if Not rs.eof then Sys_CarSimpleID=trim(rs("CarSimpleID"))
if Not rs.eof then Sys_CarAddID=trim(rs("CarAddID"))

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

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

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

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then 
	Sys_OwnerZipName=trim(rszip("ZipName"))
	Sys_OwnerAddress=replace(Sys_OwnerAddress,"臺","台")
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
end if
rszip.close

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

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then chk_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFillerTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
'else
'	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'end if
set Unit=conn.Execute(strSQL)
If Not Unit.eof Then
	SysUnit=unit("UnitName")
	SysUnitTel=trim(unit("Tel"))
end if
Unit.close

if trim(Sys_UnitLevelID)="3" then
	chkJobID="303,304,317"
else
	chkJobID="303,304"
end if

strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"
'response.write strSQL
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
Sys_UnitID=Unit("UnitID")
Sys_UnitTypeID=Unit("UnitTypeID")
Sys_UnitLevelID=Unit("UnitLevelID")
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

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if
strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;"><%
if showBarCode then
%>
<div id="Layer1" style="position:absolute; left:80px; top:0px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:80px; top:20px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:200px; top:0px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:200px; top:15px; width:202px; height:36px; z-index:5">v</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:675px; top:10px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:675px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:675px; top:40px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; left:55px; top:50px; width:233px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:520px; top:45px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" class="style1" style="position:absolute; left:530px; top:85px; width:230px; height:12px; z-index:7">　<%=BillPageUnit%>交字第　　　　　　　　　號</div>
<div id="Layer12" class="style4" style="position:absolute; left:135px; top:110px; width:150px; height:11px; z-index:3"><%
	If not ifnull(Sys_Jurgeday(0)) Then
		Response.Write "民眾檢舉舉發"&Sys_A_Name&"<br>"
		Response.Write "附檢舉照片　"&Sys_CarColor
	else
		Response.Write "逕行舉發　"&Sys_A_Name&"<br>"
		Response.Write "附採證照片　"&Sys_CarColor
	End if 
%></div>
<div id="Layer13" style="position:absolute; left:290px; top:110px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" class="style4" style="position:absolute; left:400px; top:106px; width:370px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納(7-11,全家,萊爾富,OK)"%></div>
<div id="Layer15" style="position:absolute; left:285px; top:130px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; left:465px; top:130px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:660px; top:130px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:135px; top:155px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:300px; top:155px; width:140px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:470px; top:155px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" style="position:absolute; left:140px; top:180px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; left:135px; top:205px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer23" style="position:absolute; left:175px; top:205px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer24" style="position:absolute; left:215px; top:205px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer25" style="position:absolute; left:255px; top:205px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer26" style="position:absolute; left:295px; top:205px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>
<div id="Layer27" class="style4" style="position:absolute; left:435px; top:200px; width:620px; height:31px; z-index:20"><%
	If not ifnull(Sys_Jurgeday(0)) Then
		Response.Write "民眾檢舉日期 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"
	end If 

	if int(Sys_Rule1)=5610102 then Sys_IllegalRule1="在禁止臨時停車處所停車"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>4310209 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "<br>100公里以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "<br>80公里以上未滿100公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "<br>60公里以上未滿80公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "<br>40公里以上未滿60公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "<br>20公里以上未滿40公里"
			else
				response.write "<br>未滿20公里"
			end if
			
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"

		Sys_IllegalRule1=LawRule(Sys_Rule1,Sys_IllegalRule1,Sys_CarSimpleID,Sys_CarAddID)

		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if	
	end if

	if not ifnull(Sys_Rule4) then response.write "("&Sys_Rule4&")"

	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"

		Sys_IllegalRule2=LawRule(Sys_Rule2,Sys_IllegalRule2,Sys_CarSimpleID,Sys_CarAddID)

		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end If 

	If ifnull(Sys_Jurgeday(0)) Then response.write " (經科學儀器採證)"
				
	'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>
<div id="Layer28" style="position:absolute; left:135px; top:235px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:135px; top:265px; width:267px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer30" style="position:absolute; left:175px; top:265px; width:267px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer31" style="position:absolute; left:215px; top:265px; width:267px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>

<div id="Layer32" class="style4" style="position:absolute; left:435px; top:265px; width:400px; height:49px; z-index:29"><%
	Response.Write "　　　　　道 路 交 通 管 理 處 罰 條 例<br>"
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	'response.write ""
	'response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "第"&Mid(trim(Sys_Rule1),3,1)&"項第"&Mid(trim(Sys_Rule1),4,2)&"款規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		'response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
%></div>

<div id="Layer34" class="style4" style="position:absolute; left:434px; top:320px; width:90px; height:30px; z-index:28"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer33" style="position:absolute; left:516px; top:315px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>

<div id="Layer35" class="style3" style="position:absolute; left:435px; top:400px; width:110px; height:49px; z-index:29"><%
	If chk_UnitTypeID = "7100" or chk_UnitTypeID = "7200" or chk_UnitTypeID = "7300" or chk_UnitTypeID = "J01" Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" nowrap>臺南市政府警察局<br>"&SysUnit&"</td></tr>"

		If chk_UnitTypeID <> "7100" then
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"">TEL"&SysUnitTel&"</td></tr>"
		end If 

		response.write "</table>"
	End if 
%></div>

<div id="Layer36" style="position:absolute; left:620px; top:370px; width:160px; height:43px; z-index:30"><%
'	If chk_UnitTypeID = "7400" Then
'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'		response.write "</table>"
'	End if 
%></div>

<div id="Layer37" class="style5" style="position:absolute; left:620px; top:420px; width:200px; height:46px; z-index:31"><%
	If chk_UnitTypeID = "7100" or chk_UnitTypeID = "7200" or chk_UnitTypeID = "7300" or chk_UnitTypeID = "J01" Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</td></tr>"
		response.write "</table>"
	End if 
%></div>

<div id="Layer82" class="style3" style="position:absolute; left:200px; top:450px; width:200px; height:10px; z-index:32">(自103年3月31日起，前、後段日數均改為30日)</div>

<div id="Layer38" style="position:absolute; left:175px; top:465px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer39" style="position:absolute; left:330px; top:465px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer40" style="position:absolute; left:480px; top:465px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer41" style="position:absolute; left:580px; top:465px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>


<%if showBarCode then%>
<div id="Layer42" style="position:absolute; left:80px; top:525px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer43" style="position:absolute; left:80px; top:545px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:200px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:200px; top:545px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:55px; top:580px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:520px; top:575px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer52" class="style1" style="position:absolute; left:530px; top:615px; width:230px; height:12px; z-index:7">　<%=BillPageUnit%>交字第　　　　　　　　　號</div>
<div id="Layer53" class="style4" style="position:absolute; left:135px; top:640px; width:150px; height:11px; z-index:3"><%
	If not ifnull(Sys_Jurgeday(0)) Then
		Response.Write "民眾檢舉舉發"&Sys_A_Name&"<br>"
		Response.Write "附檢舉照片　"&Sys_CarColor
	else
		Response.Write "逕行舉發　"&Sys_A_Name&"<br>"
		Response.Write "附採證照片　"&Sys_CarColor
	End if 
%></div>
<div id="Layer54" style="position:absolute; left:300px; top:640px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:400px; top:640px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; left:295px; top:665px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; left:465px; top:665px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:660px; top:665px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:135px; top:685px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:300px; top:685px; width:140px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:480px; top:685px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" style="position:absolute; left:135px; top:710px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" style="position:absolute; left:135px; top:735px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer64" style="position:absolute; left:175px; top:735px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer65" style="position:absolute; left:215px; top:735px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer66" style="position:absolute; left:255px; top:735px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer67" style="position:absolute; left:295px; top:735px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>
<div id="Layer68" class="style4" style="position:absolute; left:435px; top:730px; width:620px; height:31px; z-index:20"><%
	If not ifnull(Sys_Jurgeday(0)) Then
		Response.Write "民眾檢舉日期 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"
	end If 
	
	if int(Sys_Rule1)=5610102 then Sys_IllegalRule1="在禁止臨時停車處所停車"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>4310209 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "<br>100公里以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "<br>80公里以上未滿100公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "<br>60公里以上未滿80公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "<br>40公里以上未滿60公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "<br>20公里以上未滿40公里"
			else
				response.write "<br>未滿20公里"
			end if
			'response.write "(經科學儀器採證)"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"

		Sys_IllegalRule1=LawRule(Sys_Rule1,Sys_IllegalRule1,Sys_CarSimpleID,Sys_CarAddID)

		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if	
	end if
	if not ifnull(Sys_Rule4) then response.write "("&Sys_Rule4&")"
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"

		Sys_IllegalRule2=LawRule(Sys_Rule2,Sys_IllegalRule2,Sys_CarSimpleID,Sys_CarAddID)

		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end if
'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
	If ifnull(Sys_Jurgeday(0)) Then response.write " (經科學儀器採證)"
%></div>
<div id="Layer69" style="position:absolute; left:135px; top:765px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:135px; top:795px; width:40px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer71" style="position:absolute; left:175px; top:795px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer72" style="position:absolute; left:215px; top:795px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer73" class="style4" style="position:absolute; left:435px; top:795px; width:400px; height:49px; z-index:29"><%

	Response.Write "　　　　　道 路 交 通 管 理 處 罰 條 例<br>"
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	'response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "第"&Mid(trim(Sys_Rule1),3,1)&"項第"&Mid(trim(Sys_Rule1),4,2)&"款規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		'response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "第"&Mid(trim(Sys_Rule2),3,1)&"項第"&Mid(trim(Sys_Rule2),4,2)&"款規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if

%></div>
<div id="Layer75" class="style4" style="position:absolute; left:435px; top:845px; width:90px; height:30px; z-index:28"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>
<div id="Layer74" style="position:absolute; left:515px; top:840px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>

<div id="Layer76" class="style3" style="position:absolute; left:435px; top:920px; width:110px; height:49px; z-index:29"><%
	If chk_UnitTypeID = "7100" or chk_UnitTypeID = "7200" or chk_UnitTypeID = "7300" or chk_UnitTypeID = "J01" Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" nowrap>臺南市政府警察局<br>"&SysUnit&"</td></tr>"
		
		If chk_UnitTypeID <> "7100" then
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"">TEL"&SysUnitTel&"</td></tr>"
		end if

		response.write "</table>"
	End if 
%></div>

<div id="Layer82" style="position:absolute; left:600px; top:910px; width:160px; height:43px; z-index:30"><%
'	If chk_UnitTypeID = "7400" Then
'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'		response.write "</table>"
'	End if 
%></div>

<div id="Layer77" class="style5" style="position:absolute; left:620px; top:920px; width:200px; height:46px; z-index:31"><%
	If chk_UnitTypeID = "7100" or chk_UnitTypeID = "7200" or chk_UnitTypeID = "7300" or chk_UnitTypeID = "J01" Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</td></tr>"
		response.write "</table>"
	End if 
%></div>
<div id="Layer78" style="position:absolute; left:175px; top:995px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer79" style="position:absolute; left:330px; top:995px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer80" style="position:absolute; left:480px; top:995px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer81" style="position:absolute; left:580px; top:995px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>
</div>
<%
	if (i mod 50)=0 then response.flush

	If i=300 Then
		If not ifnull(logBillNo) Then ConnExecute "舉發單列印："&logBillNo ,360
		logBillNo=""
	End if
	
next
end if

If not ifnull(logBillNo) Then ConnExecute "舉發單列印："&logBillNo ,360

If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,5.08,5.08,5.08,5.08);
</script>