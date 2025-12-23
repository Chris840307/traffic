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
.style1 {font-size: 9px}
.style2 {font-size: 10px}
.style3 {font-size: 16px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 18px}
.style11 {font-size: 14px}
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
codebase="..\smsx.cab#Version=6,1,432,1">
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
PrintSum=request("hd_PrintSum"):Sys_Print=request("Sys_Print")
PrintSum=PrintSum-Sys_Print
if trim(Sys_Print)="" then
	PrintSum=0:Sys_Print=Ubound(PBillSN)+1
elseif Cint(request("hd_PrintSum"))-1>Ubound(PBillSN) then
	Sys_Print=Ubound(PBillSN)-(Cint(request("hd_PrintSum"))-Sys_Print)
end if
for i=0 to Cint(Sys_Print-1)
sumCnt=sumCnt+1
if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+PrintSum)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_RecordMemberID=trim(rs("RecordMemberID"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
'response.write strSql
'response.end
set rsfound=conn.execute(strSql)
'if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver"))
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
if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close
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

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='N'"
set rstwo=conn.execute(strSql)
if not rstwo.eof then
	rstwo.close
	strSql="select distinct DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and DriverHomeZip is not null"
	set rstwo=conn.execute(strSql)
	if not rstwo.eof then
		Sys_OwnerAddressDeliver=trim(rstwo("DriverHomeAddress"))
		Sys_OwnerZipDeliver=trim(rstwo("DriverHomeZip"))
		strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZipDeliver&"'"
		set rszip=conn.execute(strSQL)
		if Not rszip.eof then Sys_OwnerZipNameDeliver=trim(rszip("ZipName"))
		rszip.close
	else
		Sys_OwnerAddressDeliver=Sys_OwnerAddress
		Sys_OwnerZipDeliver=Sys_OwnerZip
		Sys_OwnerZipNameDeliver=Sys_OwnerZipName
	end if
	rstwo.close
else
	Sys_OwnerAddressDeliver=Sys_OwnerAddress
	Sys_OwnerZipDeliver=Sys_OwnerZip
	Sys_OwnerZipNameDeliver=Sys_OwnerZipName
	rstwo.close
end if
rsfound.close
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
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

strSQL="select UnitName,Address,Tel from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysAddress=Unit("Address")
SysUnitTel=Unit("Tel")
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

strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

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

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" class="pageprint" style="position:relative;"><%
if showBarCode then
%>
<div id="Layer1" style="position:absolute; left:0px; top:15px; width:10px; height:20px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:0px; top:40px; width:10px; height:20px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:120px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer4" style="position:absolute; left:120px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:625px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:625px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:625px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:0px; top:70px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>
<!--<div id="Layer42" style="position:absolute; left:210px; top:70px; width:202px; height:36px; z-index:5"><%="<font size=1>"&SysUnit&"<br>("&SysUnitTel&")</font>"%></div>-->
<div id="Layer10" style="position:absolute; left:450px; top:70px; width:233px; height:32px; z-index:6"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:485px; top:110px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer12" style="position:absolute; left:50px; top:130px; width:150px; height:11px; z-index:3">逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer13" style="position:absolute; left:210px; top:130px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:320px; top:130px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>
<div id="Layer15" style="position:absolute; left:200px; top:160px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:355px; top:160px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:550px; top:160px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:50px; top:175px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:215px; top:175px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:385px; top:175px; width:251px; height:17px; z-index:13"><%=chstr(Sys_Owner)%></div>
<div id="Layer21" style="position:absolute; left:50px; top:195px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&chstr(Sys_OwnerZipName&Sys_OwnerAddress)%></div>

<div id="Layer22" style="position:absolute; left:60px; top:220px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:110px; top:220px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:160px; top:220px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:210px; top:220px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:260px; top:220px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:345px; top:220px; width:600px; height:31px; z-index:20"><%
	response.write "<font size=2>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if
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
	response.write "</font>"
%></div>
<div id="Layer28" style="position:absolute; left:50px; top:260px; width:280px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:90px; top:290px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:170px; top:290px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:250px; top:290px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" style="position:absolute; left:360px; top:295px; width:400px; height:49px; z-index:29"><%
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

<div id="Layer34" style="position:absolute; left:345px; top:335px; width:95px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>
<div id="Layer33" style="position:absolute; left:445px; top:330px; width:400px; height:30px; z-index:28"><%if showBarCode then response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_5.jpg"">"%></div>

<!--<div id="Layer35" style="position:absolute; left:455px; top:420px; width:100px; height:49px; z-index:29"><%
	'if billprintuseimage=1 then
		'response.write "<img src=""..\UnitInfo\Picture\"&Sys_UnitFilename&""" width=""70"" height=""70"">"
	'else
		'response.write Sys_UnitName
	'end if%></div>
<div id="Layer36" style="position:absolute; left:580px; top:420px; width:100px; height:43px; z-index:30">主管</div>
<div id="Layer37" style="position:absolute; left:660px; top:410px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""..\Member\Picture\"&Sys_MemberFilename&""" width=""90"" height=""30"">"
	else
		response.write Sys_ChName
	end if
%></div>-->
<div id="Layer38" style="position:absolute; left:220px; top:485px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:380px; top:485px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:530px; top:485px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:635px; top:485px; width:80px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>

<div style="position:absolute; left:60px; top:650px;">
<table width="645" height="393" border="0">
  <tr>
    <td height="69" valign="top">&nbsp;</td>
    <th align="left" colspan="2"><span class="style8">900<br>屏東縣<br>　屏東縣<%=SysAddress%></span></th>
    <td rowspan="2" align="left" valign="top"></td>
  </tr>
  <tr>
    <td align="left" valign="top"></td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th align="left" colspan="2"><span class="style8"><%=Sys_Driver%>　台啟</span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left"><span class="style8"><%=Sys_Owner%>　台啟</span></th>
	<%end if%>
  </tr>
  <tr>
    <td>&nbsp;</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th colspan="2" align="left" valign="top" nowrap><span class="style8"><%=Sys_DriverHomeZip%><br>
    <%=Sys_DriverZipName&Sys_DriverHomeAddress%></span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left" valign="top" nowrap><span class="style8"><%=Sys_OwnerZip%><br>
    <%=Sys_OwnerZipName&Sys_OwnerAddress%></span></th>
	<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"></td>
    <td align="center">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"></td>
    <td align="center">&nbsp;</td>
    <td align="right"><p>&nbsp;</p></td>
  </tr>
  <tr>
    <td height="98" valign="top" nowrap><p>&nbsp;</p></td>
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</div>

<div id="Layer43" style="position:absolute; left:310px; top:1090px; width:280px; height:12px; z-index:36"><%=SysUnit%></div>
<div id="Layer44" style="position:absolute; left:155px; top:1110px; width:400px; height:12px; z-index:36"><%
	if Sys_BillTypeID="1" then
		response.write "<font size=2>"&Sys_Driver&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;車號："&Sys_CarNo&"</font>"
		response.write "<font size=2><br>"&Sys_DriverHomeZip&"&nbsp;&nbsp;"&Sys_DriverZipName&Sys_DriverHomeAddress&"</font>"
	else
		response.write "<font size=2>"&Sys_Owner&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;車號："&Sys_CarNo&"</font>"
		response.write "<font size=2><br>"&Sys_OwnerZipDeliver&"&nbsp;&nbsp;"&Sys_OwnerZipNameDeliver&Sys_OwnerAddressDeliver&"</font>"
	end if%></div>
<div id="Layer45" style="position:absolute; left:210px; top:1150px; width:280px; height:12px; z-index:36"><%=Sys_BillNo%></div>
<div id="Layer46" style="position:absolute; left:400px; top:1140px; width:233px; height:32px; z-index:6"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"""%>></div>
</div>
<%next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,0,5.08,5.08,5.08);
</script>