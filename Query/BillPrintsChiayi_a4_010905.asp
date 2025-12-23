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
.style2 {font-family:"標楷體"; font-size: 16px; color:#ff0000; }
.style3 {font-size: 12px;}
.style4 {font-size: 10px;}
.style5 {font-size: 14px;}

-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,5,439,72">
</object>
<%
'on Error Resume Next
if trim(request("printStyle"))<>"" then
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,ILLEGALADDRESS,IllegalSpeed,RuleSpeed,INSURANCE,RuleVer,Note,BillFillDate from BillBase where SN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)

'===初始化(8/21)==
Sys_CarNo=""
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_IllegalSpeed="":Sys_RuleSpeed=""
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
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

if trim(Sys_UnitLevelID)="3" then
	chkJobID="303,314"
elseif trim(Sys_UnitLevelID)="1" then
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

strSQL="select UnitName,Tel from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
Unit.close

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
	If Not rscolor.eof then
		If Not IsNull(rscolor("Content")) then
			Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
		End If 
	End if
	rscolor.close
next

if trim(Sys_BillTypeID)="1" then
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
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
<div id="L78" class="pageprint" style="position:relative;"><%
if showBarCode then
%>
<div id="Layer1" style="position:absolute; left:50px; top:0px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:50px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:170px; top:5px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:20px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:675px; top:10px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:675px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:675px; top:40px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; left:35px; top:55px; width:233px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:470px; top:55px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<!--<div id="Layer11" class="style4" style="position:absolute; left:520px; top:95px; width:230px; height:12px; z-index:7"><%=BillPageUnit%>交字第<%="　　　　　　　　　　"%>號</div>-->

<div id="Layer12" class="style3" style="position:absolute; left:120px; top:130px; width:150px; height:11px; z-index:3">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></div>
<div id="Layer13" style="position:absolute; left:275px; top:125px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:385px; top:125px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>
<div id="Layer15" style="position:absolute; left:280px; top:145px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; left:450px; top:145px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:645px; top:145px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:120px; top:180px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:325px; top:180px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:560px; top:180px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" style="position:absolute; left:180px; top:210px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; left:120px; top:235px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:180px; top:235px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:240px; top:235px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:300px; top:235px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:360px; top:235px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style3" style="position:absolute; left:450px; top:235px; width:620px; height:31px; z-index:20"><%
	'response.write "<font size=2>"
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
			response.write " (經雷達、雷射測速儀器採證)"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
		
	end if
	response.write "<br>(期限內自動繳納處新台幣"&Sys_Level1&"元)"

	if trim(Sys_Rule2)<>"0" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
		response.write "<br>(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if				
	'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
	'response.write "</font>"
%></div>
<div id="Layer33" style="position:absolute; left:430px; top:275px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer28" style="position:absolute; left:120px; top:260px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:140px; top:280px; width:50px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:230px; top:280px; width:50px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:320px; top:280px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style3" style="position:absolute; left:550px; top:325px; width:400px; height:49px; z-index:29"><%
	'response.write "<font size='2'>"&left(trim(Sys_Rule1),2)
	response.write left(trim(Sys_Rule1),2)
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write "　　　　　　"
	response.write Mid(trim(Sys_Rule1),3,1)&"　　　　　　"&Mid(trim(Sys_Rule1),4,2)

	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write "　　　　　　"
		response.write Mid(trim(Sys_Rule2),3,1)&"　　　　　　"&Mid(trim(Sys_Rule2),4,2)

	end if
	'response.write "</font>"
%></div>

<div id="Layer34" class="style3" style="position:absolute; left:450px; top:360px; width:100px; height:30px; z-index:28"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer35" class="style5" style="position:absolute; left:455px; top:410px; width:110px; height:49px; z-index:29"><%="嘉義縣警察局<br>交通隊<br>05-3620229"%></div>

<div id="Layer36" style="position:absolute; left:580px; top:390px; width:160px; height:43px; z-index:30"><%
	'response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=63 height=25 align=""center""><span class=""style1"">隊長&nbsp;&nbsp;&nbsp;</span><br><span class=""style2"">陳建陽</span></td></tr>"
	'response.write "</table>"
%></div>

<div id="Layer37" class="style3" style="position:absolute; left:650px; top:410px; width:200px; height:46px; z-index:31"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=25 align=""center""><span class=""style2"">"&Sys_BillJobName&"&nbsp;</span><span class=""style2"">"&Sys_ChName&"</span></td></tr>"
	response.write "</table>　　"&Sys_BillFillerMemberID
%></div>
<div id="Layer38" style="position:absolute; left:140px; top:460px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:200px; top:460px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:260px; top:460px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:320px; top:460px; width:120px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>


<%if showBarCode then%>
<div id="Layer42" style="position:absolute; left:50px; top:520px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer43" style="position:absolute; left:50px; top:540px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:170px; top:525px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:170px; top:540px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:35px; top:580px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>

<div id="Layer51" style="position:absolute; left:470px; top:580px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--
<div id="Layer52" class="style4" style="position:absolute; left:505px; top:620px; width:230px; height:12px; z-index:7"><%=BillPageUnit%>交字第<%="　　　　　　　　　　"%>號</div>
-->

<div id="Layer53" class="style3" style="position:absolute; left:120px; top:655px; width:150px; height:11px; z-index:3">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></div>
<div id="Layer54" style="position:absolute; left:285px; top:650px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:385px; top:650px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; left:280px; top:675px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; left:450px; top:675px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:645px; top:675px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:120px; top:700px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:330px; top:700px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:560px; top:700px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" style="position:absolute; left:180px; top:730px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" style="position:absolute; left:120px; top:760px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer64" style="position:absolute; left:180px; top:760px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer65" style="position:absolute; left:240px; top:760px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer66" style="position:absolute; left:300px; top:760px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer67" style="position:absolute; left:360px; top:760px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer68" class="style3" style="position:absolute; left:450px; top:755px; width:620px; height:31px; z-index:20"><%
	'response.write "<font size=2>"
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
			response.write " (經雷達、雷射測速儀器採證)"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	
	end if
	response.write "<br>(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
		response.write "<br>(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if				
	'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
	'response.write "</font>"
%></div>

<div id="Layer74" style="position:absolute; left:430px; top:795px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>

<div id="Layer69" style="position:absolute; left:120px; top:785px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:140px; top:805px; width:50px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer71" style="position:absolute; left:230px; top:805px; width:50px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer72" style="position:absolute; left:320px; top:805px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer73" class="style3" style="position:absolute; left:550px; top:850px; width:400px; height:49px; z-index:29"><%
	'response.write "<font size='2'>"&left(trim(Sys_Rule1),2)
	response.write left(trim(Sys_Rule1),2)
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write "　　　　　　"
	response.write Mid(trim(Sys_Rule1),3,1)&"　　　　　　"&Mid(trim(Sys_Rule1),4,2)

	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write "　　　　　　"
		response.write Mid(trim(Sys_Rule2),3,1)&"　　　　　　"&Mid(trim(Sys_Rule2),4,2)

	end if
	'response.write "</font>"
%></div>

<div id="Layer75" class="style3" style="position:absolute; left:450px; top:890px; width:100px; height:30px; z-index:26"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer76" class="style3" style="position:absolute; left:550px; top:930px; width:110px; height:49px; z-index:29"><%
	response.write "嘉義縣警察局<br>交通隊<br>05-3620229"
%></div>
<div id="Layer82" style="position:absolute; left:500px; top:910px; width:160px; height:43px; z-index:30"><%
	'response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=63 height=25 align=""center""><span class=""style1"">隊長&nbsp;&nbsp;&nbsp;</span><br><span class=""style2"">陳建陽</span></td></tr>"
	'response.write "</table>"
%></div>

<div id="Layer77" class="style3" style="position:absolute; left:650px; top:940px; width:200px; height:46px; z-index:31"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=25 align=""center""><span class=""style2"">"&Sys_BillJobName&"&nbsp;</span><span class=""style2"">"&Sys_ChName&"</span></td></tr>"
	response.write "</table>　　"&Sys_BillFillerMemberID
%></div>
<div id="Layer78" style="position:absolute; left:140px; top:985px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer79" style="position:absolute; left:200px; top:985px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer80" style="position:absolute; left:260px; top:985px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer81" style="position:absolute; left:320px; top:985px; width:120px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>
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
	printWindow(true,8,5.08,5.08,5.08);
</script>