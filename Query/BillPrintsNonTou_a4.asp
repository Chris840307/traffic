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
</head>

<body>
<%
on Error Resume Next
if trim(request("printStyle"))<>"" then
PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,ILLEGALADDRESS,IllegalSpeed,RuleSpeed,INSURANCE,RuleVer,Note,BillFillDate from BillBase where SN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
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
if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close
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

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSQL="select UnitName from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
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

StrBass="select a.A_Name,b.Content as CarColor,c.ID as CarStatusID,c.Content as CarStatusName,d.ID as Rule4,d.Content as Rule4Name,e.DCIStationName from BillBaseDCIReturn a,(select ID,Content from DCICode where TypeID=4) b,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DciReturnCarColor=b.ID(+) and a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+) and a.EXCHANGETYPEID='A' and a.CarNo='"&trim(rsbil("CarNo"))&"'"

Sys_A_Name="":Sys_CarColor=""
set rsCarType=conn.execute(strBass)
if not rsCarType.eof then
	Sys_A_Name=trim(rsCarType("A_Name"))
	Sys_CarColor=trim(rsCarType("CarColor"))
end if
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
%>
<div id="Layer1" style="position:absolute; left:70px; top:<%=5+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<!--<div id="Layer2" style="position:absolute; left:70px; top:<%=35+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:195px; top:<%=5+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer4" style="position:absolute; left:195px; top:<%=25+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:<%=45+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:675px; top:<%=10+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:675px; top:<%=25+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:675px; top:<%=40+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; left:35px; top:<%=55+1084*i%>px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>><%if billprintUnitID=1 then resposne.write "　　　　　　"&SysUnit%></div>
<div id="Layer10" style="position:absolute; left:550px; top:<%=55+1084*i%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" style="position:absolute; left:510px; top:<%=95+1084*i%>px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>
<div id="Layer12" style="position:absolute; left:130px; top:<%=115+1084*i%>px; width:150px; height:11px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer13" style="position:absolute; left:285px; top:<%=115+1084*i%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:385px; top:<%=115+1084*i%>px; width:324px; height:10px; z-index:4"><font size=2><%=Sys_DriverZipName&Sys_DriverHomeAddress%>*本單可至郵局或委託代收之超商繳納</font></div>
<div id="Layer15" style="position:absolute; left:290px; top:<%=135+1084*i%>px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; left:460px; top:<%=135+1084*i%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:655px; top:<%=135+1084*i%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:130px; top:<%=160+1084*i%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:295px; top:<%=160+1084*i%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:475px; top:<%=160+1084*i%>px; width:251px; height:17px; z-index:13"><%=Sys_Owner%></div>
<div id="Layer21" style="position:absolute; left:140px; top:<%=180+1084*i%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZipName&Sys_OwnerAddress%></div>

<div id="Layer22" style="position:absolute; left:130px; top:<%=205+1084*i%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer23" style="position:absolute; left:185px; top:<%=205+1084*i%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer24" style="position:absolute; left:225px; top:<%=205+1084*i%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer25" style="position:absolute; left:265px; top:<%=205+1084*i%>px; width:40px; height:16px; z-index:18"><%=Sys_IllegalDate_h%>時</div>
<div id="Layer26" style="position:absolute; left:310px; top:<%=205+1084*i%>px; width:40px; height:13px; z-index:19"><%=Sys_IllegalDate_m%>分</div>
<div id="Layer27" style="position:absolute; left:425px; top:<%=215+1084*i%>px; width:600px; height:31px; z-index:20"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "、超過20公里"
			else
				response.write "、未滿20公里"
			end if
		end if
	else
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"

		
	end if
	'smith edit for print two rule 20070621
	if len(Sys_IllegalRule2)<25 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule2))
		end if	
'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>
<div id="Layer28" style="position:absolute; left:130px; top:<%=235+1084*i%>px; width:217px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:130px; top:<%=270+1084*i%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer30" style="position:absolute; left:185px; top:<%=270+1084*i%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer31" style="position:absolute; left:225px; top:<%=270+1084*i%>px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer32" style="position:absolute; left:425px; top:<%=275+1084*i%>px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>

<div id="Layer33" style="position:absolute; left:425px; top:<%=315+1084*i%>px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>

<div id="Layer34" style="position:absolute; left:665px; top:<%=320+1084*i%>px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME%></font></div>


<div id="Layer35" style="position:absolute; left:405px; top:<%=380+1084*i%>px; width:120px; height:49px; z-index:29"><%
	if billprintuseimage=1 then
		if trim(Sys_UnitFilename)<>"" then
			response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""70"" height=""70"">"
		end if
	else
		response.write "<font size=1>"&Sys_City&"<br>"&SysUnit&"</font>"
	end if%></div>
<!--<div id="Layer36" style="position:absolute; left:580px; top:<%=420+1084*i%>px; width:100px; height:43px; z-index:30">主管</div>-->
<div id="Layer37" style="position:absolute; left:660px; top:<%=410+1084*i%>px; width:200px; height:46px; z-index:31"><%
	if billprintuseimage=1 then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		end if
	end if
%></div>
<div id="Layer38" style="position:absolute; left:175px; top:<%=470+1084*i%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer39" style="position:absolute; left:330px; top:<%=470+1084*i%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer40" style="position:absolute; left:480px; top:<%=470+1084*i%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer41" style="position:absolute; left:580px; top:<%=470+1084*i%>px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>



<div id="Layer42" style="position:absolute; left:65px; top:<%=530+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<!--<div id="Layer43" style="position:absolute; left:65px; top:<%=555+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->

<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:190px; top:<%=535+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:190px; top:<%=550+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:<%=565+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:<%=535+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:<%=550+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:<%=565+1084*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:35px; top:<%=580+1084*i%>px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:550px; top:<%=580+1084*i%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer52" style="position:absolute; left:510px; top:<%=620+1084*i%>px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>
<div id="Layer53" style="position:absolute; left:130px; top:<%=640+1084*i%>px; width:150px; height:11px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer54" style="position:absolute; left:295px; top:<%=640+1084*i%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:395px; top:<%=640+1084*i%>px; width:324px; height:10px; z-index:4"><font size=2><%=Sys_DriverZipName&Sys_DriverHomeAddress%></font></div>
<div id="Layer56" style="position:absolute; left:290px; top:<%=665+1084*i%>px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; left:460px; top:<%=665+1084*i%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:655px; top:<%=665+1084*i%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:130px; top:<%=685+1084*i%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:295px; top:<%=685+1084*i%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:475px; top:<%=685+1084*i%>px; width:251px; height:17px; z-index:13"><%=Sys_Owner%></div>
<div id="Layer62" style="position:absolute; left:130px; top:<%=710+1084*i%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZipName&Sys_OwnerAddress%></div>

<div id="Layer63" style="position:absolute; left:130px; top:<%=735+1084*i%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer64" style="position:absolute; left:185px; top:<%=735+1084*i%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer65" style="position:absolute; left:225px; top:<%=735+1084*i%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer66" style="position:absolute; left:265px; top:<%=735+1084*i%>px; width:40px; height:16px; z-index:18"><%=Sys_IllegalDate_h%>時</div>
<div id="Layer67" style="position:absolute; left:310px; top:<%=735+1084*i%>px; width:40px; height:13px; z-index:19"><%=Sys_IllegalDate_m%>分</div>
<div id="Layer68" style="position:absolute; left:425px; top:<%=750+1084*i%>px; width:600px; height:31px; z-index:20"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "、超過20公里"
			else
				response.write "、未滿20公里"
			end if
		end if
	else
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"
	
	end if
		'smith edit for print two rule
		if len(Sys_IllegalRule2)<25 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule2))
		end if	
'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>
<div id="Layer69" style="position:absolute; left:130px; top:<%=765+1084*i%>px; width:217px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:130px; top:<%=800+1084*i%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer71" style="position:absolute; left:185px; top:<%=800+1084*i%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer72" style="position:absolute; left:225px; top:<%=800+1084*i%>px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer73" style="position:absolute; left:425px; top:<%=805+1084*i%>px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>

<div id="Layer74" style="position:absolute; left:425px; top:<%=845+1084*i%>px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>
<div id="Layer75" style="position:absolute; left:665px; top:<%=850+1084*i%>px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME%></font></div>

<div id="Layer76" style="position:absolute; left:425px; top:<%=920+1084*i%>px; width:120px; height:49px; z-index:29"><%
	if billprintuseimage=1 then
		if trim(Sys_UnitFilename)<>"" then
			response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""70"" height=""70"">"
		end if
	else
		response.write "<font size=1>"&Sys_City&"<br>"&SysUnit&"</font>"
	end if%></div>
<div id="Layer77" style="position:absolute; left:630px; top:<%=930+1084*i%>px; width:200px; height:46px; z-index:31"><%
	if billprintuseimage=1 then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		end if
	end if
%></div>
<div id="Layer78" style="position:absolute; left:175px; top:<%=995+1084*i%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer79" style="position:absolute; left:330px; top:<%=995+1084*i%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer80" style="position:absolute; left:480px; top:<%=995+1084*i%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer81" style="position:absolute; left:580px; top:<%=995+1084*i%>px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>
<%next
end if%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>