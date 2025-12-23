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
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style6 {font-size: 20px; line-height:2;}
.style7 {font-size: 16px}
.style8 {font-size: 36px}
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style11 {font-size: 14px}
.style12 {font-family:"標楷體"; font-size: 8px; color:#ff0000; }
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
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

'on Error Resume Next
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close
PrintSum=0

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

for i=0 to ubound(PBillSN)
sumCnt=sumCnt+1
if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+PrintSum)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Rule4=""
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
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then
	Sys_Jurgeday=split(gArrDT(trim(rs("Jurgeday"))),"-")
else
	Sys_Jurgeday=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then Sys_RecordMemberID=trim(rs("RecordMemberID"))
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

If Not ifnull(trim(rsfound("Driver"))) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
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
Sum_Level=0:Sys_Rule1="":Sys_Rule2=""
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

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
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

'strSQL="select * from BillIllegalImage where billsn="&trim(rsbil("BillSN"))
'set rsimage=conn.execute(strSQL)
'if Not rsimage.eof then
'	Sys_IisImagePath=trim(rsimage("IisImagePath"))
'	Sys_ImageFileNameA=trim(rsimage("ImageFileNameA"))
'	Sys_ImageFileNameB=trim(rsimage("ImageFileNameB"))
'end if

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

if trim(Sys_UnitLevelID)="1" then
	chkJobID="303"
else
	chkJobID="303,314"
end if

strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"
'response.write strSQL
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

'strSQL="select UnitName from UnitInfo where UnitTypeID='"&DBUnitTypeID&"' and UnitName like '%分局'"
'set unit=conn.execute(strSQL)
'If not unit.eof Then DBUnitName=unit("UnitName")
'unit.close

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

Sys_BillUnitAddress=""

strSQL="select * from UnitInfo where UnitID in(select UnitTypeID from UnitInfo where UnitID in(select UnitID from MemberData where MemberID="& Sys_RecordMemberID &"))"
set mem=conn.execute(strSQL)
if Not mem.eof then Sys_BillUnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
mem.close

sys_UnitLevelAddr=""
strSQL="select * from UnitInfo where UnitID in(select UnitID from MemberData where MemberID="& Sys_RecordMemberID &")"
set mem=conn.execute(strSQL)
if Not mem.eof then sys_UnitLevelAddr=trim(mem("UnitLevelID"))
mem.close


If ifnull(Sys_BillUnitAddress) Then
	response.write "<font size=""10"">請至單位管理填寫單位地址才能繼續列印！！</font>"
	response.end
end if

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
If ifnull(Sys_MailNumber) Then Sys_MailNumber=0
if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,36
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",400,451,36"
	'response.end
end if
DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,60,160

Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo
strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;">

<div id="Layer1" style="position:absolute; left:60px; top:20px; z-index:5">
<%
Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_1.jpg"" hspace=""0"" vspace=""0"">"
Response.Write "<br><span class=""style7"">"&Sys_FirstBarCode&"</span>"
%>
</div>
<div id="Layer2" class="style3" style="position:absolute; left:250px; top:20px; width:350px; height:36px; z-index:5"><%
If not ifnull(request("Sys_UnitLabelKind")) Then
	response.write "<b>"&Sys_BillUnitAddress&"<br>"&Sys_BillUnitName&"</b>"
End if%>
</div>

<div id="Layer2" style="position:absolute; left:250px; top:70px; width:350px; height:36px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	if trim(Sys_BillTypeID)="1" then
		Response.Write Sys_DriverHomeZip&Sys_DriverZipName&funcCheckFont(Sys_DriverHomeAddress,20,1)&"<br>"
		Response.Write funcCheckFont(Sys_Driver,20,1)&"　台啟"
	elseif trim(Sys_BillTypeID)="2" then
		Response.Write Sys_OwnerZip&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,1)&"<br>"
		Response.Write funcCheckFont(Sys_Owner,20,1)&"　台啟"
	end if
	Response.Write "</span>"%>
</div>
<div id="Layer3" style="position:absolute; left:320px; top:125px; z-index:5"><%
	'If trim(Session("UnitLevelID"))="1" Then
	Response.Write "<span class=""style4"">"
	Response.Write "大宗郵資已付掛號函件<br>　　　第"&Sys_MailNumber&"號"
	Response.Write "</span>"
	'End if
%>
</div>
<div id="Layer4" style="position:absolute; left:300px; top:155px; z-index:5"><%
	'If trim(Session("UnitLevelID"))="1" Then
	Response.Write "<span class=""style3"">"
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""><br>"
	Response.Write Sys_MAILCHKNUMBER
	Response.Write "</span><br>"
	'end if
%>
</div>

<div id="Layer5" style="position:absolute; left:60px; top:180px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write "應到案處所："&Sys_STATIONNAME&"<br>"
   	Response.Write "應到案處所電話："&Sys_StationTel
	Response.Write "</span>"

%>
</div>
<div id="Layer5" style="position:absolute; left:550px; top:160px; z-index:5"><%
	Response.Write "<span class=""style8"">"
	Response.Write Sys_StationID
	Response.Write "</span>"
%>
</div>

<div id="Layer6" style="position:absolute; left:60px; top:320px; width:400px; height:36px; z-index:5"><span class="style7">查詢電話：<%=Sys_BillUnitTel%>（<%=DB_UnitName%>）</span></div>
<!--<div id="Layer1" style="position:absolute; left:50px; top:810px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<!--<div id="Layer2" style="position:absolute; left:50px; top:840px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%'if trim(Sys_BillTypeID)="1" then%>
<!--	<div id="Layer3" style="position:absolute; left:165px; top:815px; width:202px; height:36px; z-index:5">Ｖ</div>
<%'else%>
<div id="Layer4" style="position:absolute; left:165px; top:830px; width:202px; height:36px; z-index:5">Ｖ</div>
<%'end if%>
<div id="Layer5" style="position:absolute; left:165px; top:845px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer6" style="position:absolute; left:635px; top:410px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer7" style="position:absolute; left:635px; top:425px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer8" style="position:absolute; left:635px; top:440px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>
<div id="Layer9" style="position:absolute; left:20px; top:425px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_BillUnitName
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:415px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:590px; top:895px; width:230px; height:12px; z-index:7"><font size=1><%=Sys_BillNo%></font></div>-->
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer12" style="position:absolute; left:115px; top:510px; width:150px; height:11px; z-index:3"><span class="style7"><%=Sys_Driver%></span></div>
<%end if
if trim(Sys_BillTypeID)="1" then%>
<div id="Layer13" style="position:absolute; left:265px; top:495px; width:28px; height:11px; z-index:3"><span class="style7"><%=Sys_Sex%></span></div>
<div id="Layer14" style="position:absolute; left:360px; top:490px; width:324px; height:10px; z-index:4"><span class="style7"><%=Sys_DriverZipName&Sys_DriverHomeAddress%></Span></div>
<div id="Layer15" style="position:absolute; left:265px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(0)%></span></div>
<div id="Layer142" style="position:absolute; left:295px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(1)%></span></div>
<div id="Layer143" style="position:absolute; left:325px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(2)%></span></div>
<div id="Layer16" style="position:absolute; left:430px; top:520px; width:106px; height:13px; z-index:9"><span class="style7"><%=Sys_DriverID%></span></div>
<div id="Layer17" style="position:absolute; left:630px; top:520px; width:99px; height:12px; z-index:10"><span class="style7"><%=fastring%></span></div>
<%end if%>
<div id="Layer18" style="position:absolute; left:125px; top:545px; width:100px; height:14px; z-index:11"><span class="style7"><%=Sys_CarNo%></span></div>
<div id="Layer19" style="position:absolute; left:310px; top:545px; width:117px; height:20px; z-index:12"><span class="style7"><%=Sys_DCIRETURNCARTYPE%></span></div>
<div id="Layer20" style="position:absolute; left:570px; top:545px; width:201px; height:17px; z-index:13"><span class="style7"><%=funcCheckFont(Sys_Owner,22,1)%></span></div>
<div id="Layer21" style="position:absolute; left:165px; top:570px; width:507px; height:13px; z-index:14"><span class="style7"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></span></div>

<div id="Layer22" style="position:absolute; left:105px; top:600px; width:40px; height:13px; z-index:15"><span class="style7"><%=Sys_IllegalDate(0)%></span></div>
<div id="Layer23" style="position:absolute; left:160px; top:600px; width:40px; height:17px; z-index:16"><span class="style7"><%=Sys_IllegalDate(1)%></span></div>
<div id="Layer24" style="position:absolute; left:220px; top:600px; width:40px; height:16px; z-index:17"><span class="style7"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:270px; top:600px; width:40px; height:16px; z-index:18"><span class="style7"><%=right("00"&Sys_IllegalDate_h,2)%></span></div>
<div id="Layer26" style="position:absolute; left:330px; top:600px; width:40px; height:13px; z-index:19"><span class="style7"><%=right("00"&Sys_IllegalDate_m,2)%></span></div>
<div id="Layer27" style="position:absolute; left:430px; top:600px; width:270px; height:31px; z-index:20"><span class="style3"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "該路段限速"&Sys_RuleSpeed&"公里、經雷達(射)測速為"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
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
		Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if int(Sys_Rule1)=5610102 then Sys_IllegalRule1="在禁止臨時停車處所停車"
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
			
'	if trim(Sys_Note)<>"" then response.write "("&Sys_Note&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:110px; top:625px; width:217px; height:15px; z-index:21"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:140px; top:645px; width:34px; height:11px; z-index:22"><span class="style7"><%=Sys_DealLineDate(0)%></span></div>
<div id="Layer30" style="position:absolute; left:220px; top:645px; width:35px; height:13px; z-index:23"><span class="style7"><%=Sys_DealLineDate(1)%></span></div>
<div id="Layer31" style="position:absolute; left:300px; top:645px; width:32px; height:15px; z-index:24"><span class="style7"><%=Sys_DealLineDate(2)%></span></div>
<div id="Layer32" style="position:absolute; left:430px; top:685px; width:400px; height:49px; z-index:29"><span class="style4"><%response.write "第"&left(trim(Sys_Rule1),2)&"條"
			'if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款規定"
				response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
			if trim(Sys_Rule2)<>"0" then
				response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
				'if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
				response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款規定"
				response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
			end if
			%></span></div>
<div id="Layer33" style="position:absolute; left:420px; top:730px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer34" style="position:absolute; left:660px; top:740px; width:90px; height:30px; z-index:28"><span class="style3"><%=Sys_STATIONNAME%></span></font></div>

<div id="Layer35" style="position:absolute; left:423px; top:785px; width:100px; height:49px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	if Session("Unit_ID") <>"0207" and Session("Unit_ID") <>"0240" then 
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style10"">&nbsp;"&Sys_UnitName&"&nbsp;</span><br><span class=""style10"">&nbsp;"&Sys_UnitTEL&"&nbsp;</span></td></tr>"
	end if
	response.write "</table>"
%></div>
<div id="Layer36" style="position:absolute; left:500px; top:745px; width:100px; height:43px; z-index:30"><%
'	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'	if Session("Unit_ID") <>"0207" then 
'		response.write "<tr><td nowrap style=""border-color:#ff0000;border-style:solid;"" align=""center""><span class=""style9"">主管職名章</span><br><span class=""style10"">"&Sys_JobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'	end if
'	response.write "</table>"
%></div>
<div id="Layer37" style="position:absolute; left:615px; top:760px; width:200px; height:46px; z-index:31"><%
'	if trim(Sys_MemberFilename)<>"" then
'		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""110"" height=""40"">"
'	else
'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=63 height=25 align=""center""><span class=""style9"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
'		response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
'	end if
%></div>

<div id="Layer47" class="style12" style="position:absolute; left:240px; top:825px; width:200px; height:10px; z-index:32">(自103年3月31日起，前、後段日數均改為30日)</div>
<div id="Layer38" style="position:absolute; left:135px; top:835px; width:60px; height:10px; z-index:32"><span class="style7"><%=sys_Date(0)%></span></div>
<div id="Layer39" style="position:absolute; left:195px; top:835px; width:60px; height:13px; z-index:33"><span class="style7"><%=sys_Date(1)%></span></div>
<div id="Layer40" style="position:absolute; left:255px; top:835px; width:60px; height:11px; z-index:34"><span class="style7"><%=sys_Date(2)%></span></div>
<div id="Layer41" style="position:absolute; left:300px; top:835px; width:80px; height:12px; z-index:36"><span class="style7"><%
	if Session("Unit_ID") <>"0207" then Response.Write Sys_BillFillerMemberID
%></span></div>


<div id="Layer45" style="position:absolute; left:180px; top:1005px; width:100px; height:12px; z-index:36"><span class="style3"><%
	Response.Write Sys_BillNo
%></span></div>

<div id="Layer43" style="position:absolute; left:210px; top:1022px; width:350px; height:12px; z-index:36"><span class="style3"><%=Sys_MAILCHKNUMBER%></span></div>

<div id="Layer44" style="position:absolute; left:370px; top:1030px; width:350px; height:12px; z-index:10"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>></div>

<div id="Layer45" style="position:absolute; left:185px; top:1050px; width:100px; height:12px; z-index:36"><span class="style3"><%
	Response.Write funcCheckFont(Sys_Owner,22,1)
%></span></div>

<div id="Layer42" style="position:absolute; left:185px; top:1070px; width:230px; height:12px; z-index:36; background-color:#FFFFFF"><span class="style3"><%
	Response.Write Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)
%></span></div>

<div id="Layer46" style="position:absolute; left:450px; top:1140px; width:230px; height:12px; z-index:36; background-color:#FFFFFF"><span class="style3"><%
	Response.Write "<B>□ 本人　□ 代收</B>"
%></span></div>

<div id="Layer42" style="position:absolute; left:190px; top:1285px; width:500px; z-index:36; background-color:#FFFFFF"><span class="style7"><%
If cdbl(sys_UnitLevelAddr) > 1 Then
	response.write "<font color=""red"">請繳回："
	Sys_BillUnitAddress=replace(Sys_BillUnitAddress,"一","１")
	Sys_BillUnitAddress=replace(Sys_BillUnitAddress,"二","２")
	Sys_BillUnitAddress=replace(Sys_BillUnitAddress,"三","３")
	Sys_BillUnitAddress=replace(Sys_BillUnitAddress,"四","４")
	Response.Write Sys_BillUnitAddress
	Response.Write "　　"
	
	Sys_BillUnitName=replace(Sys_BillUnitName,"一","１")
	Sys_BillUnitName=replace(Sys_BillUnitName,"二","２")
	Sys_BillUnitName=replace(Sys_BillUnitName,"三","３")
	Sys_BillUnitName=replace(Sys_BillUnitName,"四","４")
	Response.Write Sys_BillUnitName
	Response.Write "</font>"
End if
%></span></div>

</div>

<%
'	If trim(Sys_BillUnitTypeID) = "0207" Then
'		if (i mod 30)=0 then response.flush
'	else
		response.flush
'	End if 
	
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
	printWindow(true,5.08,5.08,5.08,5.08);
</script>