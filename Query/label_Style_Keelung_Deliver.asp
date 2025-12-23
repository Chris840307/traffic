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
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

PBillSN=split(trim(request("PBillSN")),",")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

Server.ScriptTimeout=6000

for i=0 to Ubound(PBillSN)
if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sum_Level=0
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_CarNo=trim(rs("CarNo"))
if Not rs.eof then Sys_Rule1=trim(rs("Rule1"))
if Not rs.eof then Sys_Rule2=trim(rs("Rule2"))
if Not rs.eof then Sys_Level1=trim(rs("FORFEIT1"))
if Not rs.eof then Sys_Level2=trim(rs("FORFEIT2"))
Sum_Level=funTnumber(Sys_Level1)+funTnumber(Sys_Level2)
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
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
Sys_DriverHomeAddress="":Sys_DriverHomeZip=""
if Instr(request("Sys_BatchNumber"),"N")>0 then
	strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
else
	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"
end if

set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Driver=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))

if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsfound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsfound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
	End if
end if
rsfound.close

If ifnull(Sys_DriverHomeAddress) or ifnull(Sys_Driver) Then
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

	elseif Sys_BillTypeID=2 then
		strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

	else
		strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

	end if

	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID=1 Then
		if Not rsdata.eof then Sys_Driver=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Driver=trim(rsdata("Owner"))
	End if
	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then Sys_DriverHomeAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_DriverHomeZip=trim(rsdata("DriverHomeZip"))
	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_DriverHomeAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_DriverHomeZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then Sys_DriverHomeAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_DriverHomeZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
If ifnull(Sys_Driver) Then
	Sys_Driver=trim(rsfound("Owner"))
end if
If ifnull(Sys_DriverHomeAddress) Then
	if Not rsfound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
end if

	Sys_DriverZipName=""
	strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof then Sys_DriverZipName=replace(trim(rszip("ZipName")),"臺","台")
	rszip.close
'	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
'	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
'	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
'	strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
'	set rszip=conn.execute(strSQL)
'	if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
'	rszip.close
	If not ifnull(Sys_DriverHomeAddress) Then
		Sys_DriverHomeAddress=replace(replace(Sys_DriverHomeAddress,"臺","台"),Sys_DriverZipName,"")
	End if
'end if

Sys_DCIReturnStation=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
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
end If 

Sys_BillUnitAddress=""

strSQL="select * from UnitInfo where UnitID in(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"')"
set mem=conn.execute(strSQL)
if Not mem.eof then Sys_BillUnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
mem.close

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

if ifnull(Sys_Note) then Sys_Note=""

'if sys_City="台中市" then
'	strSql="select StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
'	set rs=conn.execute(strSql)
'	if Not rs.eof then Sys_MailNumber=trim(rs("StoreAndSendMailNumber"))
'	if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
'	if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
'else
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		MailKindType=17
	else
		MailKindType=36
	end if

	If Not ifnull(request("Sys_LabelKind")) and sys_City<>"台中市" then
		if instr(Sys_Note,"郵寄日")<=0 then
			strSql="select MailNumber_Sn.NextVal StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
			set rs=conn.execute(strSql)
			if Not rs.eof then Sys_MailNumber=trim(rs("StoreAndSendMailNumber"))
			if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
			if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
		else
			strSql="select MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
			set rs=conn.execute(strSql)
			if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
			if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
			Sys_MailNumber=mid(Sys_Note,instr(Sys_Note,"大宗")+3,6)
		end if
	elseif Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MailNumber=trim(rs("StoreAndSendMailNumber"))
		if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
		if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
	else
		strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
		if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
		if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
	end if
'end if

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
if sys_City="台中市" then
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,MailKindType
elseif sys_City="基隆市" then
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,MailKindType
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",400,451,17"
	'response.end
end if
DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,60,160

Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo
Sys_MAILCHKNUMBER=""
'if sys_City="台中市" then
'	strSql="select OpenGOVReportnumber from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
'	set rs=conn.execute(strSql)
'	if Not rs.eof then Sys_MAILCHKNUMBER=left(trim(rs("OpenGOVReportnumber")),6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),7,6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),13,2)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),15)
'	rs.close
'else
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select OpenGOVReportnumber from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MAILCHKNUMBER=left(trim(rs("OpenGOVReportnumber")),6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),7,6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),13,2)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),15)
		rs.close
	else
		strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
		rs.close
	end if
'end if

If Not ifnull(request("Sys_LabelKind")) and instr(Sys_Note,"郵寄日")<=0 Then
	strSQL="select Note from BillBase where sn="&trim(rsbil("BillSN"))
	set rs=conn.execute(strSQL)
		strSQL="Update BillBase set Note='"&trim(rs("Note"))&" 郵寄日:"&gInitDT(date)&" 大宗:"&Sys_MAILCHKNUMBER&"' where sn="&trim(rsbil("BillSN"))
		conn.execute(strSQL)
		strSQL="Update BillMailHistory set StoreAndSendMailNumber=null,OpenGOVReportnumber=null where sn="&trim(rsbil("BillSN"))
		conn.execute(strSQL)
	rs.close
end if

rsbil.close
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;">

<div id="Layer1" style="position:absolute; left:60px; top:0px; z-index:5">
<%If not ifnull(request("Sys_UnitLabelKind")) Then
	response.write "請繳回："&Sys_BillUnitAddress&"　　"&Sys_BillUnitName&"<br>"
End if
Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_1.jpg"" hspace=""0"" vspace=""0"">"
Response.Write "<br><span class=""style7"">"&Sys_FirstBarCode&"</span>"
%>
</div>
<div id="Layer2" style="position:absolute; left:250px; top:20px; width:400px; height:36px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write Sys_DriverHomeZip&"<br>"
	Response.Write Sys_DriverZipName&funcCheckFont(Sys_DriverHomeAddress,20,1)&"<br><br>"
	Response.Write funcCheckFont(Sys_Driver,20,1)&"　台啟"
	Response.Write "</span>"%>
</div>
<div id="Layer3" style="position:absolute; left:320px; top:100px; z-index:5"><%
	Response.Write "<span class=""style4"">"
	Response.Write "大宗郵資已付掛號函件<br>　　　第"&Sys_MailNumber&"號"
	Response.Write "</span>"
	
%>
</div>

<div id="Layer3" style="position:absolute; left:60px; top:120px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write "二次郵寄<br>"
	Response.Write "</span>"
	Response.Write "<span class=""style3"">"
	Response.Write "行政文書"
	Response.Write "</span>"
	
%>
</div>

<div id="Layer4" style="position:absolute; left:300px; top:130px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""><br>"
    Response.Write Sys_MAILCHKNUMBER
	Response.Write "</span><br>"
%>
</div>

<div id="Layer5" style="position:absolute; left:60px; top:170px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write "應到案處所："&Sys_STATIONNAME&"<br>"
   	Response.Write "應到案處所電話："&Sys_StationTel
	Response.Write "</span>"

%>
</div>
<div id="Layer5" style="position:absolute; left:550px; top:150px; z-index:5"><%
	Response.Write "<span class=""style8"">"
	Response.Write Sys_StationID
	Response.Write "</span>"
%>
</div>

<div id="Layer45" style="position:absolute; left:80px; top:990px; width:400px; height:12px; z-index:36"><span class="style3"><%
	If not ifnull(request("Sys_UnitLabelKind")) Then
		response.write "請繳回："&Sys_BillUnitAddress&"　　"&Sys_BillUnitName
	else
		Response.Write Sys_BillNo
	end If 
%></span></div>

<div id="Layer44" style="position:absolute; left:520px; top:1000px; width:350px; height:12px; z-index:10"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>></div>

<div id="Layer45" style="position:absolute; left:100px; top:1030px; width:100px; height:12px; z-index:36"><span class="style3"><%
	Response.Write funcCheckFont(Sys_Driver,22,1)
%></span></div>

<div id="Layer42" style="position:absolute; left:100px; top:1050px; width:260px; height:40px; z-index:36; background-color:#FFFFFF"><span class="style3"><%
	Response.Write Sys_DriverHomeZip&" "&Sys_DriverZipName&funcCheckFont(Sys_DriverHomeAddress,22,1)
%></span></div>
</div>

<%
	response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>