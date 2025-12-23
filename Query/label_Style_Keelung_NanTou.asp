<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單信封黏貼標籤</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 9px}
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
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

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
Server.ScriptTimeout=6000
pagefile=0

if Instr(request("Sys_BatchNumber"),"N")>0 then
	KindType="('1','3','9','a','j','A','H','K','T','n')"

	tempSQL="where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.billno=i.billno and a.CarNo=i.CarNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.BillTypeID='2' and a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&request("sys_strSQL")&" and NVL(f.EquiPmentID,1)<>-1 and a.DciReturnStatusID<>'n'"
else
	KindType="('1','3','9','a','j','A','H','K','T')"

	tempSQL="where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.billno=i.billno and a.CarNo=i.CarNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.BillTypeID='2' and a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&request("sys_strSQL")&" and NVL(f.EquiPmentID,1)<>-1"
end if



if Instr(request("Sys_BatchNumber"),"N")>0 then
	strSQL="select distinct a.BillSN from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL

	strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
	chk_MailNumKind=1
else
	strSQL="select distinct a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL

	strSQL=strSQL&" order by f.RecordDate"
end if

set rssn=conn.execute(strSQL)
BillSN="":tempBillSN=""
while Not rssn.eof
	If trim(tempBillSN)<>trim(rssn("BillSN")) Then
		tempBillSN=trim(rssn("BillSN"))
		if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
		BillSN=BillSN&trim(rssn("BillSN"))
	end if
	rssn.movenext
wend
rssn.close

PBillSN=split(trim(BillSN),",")
for i=0 to Ubound(PBillSN)
if ((i+1) mod 20)=1 and i>1 then response.write "<div class=""PageNext""></div>"

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

Sys_DriverHomeZip=""
Sys_DriverZipName=""
Sys_DriverHomeAddress=""

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

Sys_DriverHomeAddress="":Sys_DriverHomeZip="":Sys_Driver=""


'if Instr(request("Sys_BatchNumber"),"N")>0 and Sys_BillTypeID=2 Then
'	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='N'"
'else
	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"
'end if

set rsfound=conn.execute(strSql)

Sys_Driver=trim(rsfound("Owner"))

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

	strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(PBillSN(i))&" and ExchangetypeID='A') and ExchangetypeID='A'"

	set rsdata=conn.execute(strsql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		If Sys_BillTypeID=2 Then
			if Not rsdata.eof then 
				Sys_DriverHomeAddress=trim(rsdata("DriverHomeAddress"))
				Sys_DriverHomeZip=trim(rsdata("DriverHomeZip"))
			end if
		end if
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

If not ifnull(Sys_DriverHomeAddress) Then
	Sys_DriverHomeAddress=replace(trim(Sys_DriverHomeAddress),"台","臺")
end if

Sys_Driver=""
strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then

	if Not rsfound.eof then Sys_Driver=trim(rsfound("Driver"))

	If ifnull(Sys_Driver) Then
		Sys_Driver=trim(rsfound("Owner"))
	end if

else
	if Not rsfound.eof then Sys_Driver=trim(rsfound("Owner"))
End if

If ifnull(Sys_DriverHomeAddress) Then
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))

		If ifnull(Sys_DriverHomeAddress) Then
			if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
		end if
	else
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
	End if
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close
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
end if

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

rsbil.close
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160

if (i mod 20)=0 then pagefile=0
pagefile=pagefile+1
	if (i mod 20)=0 or i=0 then Response.Write "<div id=""L78"" style=""position:relative;"">"
		if (pagefile mod 2)=1 then
			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:20px; "
			Response.Write "top:"& 5+(fix((pagefile-1)/2)*111) &"px; width:270px; height:5px; z-index:13"">"
			Response.Write "姓名："&funcCheckFont(Sys_Driver,16,1)&"</div>"
Sys_DriverZipName=replace(trim(Sys_DriverZipName),"台","臺")
Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)

			Response.Write "<div id=""Layer02"" class=""style3"" style=""position:absolute; left:20px; "
			Response.Write "top:"& 20+(fix((pagefile-1)/2)*111) &"px; width:340px; height:17px; z-index:13"">"
			Response.Write "地址："&Sys_DriverHomeZip&" "&funcCheckFont(Sys_DriverHomeAddress,16,1)
			Response.Write "</div>"

			Response.Write "<div id=""Layer03"" class=""style3"" style=""position:absolute; left:20px; "
			Response.Write "top:"& 50+(fix((pagefile-1)/2)*111) &"px; width:250px; height:36px; z-index:13"">"
			Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&".jpg"" height=""45"">"
			Response.Write "</div>"
		end if
		if (pagefile mod 2)=0 then
			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:400px; "
			Response.Write "top:"& 5+(fix((pagefile-1)/2)*111) &"px; width:270px; height:17px; z-index:13"">"
			Response.Write "姓名："&funcCheckFont(Sys_Driver,16,1)&"</div>"
Sys_DriverZipName=replace(trim(Sys_DriverZipName),"台","臺")
Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)
			Response.Write "<div id=""Layer02"" class=""style3"" style=""position:absolute; left:400px; "
			Response.Write "top:"& 20+(fix((pagefile-1)/2)*111) &"px; width:340px; height:17px; z-index:13"">"
			Response.Write "地址："&Sys_DriverHomeZip&" "&funcCheckFont(Sys_DriverHomeAddress,16,1)
			Response.Write "</div>"

			Response.Write "<div id=""Layer03"" class=""style3"" style=""position:absolute; left:400px; "
			Response.Write "top:"& 50+(fix((pagefile-1)/2)*111) &"px; width:250px; height:36px; z-index:13"">"
			Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&".jpg"" height=""45"">"
			Response.Write "</div>"
		end if
	if ((i+1) mod 20)=0 then Response.Write "</div>"
	if (i mod 100)=0 then response.flush
next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,0,0,0,0);
</script>