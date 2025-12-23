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
.style3 {font-size: 16px}
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
' jafe 20111202 by  and f.recordstateid=0,刪除的不出現
If instr(Request("Sys_Batchnumber"),"W") Then
	strSQL="select distinct a.BillSN,f.RecordDate from DCILog a,BillBase f,(select distinct BillNo,CarNo from BillBaseDCIReturn where EXCHANGETYPEID='W') i where a.DciReturnStatusID not in('N') and a.BillSN=f.sn and a.BillNo=i.BillNo and a.CarNo=i.CarNo and f.recordstateid=0 "&request("sys_strSQL")&" order by RecordDate"
else
	strSQL="select distinct a.BillSN from DCILog a,BillBase f,(select distinct BillNo,CarNo from BillBaseDCIReturn where EXCHANGETYPEID='W') i where a.DciReturnStatusID not in('n','N') and a.BillSN=f.sn and a.BillNo=i.BillNo and a.CarNo=i.CarNo  and f.recordstateid=0 "&request("sys_strSQL")

	strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
End if

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

Sys_DriverHomeAddress="":Sys_Driver=""
strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

Sys_Level1=0:Sys_Level2=0

if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(funTnumber(Sys_Level1))+cdbl(funTnumber(Sys_Level2))

If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Driver=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
	
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
		If ifnull(Sys_DriverHomeAddress) Then
			if Not rsFound.eof then Sys_Driver=trim(rsfound("Owner"))
			if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
		end if
	else
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_DriverHomeAddress) Then
	strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
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

		If ifnull(Sys_DriverHomeAddress) Then
			if Not rsFound.eof then Sys_Driver=trim(rsfound("Owner"))
			if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
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

If ifnull(Sys_DriverHomeAddress) Then
	if Not rsfound.eof then Sys_Driver=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
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
elseif left(trim(Sys_Rule1),3)="294" or left(trim(Sys_Rule1),3)="295" or left(trim(Sys_Rule1),3)="303" or left(trim(Sys_Rule1),3)="314" or left(trim(Sys_Rule1),3)="362" or left(trim(Sys_Rule1),3)="363" or left(trim(Sys_Rule1),3)="621" or left(trim(Sys_Rule1),3)="624" or left(trim(Sys_Rule1),3)="625" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="13" or left(trim(Sys_Rule1),2)="18" or left(trim(Sys_Rule1),2)="20" or left(trim(Sys_Rule1),2)="21" or left(trim(Sys_Rule1),2)="23" or left(trim(Sys_Rule1),2)="24" or left(trim(Sys_Rule1),2)="26" or left(trim(Sys_Rule1),2)="27" or left(trim(Sys_Rule1),2)="35" or left(trim(Sys_Rule1),2)="37" or left(trim(Sys_Rule1),2)="43" or left(trim(Sys_Rule1),2)="54" or left(trim(Sys_Rule1),2)="61" then
	showBarCode=false
elseif int(Sys_Rule1)=29300012 or int(Sys_Rule1)=29300022 or int(Sys_Rule1)=3400003 or int(Sys_Rule1)=3400004 then
	showBarCode=false
end if
if (i mod 20)=0 then pagefile=0
pagefile=pagefile+1
	if (i mod 20)=0 or i=0 then Response.Write "<div id=""L78"" style=""position:relative;"">"
		if (pagefile mod 2)=1 then
			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:70px; "
			Response.Write "top:"& 20+(fix((pagefile-1)/2)*110) &"px; height:17px; z-index:13"">"
			Response.Write "單號："&Sys_BillNo
			'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then Response.Write "(罰鍰："&Sum_Level&"元)"
			Response.Write "</div>"

			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:70px; "
			Response.Write "top:"& 35+(fix((pagefile-1)/2)*110) &"px; width:280px; height:17px; z-index:13"">"
			Response.Write "姓名："&funcCheckFont(Sys_Driver,16,1)&"</div>"

Sys_DriverZipName=replace(trim(Sys_DriverZipName),"台","臺")
Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)

			Response.Write "<div id=""Layer02"" class=""style3"" style=""position:absolute; left:70px; "
			Response.Write "top:"& 50+(fix((pagefile-1)/2)*110) &"px; width:280px; height:17px; z-index:13"">"
			Response.Write "地址："&Sys_DriverHomeZip&" "&funcCheckFont(Sys_DriverHomeAddress,16,1)
			Response.Write "</div>"
		end if
		if (pagefile mod 2)=0 then
			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:450px; "
			Response.Write "top:"& 20+(fix((pagefile-1)/2)*110) &"px; height:17px; z-index:13"">"
			Response.Write "單號："&Sys_BillNo
			'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then Response.Write "(罰鍰："&Sum_Level&"元)"
			Response.Write "</div>"

			Response.Write "<div id=""Layer01"" class=""style3"" style=""position:absolute; left:450px; "
			Response.Write "top:"& 35+(fix((pagefile-1)/2)*110) &"px; width:280px; height:17px; z-index:13"">"
			Response.Write "姓名："&funcCheckFont(Sys_Driver,16,1)&"</div>"

Sys_DriverZipName=replace(trim(Sys_DriverZipName),"台","臺")
Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)

			Response.Write "<div id=""Layer02"" class=""style3"" style=""position:absolute; left:450px; "
			Response.Write "top:"& 50+(fix((pagefile-1)/2)*110) &"px; width:280px; height:17px; z-index:13"">"
			Response.Write "地址："&Sys_DriverHomeZip&" "&funcCheckFont(Sys_DriverHomeAddress,16,1)
			Response.Write "</div>"
		end if
	if ((i+1) mod 20)=0 then Response.Write "</div>"
	if (i mod 5)=0 then response.flush
next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,0,0,0,0);
</script>