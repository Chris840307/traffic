<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-size: 10px}
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style12 {font-family:"標楷體"; font-size: 8px;}
.style13 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style14 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style15 {font-family:"標楷體"; font-size: 20px;}
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsxie8.cab#Version=6,5,439,50">
</object>
<%
'on Error Resume Next
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

Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""

Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

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

if not ifnull(Sys_OwnerAddress) then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"  "," ")," ","　")
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerAddress) Then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"臺","台"),Sys_OwnerZipName,"")
end if

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then
		Sys_Owner=trim(replace(Sys_Owner," ","*"))
		errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
	end if
end if

Sys_Owner=trim(replace(""&Sys_Owner," ","*"))

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cdbl(Sys_Level1)+Cdbl(Sys_Level2)
Sys_DCIRETURNCARTYPEID=""
if Not rsfound.eof then Sys_DCIRETURNCARTYPEID=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPEID&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex=""
strSql="select distinct BillMemID1,BillMemID2,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where SN="&PBillSN(i)
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
Sys_BillMemID1=0
if Not rssex.eof then Sys_BillMemID1=trim(rssex("BillMemID1"))
if Not rssex.eof then Sys_BillMemID2=trim(rssex("BillMemID2"))

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillMemID1
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

Sys_BillMemName2=""
Sys_BillJobName2=""
If ifnull(Sys_BillMemID2) Then Sys_BillMemID2=0

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_BillMemID2)
set Unit2=conn.execute(strSQL)
If not Unit2.eof Then
	Sys_BillMemName2=Unit2("chName")
	Sys_BillJobName2=Unit2("Content")
End if
Unit2.close

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_RecordMemberID)
set Unit=conn.execute(strSQL)
Sys_UnitID=Unit("UnitID")
Sys_RedUnitName=Unit("UnitName")
Sys_UnitTypeID=Unit("UnitTypeID")
Sys_UnitLevelID=Unit("UnitLevelID")
Sys_RecordName=Unit("chName")
Sys_RecordJobName=Unit("Content")
Unit.close

If Sys_UnitLevelID=1 or instr(Sys_RedUnitName,"分隊")>0 Then
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

'HHH	550cc重機
'HH	重型機車
'H	重機
'L	輕機
'TC	汽車試車
'TM	機車試車
'S	小型輕機
Sys_IllegalRule1=""

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then

	If Sys_DCIRETURNCARTYPEID = "HHH" or Sys_DCIRETURNCARTYPEID = "HH" Then
		strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and IllegalRule like '%大型%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing

		If Sys_IllegalRule1 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	End if 

	If Sys_DCIRETURNCARTYPEID = "H" or Sys_DCIRETURNCARTYPEID = "L" or Sys_DCIRETURNCARTYPEID = "TC" or Sys_DCIRETURNCARTYPEID = "TM" or Sys_DCIRETURNCARTYPEID = "S" Then
		strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and IllegalRule not like '%大型%' and IllegalRule like '%機%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing

		If Sys_IllegalRule1 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	End if

	If Sys_DCIRETURNCARTYPEID <> "HHH" and Sys_DCIRETURNCARTYPEID <> "HH" and Sys_DCIRETURNCARTYPEID <> "H" and Sys_DCIRETURNCARTYPEID <> "L" and Sys_DCIRETURNCARTYPEID <> "TC" and Sys_DCIRETURNCARTYPEID <> "TM" and Sys_DCIRETURNCARTYPEID <> "S" Then

		strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and IllegalRule like '%汽車%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing

		If Sys_IllegalRule1 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	end If 
end If 

rssex.close

Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

	If Sys_DCIRETURNCARTYPEID = "HHH" or Sys_DCIRETURNCARTYPEID = "HH" Then
		strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and IllegalRule like '%大型%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing

		If Sys_IllegalRule2 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	End if 

	If Sys_DCIRETURNCARTYPEID = "H" or Sys_DCIRETURNCARTYPEID = "L" or Sys_DCIRETURNCARTYPEID = "TC" or Sys_DCIRETURNCARTYPEID = "TM" or Sys_DCIRETURNCARTYPEID = "S" Then
		strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and IllegalRule not like '%大型%' and IllegalRule like '%機%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing

		If Sys_IllegalRule2 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	End If 

	If Sys_DCIRETURNCARTYPEID <> "HHH" and Sys_DCIRETURNCARTYPEID <> "HH" and Sys_DCIRETURNCARTYPEID <> "H" and Sys_DCIRETURNCARTYPEID <> "L" and Sys_DCIRETURNCARTYPEID <> "TC" and Sys_DCIRETURNCARTYPEID <> "TM" and Sys_DCIRETURNCARTYPEID <> "S" Then

		strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and IllegalRule like '%汽車%' and VerSion='"&Sys_RuleVer&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing
	

		If Sys_IllegalRule2 = "" Then

			strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		
		End if 
	end If 
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
end if

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
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,802,451,17
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,802,451,17

'	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,28,160,0

'	response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",1"
'	response.end
end if
strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close

Sys_ReportCaseNo=""
strSQL="select ReportCaseNo from billbaseTmp where billsn="&trim(rsbil("BillSN"))
set rs=conn.execute(strSQL)
if Not rs.eof then Sys_ReportCaseNo=trim(rs("ReportCaseNo"))
rs.close

rsbil.close

pageTop=0
pageLeft=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" style="position:relative;">

<div id="Layer44" class="style14" style="position:absolute; left:160px; top:0px; height:12px; z-index:36"><%=SysUnit&"送達證書"%></div>

<div id="Layer42" class="style14" style="position:absolute; left:295px; top:0px; width:500px; height:12px; z-index:36"><%="請繳回："&SysAddress%></div>

<div id="Layer41" class="style3" style="position:absolute; left:650px; top:0px; width:200px; height:12px; z-index:36"><%=Sys_MailNumber%></div>

<div id="Layer01" class="style3" style="position:absolute; left:120px; top:20px; z-index:3"><%
	'response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp; "&Sys_CarNo
	If chkIllegalDate Then Response.Write "(車主自取)"
	Response.Write "<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress
%>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:300px; top:60px; z-index:2"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:95px; top:50px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:390px; top:275px; z-index:1"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:120px; top:290px; z-index:1"><b><%
	if trim(Sys_BillTypeID)="1" then
		response.write Sys_Driver
	elseif trim(Sys_BillTypeID)="2" then
		response.write funcCheckFont(Sys_Owner,16,1)
	end If 
	If chkIllegalDate Then Response.Write "(車主自取)"
%>　台啟</b>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:100px; top:310px; width:370px; z-index:5"><b><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			if trim(Sys_BillTypeID)="1" then
				response.write Orz_DriverHomeZip&"　"
				response.write replace(Orz_DriverHomeZipName&Orz_DriverHomeAddress,Orz_DriverHomeZipName&Orz_DriverHomeZipName,Orz_DriverHomeZipName)&chkaddress
			elseif trim(Sys_BillTypeID)="2" then
				response.write Sys_OwnerZip&" "
				response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress
			end if
			response.write "<br><br>"
	%></b>
</div>

<div id="Layer45" class="style3" style="position:absolute; left:265px; top:330px; height:12px; z-index:1"><b><%=SysUnit%></b></div>

<div id="Layer07" class="style3" style="position:absolute; left:60px; top:370px; z-index:5"><%
	response.write Sys_BatchNumber&"　"&SysUnit&"　"&"("&cdbl(i+1)&"/"&cdbl(Ubound(PBillSN)+1)&")"
	Response.Write "<br>"
	Response.Write Sys_ReportCaseNo
%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:235px; top:345px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer43" style="position:absolute; left:320px; top:395px; width:250px; height:12px; z-index:36"><%=Sys_DCIRETURNCARTYPE%></div>

<div id="Layer08" class="style3" style="position:absolute; left:360px; top:370px; z-index:1"><%
	'Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>
<!-------------------------- 判斷 billillegalimage 有沒有這些圖檔 ------------------------------>
<!---------- todo 依據法條判斷, 如果是闖紅燈, 要把 a 檔的 xxxxx_a.jpg 換成 b檔的xxxxxx_b.jpg ---------------------------------------->

<div id="Layer46" style="position:absolute; left:45px; top:775px; width:202px; height:36px; z-index:5"><%
	Response.Write Sys_ReportCaseNo
%></div>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:45px; top:810px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:45px; top:845px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:170px; top:810px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:830px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<div id="Layer9" style="position:absolute; left:40px; top:865px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:860px; width:233px; height:32px; z-index:3"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer12" style="position:absolute; left:110px; top:920px; width:250px; height:11px; z-index:6"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>

<div id="Layer13" style="position:absolute; left:260px; top:915px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:370px; top:915px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>

<div id="Layer15" style="position:absolute; left:260px; top:925px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:425px; top:925px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:620px; top:925px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:125px; top:965px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:270px; top:965px; width:250px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:510px; top:965px; width:300px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" style="position:absolute; left:125px; top:990px; width:610px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress%></div>

<div id="Layer22" style="position:absolute; left:115px; top:1010px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:165px; top:1010px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:215px; top:1010px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:265px; top:1010px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:315px; top:1010px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:390px; top:1015px; width:350px; height:31px; z-index:20"><span class="style3"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>4310209 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
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
	response.write "</font>"
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		response.write "<br>"&Sys_IllegalRule2
	end if
	if trim(Sys_Rule4)<>"" then response.write "("&Sys_Rule4&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:115px; top:1030px; width:220px; height:15px; z-index:21"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:120px; top:1065px; width:50px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:200px; top:1065px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:280px; top:1065px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" style="position:absolute; left:405px; top:1080px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	'if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		'if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:390px; top:1115px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" style="position:absolute; left:635px; top:1115px; width:100px; height:40px; z-index:28"><span class="style7"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>

<div id="Layer36" class="style7" style="position:absolute; left:660px; top:1190px; width:140px; height:43px; z-index:30"><%=Sys_ChName%></div>

<div id="Layer38" style="position:absolute; left:210px; top:1250px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:365px; top:1250px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:515px; top:1250px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:690px; top:1250px; width:80px; height:12px; z-index:36"><%%></div>

</div>

<%
		if (i mod 10)=0 then response.flush
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
	//window.print();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>