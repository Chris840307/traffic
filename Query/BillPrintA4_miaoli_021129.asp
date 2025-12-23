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
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 12px}
.style2 {font-family:"標楷體"; font-size: 16px}
.style3 {font-family:"標楷體"; font-size: 14px}
.style7 {font-family:"標楷體"; font-size: 13px}
.style8 {font-family:"標楷體"; font-size: 13px; color:#0000FF;}
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
codebase="..\smsx.cab#Version=6,5,439,72">
</object>
<%
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BatchNumber,BillSN,BillNo,CarNo,ExchangeDate from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
CNum=""
strSQL="select cnt from (select RowNum cnt,BillSN from (select BillSN from DCILog where BatchNumber='"&trim(rsbil("BatchNumber"))&"' order by BillSN) order by BillSN) where BillSN="&PBillSN(i)

set dci=conn.execute(strSQL)
if not dci.eof then CNum=dci("cnt")
dci.close
Sys_BatChNumber=""
If not ifnull(request("Sys_BatchNumber")) Then
	Sys_BatChNumber=gInitDT(trim(rsbil("ExchangeDate")))&"_"&trim(rsbil("BatchNumber"))&"_"&(CNum)
End if

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_CarSimpleID="":Sys_CarAddID="":Sys_ProjectID="":Sys_Jurgeday=""
Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""
Sys_DriverBirth="":Sys_DriverID="":Sys_BillTypeID=""

if Not rs.eof then
	If trim(rs("BillTypeID")) = "1" Then

		Sys_DriverID=trim(rs("DriverID"))

	End if 
End if 

if Not rs.eof then Sys_CarAddID=trim(rs("CarAddiD"))
if Not rs.eof then Sys_CarSimpleID=trim(rs("CarSimpleID"))
if Not rs.eof then Sys_ProjectID=trim(rs("ProjectID"))
if Not rs.eof then Sys_BillUnitID=trim(rs("BillUnitID"))
if Not rs.eof then Sys_RecordMemberID=trim(rs("RecordMemberID"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
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
end if

if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close


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

Sys_Driver="":Sys_DriverHomeAddress="":Sys_DriverHomeZip=""

set rsfound=conn.execute(strSql)

If ifnull(Sys_OwnerAddress) Then

	chkaddress=""

	If Sys_BillTypeID = "1" Then
		
		if Not rsfound.eof then Sys_Driver=trim(rsfound("Driver"))
		if Not rsfound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))

	end If 

	if Not rsfound.eof then Sys_Owner=rsfound("Owner")

	If Not ifnull(trim(rsfound("OwnerAddress"))) Then
		If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") or instr(replace(rsfound("OwnerAddress"),"（","("),"(通") Then
			chkaddress=""
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
	end If 
	

	If ifnull(Sys_OwnerAddress) Then
		chkaddress="(車)"
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if

	If Sys_BillTypeID = "2" Then
		If not ifnull(Sys_OwnerAddress) Then

			strSQL="Update Billbase set Owner='"&rsfound("Owner")&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&chkaddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"

			conn.execute(strSQL)
		end If 
	End if 
end If 

If instr(Sys_OwnerAddress,"信箱") > 0 or instr(Sys_OwnerAddress,"信相") > 0 Then
	errBillNo=errBillNo&rsbil("BillNo")&","&Sys_OwnerAddress&"「為郵政地址請確認」\n"
End If 

If instr(Sys_OwnerAddress,"國外") > 0 Then
	errBillNo=errBillNo&rsbil("BillNo")&","&Sys_OwnerAddress&"「地址 包含 國外，請確認」\n"
End if

if not ifnull(Sys_OwnerAddress) then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"  "," ")," ","　")
	Sys_OwnerAddress=replace(Sys_OwnerAddress,"臺","台")
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerAddress) Then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
end If 

If not ifnull(Sys_DriverHomeZip) Then 
	strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof then Sys_DriverHomeZipName=trim(rszip("ZipName"))
	rszip.close
	
	Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,Sys_DriverHomeZipName,"")
end if

If Sys_BillTypeID="2" Then
	If len(trim(Sys_Owner))<3 or instr(Sys_Owner," ")>0 or instr(Sys_Owner,"　")>0 Then
		Sys_Owner=trim(replace(Sys_Owner," ","*"))
		errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"「姓名請確認」\n"
	end If 

	If instr(Sys_OwnerAddress,"?")>0 Then
		Sys_OwnerAddress=trim(replace(Sys_OwnerAddress,"?","*"))
		errBillNo=errBillNo&rsbil("BillNo")&","&Sys_OwnerAddress&"「地址請確認」\n"
	end If 
end if

Sys_Owner=trim(replace(""&Sys_Owner," ","*"))

Sys_DCIReturnStation=0
Sum_Level=0:Sys_Level1=0:Sys_Level2=0
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
Sys_Sex="":Sys_DriverBirth=""
strSql="select distinct BillFillerMemberID,BillMemID2,BillMemID3,BillMemID4,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME from BillBase where SN="&trim(rsbil("BillSN"))
set rssex=conn.execute(strSql)
if trim(Sys_BillTypeID)="1" then
	if Not rssex.eof then
		if trim(rssex("DriverSex"))="1" then
			Sys_Sex="男"
		else
			Sys_Sex="女"
		end if
	end If 
	
	if Not rssex.eof then
		Sys_DriverBirth=split(gArrDT(trim(rssex("DriverBirth"))),"-")
	else
		Sys_DriverBirth=split(gArrDT(trim("")),"-")
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

if Not rssex.eof then Sys_IMAGEFILENAME=trim(rssex("IMAGEFILENAME"))
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))

Sys_BillMemID2="null":Sys_BillMemID3="null":Sys_BillMemID4="null"

if Not rssex.eof then
	If not ifnull(rssex("BillMemID2")) Then Sys_BillMemID2=trim(rssex("BillMemID2"))
	If not ifnull(rssex("BillMemID3")) Then Sys_BillMemID3=trim(rssex("BillMemID3"))
	If not ifnull(rssex("BillMemID4")) Then Sys_BillMemID4=trim(rssex("BillMemID4"))
end if


Sys_BillMemName2="":Sys_BillJobName2=""
Sys_BillMemName3="":Sys_BillJobName3=""
Sys_BillMemName4="":Sys_BillJobName4=""

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_BillMemID2)

set Unit2=conn.execute(strSQL)
If not Unit2.eof Then
	Sys_BillMemName2=Unit2("chName")
	Sys_BillJobName2=Unit2("Content")
End if
Unit2.close

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_BillMemID3)
set Unit2=conn.execute(strSQL)
If not Unit2.eof Then
	Sys_BillMemName3=Unit2("chName")
	Sys_BillJobName3=Unit2("Content")
End if
Unit2.close

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_BillMemID4)
set Unit2=conn.execute(strSQL)
If not Unit2.eof Then
	Sys_BillMemName4=Unit2("chName")
	Sys_BillJobName4=Unit2("Content")
End if
Unit2.close


strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.Tel,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close


If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
unit.close

'strSQL="select UnitName,Tel from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
'set Unit=conn.execute(strSQL)
'SysUnit=Unit("UnitName")
'SysUnitTel=Unit("Tel")
'Unit.close

chkJobID=""

if trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
	chkJobID="303,304,305,307,314,318,1815,1838,1936,1937,1935,1938,1947,1948,1949" 

elseif trim(Sys_UnitLevelID)="1" then
	chkJobID="303,304,318,307,1947,1948,1949"
end If

Sys_jobName="":Sys_MainChName=""

strSQL="select a.ChName,b.Content,b.ID,b.showorder from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,showorder,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by b.showorder,b.id"
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

	Sys_CarSimpleName=""
	If cdbl(Sys_CarSimpleID)=1 or cdbl(Sys_CarSimpleID)=2 Then
		Sys_CarSimpleName="汽車"
	elseIf cdbl(Sys_CarSimpleID)=3 or cdbl(Sys_CarSimpleID)=4 Then
		Sys_CarSimpleName="機車"
	else
		Sys_CarSimpleName=""
	End If 

	Sys_IllegalRule1=""
	if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
		If not ifnull(Sys_CarSimpleName) Then
			strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and IllegalRule like '%"&Sys_CarSimpleName&"%' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				'Sys_Level1=trim(rsRule1("Level1"))
				Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		End if
		

		If ifnull(Sys_IllegalRule1) Then
			strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				'Sys_Level1=trim(rsRule1("Level1"))
				Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing	
		End if
		
	end if
	rssex.close

	Sys_IllegalRule2=""
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		If not ifnull(Sys_CarSimpleName) Then
			strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and IllegalRule like '%"&Sys_CarSimpleName&"%' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				'Sys_Level1=trim(rsRule1("Level1"))
				Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		End if
		
		If ifnull(Sys_IllegalRule2) Then
			strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				'Sys_Level1=trim(rsRule1("Level1"))
				Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing	
		End if
	end if

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close
strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
Sys_MailNumber=""
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
if trim(Sys_CarColorID(y))<>"" and not isnull(Sys_CarColorID(y)) then
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	if not rscolor.eof then
		Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	end if
	rscolor.close
end if
next

If ifnull(Sys_MailNumber) Then Sys_MailNumber=0

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,364,000,17
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",802,451,17
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,364,000,17
	
'	DelphiASPObj.CreateBarCode Sys_MailNumber&"80026336"
'	response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",802,451,17"
'	response.end
end if

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
Sys_MAILCHKNUMBER=Sys_MailNumber&"36400017"
rs.close
If Sys_OwnerZip="001" then Sys_OwnerZip=""
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer42" class="style2" style="position:absolute; left:60px; top:10px; z-index:5">
<%
	Response.Write "大宗郵資已付掛號函件<br>"
	Response.Write "　　第"&Sys_MailNumber&"號"

%>
</div>

<div id="Layer43" class="style2" style="position:absolute; left:270px; top:0px; z-index:5">
<%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""><br>"
	Response.Write "　"&Sys_MAILCHKNUMBER
%>
</div>

<div id="Layer44" class="style2" style="position:absolute; left:450px; top:70px; z-index:5">
<%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"
%>
</div>

<div id="Layer48" class="style2" style="position:absolute; left:100px; top:70px; z-index:5">
<%
	Response.Write Sys_BatChNumber	
%>
</div>

<div id="Layer45" class="style3" style="position:absolute; left:170px; top:115px; width:550px; height:13px; z-index:14">
<%
If Sys_BillTypeID = "1" Then
		
	'Sys_Owner="":Sys_OwnerZip="":Sys_OwnerZipName="":Sys_OwnerAddress=""

	Response.Write funcCheckFont(Sys_Driver,20,1)&"<br>"&Sys_DriverHomeZip&"　"&funcCheckFont(Sys_DriverHomeZipName&Sys_DriverHomeAddress,20,1)
else
	Response.Write funcCheckFont(Sys_Owner,20,1)&"<br>"&Sys_OwnerZip&"　"&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress&chkaddress,20,1)
end If 

%>
</div>

<div id="Layer46" class="style2" style="position:absolute; left:170px; top:155px; width:510px; height:13px; z-index:14">
<%=Sys_BillNo%>
</div>

<div id="Layer47" class="style2" style="position:absolute; left:170px; top:175px; width:510px; height:13px; z-index:14">
<%'="舉發違反道路交通管理事件通知單"%>
</div>

<div id="Layer43" class="style2" style="position:absolute; left:100px; top:510px; z-index:5">
<%
	If Sys_RecordMemberID = 3552 Then
		Response.Write "停管入案"
	End If 
%>
</div>


<div id="Layer48" class="style2" style="position:absolute; left:10px; top:565px; z-index:5">
<%
	Response.Write Sys_BatChNumber

	If Sys_RecordMemberID = 3552 Then
		Response.Write "<br><span class=""style10"">注意：尚有停車欠費未繳，補繳請至統一、全家、萊爾富、ok等超商多媒體事務機查詢補單繳納。</span>"
	End If 
%>
</div>

<!--<div id="Layer6" style="position:absolute; left:40px; top:540px; width:400px; height:36px; z-index:5"><span class="style7">查詢電話：<%=DB_UnitTel%><%'=Sys_UnitTel%>（<%=DB_UnitName%><%'=Sys_UnitName%>）</span></div>-->
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:10px; top:630px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:10px; top:665px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:130px; top:625px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:130px; top:640px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<div id="Layer5" style="position:absolute; left:165px; top:845px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer6" style="position:absolute; left:625px; top:610px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer7" style="position:absolute; left:625px; top:625px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer8" style="position:absolute; left:625px; top:640px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:0px; top:685px; width:202px; height:36px; z-index:5"><%
	If Sys_BillTypeID = "1" Then 
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		if showBarCode then
			response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
		end if
	end If 
%></div>
<div id="Layer10" style="position:absolute; left:510px; top:670px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer11" class="style2" style="position:absolute; left:90px; top:735px; width:300px; height:11px; z-index:10">
<%
If Sys_BillTypeID = "1" Then Response.Write "<br><font size=""4"">"&Sys_Driver&"</font>"
%>
</div>

<div id="Layer43" class="style2" style="position:absolute; left:70px; top:755px; z-index:5">
<%
	If Sys_RecordMemberID = 3552 Then
		Response.Write "██████"
	End If 
%>
</div>
<div id="Layer43" class="style8" style="position:absolute; left:70px; top:775px; z-index:5">
<%
	If Sys_RecordMemberID = 3552 Then
		Response.Write "停車收費逾期未繳案件"
	End If 
%>
</div>

<div id="Layer12" class="style7" style="position:absolute; left:170px; top:735px; width:300px; height:11px; z-index:10">
<%
If Sys_BillTypeID = "2" Then
	Response.Write funcCheckFont(Sys_A_Name,16,1)&"<br>"&Sys_CarColor
end If 
%>
</div>

<div id="Layer13" class="style2" style="position:absolute; left:220px; top:735px; width:300px; height:11px; z-index:10">
<%
If Sys_BillTypeID = "1" Then
	If Sys_Sex="男" Then
		Response.Write "v"
	else
		Response.Write "<br>v"
	End if 
end If 
%>
</div>

<div id="Layer14" class="style1" style="position:absolute; left:330px; top:740px; width:500px; height:10px; z-index:10">
<%
If Sys_BillTypeID = "1" Then
	Response.Write "<font size=""3"">"
	Response.Write Sys_DriverHomeZip&"　"&funcCheckFont(Sys_DriverHomeZipName&Sys_DriverHomeAddress,20,1)
	Response.Write "</font>"
else
	if showBarCode then response.write "*本單可至郵局或全國7-11、全家、萊爾富、OK等超商門市繳納。"
end if
%>
</div>

<div id="Layer15" class="style2" style="position:absolute; left:220px; top:758px; width:300px; height:11px; z-index:10">
<%
If Sys_BillTypeID = "1" Then
	Response.Write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"
end If 
%>
</div>

<div id="Layer16" class="style2" style="position:absolute; left:340px; top:755px; width:300px; height:11px; z-index:10">
<%
If Sys_BillTypeID = "1" Then
	Response.Write DriverID
end If 
%>
</div>

<div id="Layer17" class="style2" style="position:absolute; left:620px; top:750px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" class="style2" style="position:absolute; left:90px; top:790px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" class="style2" style="position:absolute; left:240px; top:790px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style2" style="position:absolute; left:530px; top:790px; width:201px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,20,1)%></div>
<div id="Layer21" class="style2" style="position:absolute; left:150px; top:825px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress&chkaddress,20,1)%></div>

<div id="Layer22" class="style2" style="position:absolute; left:80px; top:855px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style2" style="position:absolute; left:130px; top:855px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style2" style="position:absolute; left:180px; top:855px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style2" style="position:absolute; left:230px; top:855px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>

<div id="Layer26" class="style2" style="position:absolute; left:280px; top:855px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>

<div id="Layer27" class="style2" style="position:absolute; left:350px; top:850px; width:390px; height:31px; z-index:20"><%
	If trim(Sys_ProjectID) = "1" Then Response.Write "民眾檢舉案件"

	If ifnull(Sys_Jurgeday(0)) Then
		Response.Write "<br>"
	else
		Response.Write "，檢舉日期 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"
	end if

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "此路段限速"&Sys_RuleSpeed&"公里，駕駛人行車速度經測時速"&Sys_IllegalSpeed&"公里，超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
'			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
'				response.write "(滿100公里以上)"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
'				response.write "(80公里以上未滿100公里)"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
'				response.write "(60公里以上未滿80公里)"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
'				response.write "(40公里以上未滿60公里)"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
'				response.write "(20公里以上未滿40公里)"
'			else
'				response.write "(未滿20公里)"
'			end if
		end if
	else
		
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1
		if int(Sys_Rule1)=5620001 then	Sys_IllegalRule1=Sys_IllegalRule1&"(掛號催繳通知補繳逾7日期限)"
		If trim(Sys_Rule4)<>"" Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		'if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		'else
		'	response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		'end if
	end if
	if trim(Sys_Rule2)<>"" then
		if left(trim(Sys_Rule2),2)="40" or (int(Sys_Rule2)>4310200 and int(Sys_Rule2)<4310219) or (int(Sys_Rule2)>3310101 and int(Sys_Rule2)<3310111) then
			if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
				response.write "<br>此路段限速"&Sys_RuleSpeed&"公里，駕駛人行車速度經測時速"&Sys_IllegalSpeed&"公里，超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
	'			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
	'				response.write "(滿100公里以上)"
	'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
	'				response.write "(80公里以上未滿100公里)"
	'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
	'				response.write "(60公里以上未滿80公里)"
	'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
	'				response.write "(40公里以上未滿60公里)"
	'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
	'				response.write "(20公里以上未滿40公里)"
	'			else
	'				response.write "(未滿20公里)"
	'			end if
			end if
		else
			'smith edit for print two law 20070621
			if int(Sys_Rule2)=5620001 then	Sys_IllegalRule2=Sys_IllegalRule2&"(掛號催繳通知補繳逾7日期限)"
			if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
			'if len(Sys_IllegalRule2)<26 then
				response.write "<br>"&Sys_IllegalRule2
			'else
			'	response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
			'end if
		end if

	end If 
	
	if trim(Sys_CarAddID)="8" then response.write "(違規拖吊)"

	If Sys_RecordMemberID = 3552 Then Response.Write "苗栗市公所辦理路邊停車收費逾期未繳費案件(申訴電話037-359960)"

'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>

<div id="Layer28" class="style2" style="position:absolute; left:80px; top:875px; width:220px; height:15px; z-index:21">
<%
	Response.Write Sys_ILLEGALADDRESS
%>
</div>
<div id="Layer29" class="style2" style="position:absolute; left:100px; top:910px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style2" style="position:absolute; left:180px; top:910px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style2" style="position:absolute; left:250px; top:910px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style2" style="position:absolute; left:365px; top:965px; width:400px; height:49px; z-index:29"><%
	response.write left(trim(Sys_Rule1),2)&"　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1

	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
%></div>

<div id="Layer34" class="style2" style="position:absolute; left:360px; top:1000px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" class="style1" style="position:absolute; left:610px; top:1005px; width:300px; height:40px; z-index:28"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer35" style="position:absolute; left:390px; top:1055px; width:100px; z-index:29"><%
	if billprintuseimage=1 then
		response.write "<img src=""../billpng/"&Sys_UnitID&".png"" height=""55"">"
	end if%></div>

<div id="Layer37" style="position:absolute; left:570px; top:1055px; width:200px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" height=""45"">"
	else

		response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=0>"
		response.write "<tr>"

		Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName&"</span>　<span class=""style9"">"&Sys_ChName&"&nbsp;</span></td>"

		If not ifnull(Sys_BillMemName2) Then
			Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName2&"</span>&nbsp;<span class=""style9"">"&Sys_BillMemName2&"&nbsp;</span></td>"
		End if
		
		Response.Write "</tr>"

		If not ifnull(Sys_BillMemName3) Then

			response.write "<tr>"

			Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName3&"</span>&nbsp;<span class=""style9"">"&Sys_BillMemName3&"&nbsp;</span></td>"

			If not ifnull(Sys_BillMemName4) Then
				Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName4&"</span>&nbsp;<span class=""style9"">"&Sys_BillMemName4&"&nbsp;</span></td>"
			End if
			
			Response.Write "</tr>"
		
		end if

		response.write "</table>"
	end if
%></div>

<div id="Layer38" class="style2" style="position:absolute; left:90px; top:1090px; width:60px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" class="style2" style="position:absolute; left:160px; top:1090px; width:60px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" class="style2" style="position:absolute; left:210px; top:1090px; width:60px; z-index:34"><%=sys_Date(2)%></div>

</div>
<%
	if (i mod 100)=0 then response.flush
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
	printWindow(true,0,0,0,0);
</script>