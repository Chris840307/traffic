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
'Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BatchNumber,BillSN,BillNo,CarNo,ExchangeDate from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
CNum=""
strSQL="select cnt from (select RowNum cnt,BillSN from (select BillSN from DCILog where BatchNumber='"&trim(rsbil("BatchNumber"))&"' order by BillSN) order by BillSN) where BillSN="&PBillSN(i)

set dci=conn.execute(strSQL)
if not dci.eof then CNum=dci("cnt")
dci.close

Sys_BatChNumber=trim(rsbil("BatchNumber"))&"_"&(CNum)

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_CarSimpleID="":Sys_CarAddID="":Sys_ProjectID=""

if Not rs.eof then Temp_BillNo=trim(rs("BillNo"))

if Not rs.eof then Sys_CarAddID=trim(rs("CarAddiD"))
if Not rs.eof then Sys_CarSimpleID=trim(rs("CarSimpleID"))
if Not rs.eof then Sys_ProjectID=trim(rs("ProjectID"))
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
end If 

if Not rs.eof then
	sys_CaseInDate=split(gArrDT(trim(rs("CaseInDate"))),"-")
else
	sys_CaseInDate=split(gArrDT(trim("")),"-")
end If 

rs.close

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

strSql="select a.*,b.DriverHomeZip DriverZip,b.DriverHomeAddress DriverAddress from (select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W') a,(select CarNo,DriverHomeZip,DriverHomeAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A') b where a.carno=b.carno(+)"

set rsfound=conn.execute(strSql)

If ifnull(Sys_OwnerAddress) Then

	Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""

	if Not rsfound.eof then Sys_Owner=rsfound("Owner")

	chkaddress=""
	If Not ifnull(trim(rsfound("OwnerAddress"))) Then
		If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") Then
			chkaddress="(通)"
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
		chkaddress="(車)"
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if

	If not ifnull(Sys_OwnerAddress) Then
		strSQL="Update Billbase set Owner='"&rsfound("Owner")&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&chkaddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
		conn.execute(strSQL)
	end if
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
end if

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 or instr(Sys_Owner," ")>0 or instr(Sys_Owner,"　")>0 Then
		Sys_Owner=trim(replace(Sys_Owner," ","*"))
		errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"「姓名請確認」\n"
	end if
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
Sys_Sex=""
strSql="select distinct BillFillerMemberID,BillMemID2,BillMemID3,BillMemID4,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID from BillBase where SN="&trim(rsbil("BillSN"))
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

Sys_BillMemID2="null":Sys_BillMemID3="null":Sys_BillMemID4="null"

if Not rssex.eof then
	If not ifnull(rssex("BillMemID2")) Then Sys_BillMemID2=trim(rssex("BillMemID2"))
	If not ifnull(rssex("BillMemID3")) Then Sys_BillMemID2=trim(rssex("BillMemID3"))
	If not ifnull(rssex("BillMemID4")) Then Sys_BillMemID2=trim(rssex("BillMemID4"))
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
strSql="select MailNumber,MailTypeID,MailDate,SignDate,(select Content from DCICode where TypeID=7 and ID=BillMailHistory.SignResonID) BillStatus,OpenGovNumber from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
Sys_MailNumber=""
set rs=conn.execute(strSql)

if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_BillStatus=trim(rs("BillStatus"))
if Not rs.eof then Sys_SignDate=trim(rs("SignDate"))
if Not rs.eof then Sys_OpenGovNumber=trim(rs("OpenGovNumber"))
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

if trim(Sys_BillTypeID)="1" then
	If ifnull(Sys_MailNumber) Then Sys_MailNumber="0"
	Sys_MailNumber=right("00000000"&Sys_MailNumber,6)
	'DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,0,"360000",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	'DelphiASPObj.CreateBarCode Sys_MailNumber&"22007317",128,35,260
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	If ifnull(Sys_MailNumber) Then Sys_MailNumber=0	
	Sys_MailNumber=right("00000000"&Sys_MailNumber,6)
	'DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"36000017","360000",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	'DelphiASPObj.CreateBarCode Sys_MailNumber&"36000017",128,60,150

	'DelphiASPObj.CreateBarCode Sys_MailNumber&"22007317",128,60,150
end if

Sys_FirstBarCode=Sys_BillNo

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
Sys_MAILCHKNUMBER=Sys_MailNumber&"36000017"
rs.close
If Sys_OwnerZip="001" then Sys_OwnerZip=""
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="Layer000" style="position:absolute; left:30px; top:0px; z-index:1"><%
		Response.Write "<img src="".\img\BillBaseFastPaper_miaoli.jpg"" width=""620"" height=""520"">"
	%>
</div>
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:45px; top:45px; width:202px; height:36px; z-index:5">V</div>
<%else%>
<div id="Layer2" style="position:absolute; left:45px; top:75px; width:202px; height:36px; z-index:5">V</div>
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
<!--
<div id="Layer9" style="position:absolute; left:10px; top:100px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:440px; top:80px; width:233px; height:32px; z-index:5"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"""%>></div>
-->
<div id="Layer12" class="style2" style="position:absolute; left:520px; top:120px; width:300px; height:11px; z-index:5">
<%=Sys_BillNo%>
</div>
<div id="Layer12" class="style7" style="position:absolute; left:90px; top:150px; width:300px; height:11px; z-index:5">
<%=funcCheckFont(Sys_A_Name,16,1)%><br><%=Sys_CarColor%>
</div>

<div id="Layer14" class="style1" style="position:absolute; left:290px; top:150px; width:500px; height:10px; z-index:5">
<%'if showBarCode then response.write "*本單可至郵局或全國7-11、全家、萊爾富、OK等超商門市繳納。"%>
</div>

<div id="Layer17" class="style2" style="position:absolute; left:590px; top:155px; width:99px; height:12px; z-index:5"><%=fastring%></div>
<div id="Layer18" class="style2" style="position:absolute; left:95px; top:200px; width:100px; height:14px; z-index:5"><%=Sys_CarNo%></div>
<div id="Layer19" class="style2" style="position:absolute; left:220px; top:200px; width:117px; height:20px; z-index:5"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style2" style="position:absolute; left:410px; top:200px; width:201px; height:17px; z-index:5"><%=funcCheckFont(Sys_Owner,20,1)%></div>
<div id="Layer21" class="style2" style="position:absolute; left:100px; top:235px; width:507px; height:13px; z-index:5"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress&chkaddress,20,1)%></div>

<div id="Layer22" class="style2" style="position:absolute; left:95px; top:265px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style2" style="position:absolute; left:135px; top:265px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style2" style="position:absolute; left:175px; top:265px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style2" style="position:absolute; left:215px; top:265px; width:40px; height:16px; z-index:15"><%=right("00"&Sys_IllegalDate_h,2)%></div>

<div id="Layer26" class="style2" style="position:absolute; left:255px; top:265px; width:40px; height:13px; z-index:15"><%=right("00"&Sys_IllegalDate_m,2)%></div>

<div id="Layer27" class="style2" style="position:absolute; left:320px; top:265px; width:320px; height:31px; z-index:5"><%
	If trim(Sys_ProjectID) = "1" Then Response.Write "民眾檢舉案件"

	If not ifnull(Sys_Jurgeday(0)) Then
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
		Response.Write "<br>"
		if left(trim(Sys_Rule2),2)="40" or (int(Sys_Rule2)>4310200 and int(Sys_Rule2)<4310219) or (int(Sys_Rule2)>3310101 and int(Sys_Rule2)<3310111) then
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
			'smith edit for print two law 20070621
			if int(Sys_Rule2)=5620001 then	Sys_IllegalRule2=Sys_IllegalRule2&"(掛號催繳通知補繳逾7日期限)"
			if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
			'if len(Sys_IllegalRule2)<26 then
				response.write Sys_IllegalRule2
			'else
			'	response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
			'end if
		end if

	end If 
	
	if trim(Sys_CarAddID)="8" then response.write "(違規拖吊)"	

'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></div>

<div id="Layer28" class="style2" style="position:absolute; left:90px; top:285px; width:220px; height:15px; z-index:5">
<%
	Response.Write Sys_ILLEGALADDRESS
%>
</div>
<div id="Layer29" class="style2" style="position:absolute; left:110px; top:315px; width:34px; height:11px; z-index:5"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style2" style="position:absolute; left:175px; top:315px; width:35px; height:13px; z-index:5"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style2" style="position:absolute; left:240px; top:315px; width:32px; height:15px; z-index:5"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style2" style="position:absolute; left:325px; top:325px; width:400px; height:49px; z-index:15"><%
	response.write left(trim(Sys_Rule1),2)&"　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"&nbsp;&nbsp;&nbsp;"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　"&Sys_Level1

	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"&nbsp;&nbsp;&nbsp;"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　"&Sys_Level2
	end if
%></div>

<div id="Layer34" class="style2" style="position:absolute; left:310px; top:360px; width:400px; height:30px; z-index:5"><%
	'if showBarCode then	response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" class="style2" style="position:absolute; left:400px; top:365px; width:300px; height:40px; z-index:5"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer35" style="position:absolute; left:350px; top:425px; width:100px; z-index:5"><%
	if billprintuseimage=1 then
		response.write "<img src=""..\billpng\"&Sys_UnitID&".png"" height=""55"">"
	end if%></div>

<div id="Layer37" style="position:absolute; left:520px; top:435px; width:200px; z-index:5"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""..\Member\Picture\"&Sys_MemberFilename&""" height=""45"">"
	else

		response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=0>"
		response.write "<tr>"

		Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName&"</span>&nbsp;&nbsp;<span class=""style9"">"&Sys_ChName&"&nbsp;</span></td>"

		If not ifnull(Sys_BillMemName2) Then
			Response.Write "</tr><tr><td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName2&"</span>&nbsp;<span class=""style9"">"&Sys_BillMemName2&"&nbsp;</span></td>"
		End if
		
		Response.Write "</tr>"

		If not ifnull(Sys_BillMemName3) Then

			response.write "<tr>"

			Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName3&"</span>　<span class=""style9"">"&Sys_BillMemName3&"&nbsp;</span></td>"

			If not ifnull(Sys_BillMemName4) Then
				Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" height=""25"" align=""center"" nowrap><span class=""style9"">&nbsp;"&Sys_BillJobName4&"</span>&nbsp;<span class=""style9"">"&Sys_BillMemName4&"&nbsp;</span></td>"
			End if
			
			Response.Write "</tr>"
		
		end if

		response.write "</table>"
	end if
%></div>

<div id="Layer38" class="style2" style="position:absolute; left:180px; top:495px; width:60px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" class="style2" style="position:absolute; left:260px; top:495px; width:60px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" class="style2" style="position:absolute; left:350px; top:495px; width:60px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" class="style2" style="position:absolute; left:40px; top:520px; width:600px; z-index:34"><%
Response.Write "作業批號："&Sys_BatChNumber&"&nbsp;&nbsp;"

if trim(Sys_BillTypeID)="1" then
	Response.Write "舉發類別：攔停&nbsp;&nbsp;"

else
	Response.Write "舉發類別：逕舉&nbsp;&nbsp;"

end if

Response.Write "入案日："&sys_CaseInDate(0)&sys_CaseInDate(1)&sys_CaseInDate(2)&"&nbsp;&nbsp;"
Response.Write "投遞日："&gInitDT(Sys_MailDate)
Response.Write "<br>"
Response.Write "大宗掛號："&Sys_MailNumber&"&nbsp;&nbsp;"
Response.Write "簽收狀態："&Sys_BillStatus&"&nbsp;&nbsp;"
Response.Write "簽收日："&gInitDT(Sys_SignDate)&"&nbsp;&nbsp;"
Response.Write "公示文號："&Sys_OpenGovNumber
%>
</div>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<%

	strImgKS="select * from BILLILLEGALIMAGE where billsn="&PBillSN(i)
	set rsImgKS=conn.execute(strImgKS)
	if not rsImgKS.eof then
		if ifnull(rsImgKS("ImageFileNameB")) or trim(rsImgKS("ImageFileNameC"))<>"" then 
			Response.Write "<table border=0><tr><td>"
			response.write "<div id=""Layer45"" class=""style2"" z-index:34"">"
		
			if trim(rsImgKS("ImageFileNameA"))<>"" then
		%>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameA"))%>" name="imgB1" width="680" alt="">
		<%
			end if
			Response.Write "</Div>"
			Response.Write "</td></tr></table>"
		end if

	end if
	rsImgKS.close
	set rsImgKS=Nothing

	strImgKS="select * from BILLILLEGALIMAGE where billsn="&PBillSN(i)
	set rsImgKS=conn.execute(strImgKS)
	if not rsImgKS.eof then
		if not ifnull(rsImgKS("ImageFileNameB")) then 
			Response.Write "<table border=0><tr><td>"
			response.write "<div id=""Layer45"" class=""style2"" z-index:34"">"
		
			if trim(rsImgKS("ImageFileNameA"))<>"" then
		%>
				<div class="PageNext">　</div>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameA"))%>" name="imgB1" width="580" alt="">
		<%
			end If 

			if trim(rsImgKS("ImageFileNameB"))<>"" then
		%>
				<br><img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameB"))%>" name="imgB2" width="680" alt="">
		<%
			end if
			

			Response.Write "</Div>"
			Response.Write "</td></tr></table>"
		end If 
		
		if trim(rsImgKS("ImageFileNameC"))<>"" then
			Response.Write "<table border=0><tr><td>"
			response.write "<div id=""Layer45"" class=""style2"" z-index:34"">"
		%>
			<br><img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameC"))%>" name="imgB3" width="680" onclick="OpenPic('<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameC"))%>')">
		<%
			Response.Write "</Div>"
			Response.Write "</td></tr></table>"
		end if

	end if
	rsImgKS.close
	set rsImgKS=Nothing
			
		
	'送達証書
	strScan="select * from BillAttatchImage where BillNo='"&trim(Temp_BillNo)&"' and TypeID=0 and Recordstateid=0 order by RecordDate"
	set rsScan=conn.execute(strScan)
	while Not rsScan.eof
		Response.Write "<div class=""PageNext"">　</div>"
		Response.Write "<table border=0><tr><td>"
		response.write "<div id=""Layer45"" class=""style2"" z-index:34"">"
	%>
		<br><img src='<%=replace(trim(rsScan("FileName")),"/img/","/scanimg/")%>' name='imgB1' width='680'>
	<%
		Response.Write "</Div>"
		Response.Write "</td></tr></table>"
		rsScan.movenext
	wend
	rsScan.close
	set rsScan=Nothing
	if (i mod 100)=0 then response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
