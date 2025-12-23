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
.style1 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style2 {font-family:"標楷體"; font-size: 10px}
.style3 {font-family:"標楷體"; font-size: 14px}
.style4 {font-family:"標楷體"; font-size: 18px}
.style5 {font-family:"標楷體"; font-size: 12px}
.style7 {font-family:"標楷體"; font-size: 13px}
.style8 {font-family:"標楷體"; font-size: 36px}
.style11 {font-family:"標楷體"; font-size: 14px}
.style15 {font-family:"標楷體"; font-size: 15px}
-->
</style>
</head>

<body>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
Server.ScriptTimeout=6000
'On Error Resume Next
PBillSN=split(trim(request("PBillSN")),",")

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

strCity="select value from Apconfigure where id=30"
set rsCity=conn.execute(strCity)
sys_title=trim(rsCity("value"))
rsCity.close

for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"
set rsbil=conn.execute(strBil)

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
'===初始化(8/21)==
Sys_BillNo=""
Sys_CarNo=""
Sys_DriverHomeZip=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_Driver=""
Sys_Owner=""
'================
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""
Orz_DriverHomeZip="":Orz_DriverHomeAddress="":Sys_DriverID=""

if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Orz_DriverHomeZip=trim(rs("OwnerZip"))
if Not rs.eof then Orz_DriverHomeAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
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

If ifnull(Sys_OwnerAddress) Then
	strSQL="select OwnerNotIfyAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
	
	set rsfi=conn.execute(strSql)

	if Not rsfi.eof then
		If Not ifnull(trim(rsfi("OwnerNotIfyAddress"))) Then

			notify_Zip="":notify_Addr=""

			if IsNumeric(left(trim(rsfi("OwnerNotIfyAddress")),3)) then

				notify_Zip=left(trim(rsfi("OwnerNotIfyAddress")),3)
			end If 

			
			notify_Addr=trim(rsfi("OwnerNotIfyAddress"))
			
			If not ifnull(notify_Zip) Then
				
				notify_Addr=mid(trim(rsfi("OwnerNotIfyAddress")),4)
			End if 

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
	If not ifnull(Sys_DriverID) Then

		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

		Orz_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		Orz_DriverHomeZip=trim(rsfound("DriverHomeZip"))
	End if 

	If ifnull(Sys_OwnerAddress) Then

		if Not rsfound.eof then Sys_Owner=rsfound("Owner")

		chkaddress=""
		If Not ifnull(trim(rsfound("OwnerAddress"))) Then
			If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") or instr(replace(rsfound("OwnerAddress"),"（","("),"(通") Then

				Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
				Sys_OwnerZip=trim(rsfound("OwnerZip"))

				Orz_DriverHomeAddress=trim(rsfound("OwnerAddress"))
				Orz_DriverHomeZip=trim(rsfound("OwnerZip"))
			End if

			If ifnull(Sys_OwnerAddress) Then
				chkaddress="(戶)"
				Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

				Orz_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
				Orz_DriverHomeZip=trim(rsfound("DriverHomeZip"))
			End if

		else
			If ifnull(Sys_OwnerAddress) Then
				chkaddress="(戶)"

				Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

				Orz_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
				Orz_DriverHomeZip=trim(rsfound("DriverHomeZip"))
			End if
		end If 
	End if 

	If ifnull(Sys_OwnerAddress) Then
		chkaddress=""
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))

		Orz_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		Orz_DriverHomeZip=trim(rsfound("OwnerZip"))
	End If 

	If not ifnull(Sys_OwnerAddress) Then
		strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
		conn.execute(strSQL)
	end if

end If 

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&"",replace(Sys_OwnerZipName&"","台","臺"),"")

If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Orz_Driver=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Orz_Driver=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Orz_DriverID=trim(rsfound("DriverID"))


strSQL="select ZipName from Zip where ZipID='"&Orz_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Orz_DriverHomeZipName=trim(rszip("ZipName"))
rszip.close

if not ifnull(Orz_DriverHomeZipName) then Orz_DriverHomeAddress=replace(Orz_DriverHomeAddress&"",replace(Orz_DriverHomeZipName&"","台","臺"),"")


'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
end if

Sys_DCIReturnStation=0
Sum_Level=0:Sys_Level1=0:Sys_Level2=0:point=""
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
if Not rsfound.eof then point=replace(replace(replace(replace(trim(rsfound("POINT"))&" ","A","10"),"B","11"),"C","12"),"D","13")
point=trim(point)
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
Sys_DCIRETURNCARTYPEID=""
if Not rsfound.eof then Sys_DCIRETURNCARTYPEID=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPEID&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex=""
strSql="select distinct BillFillerMemberID,BillMemID2,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID from BillBase where SN="&trim(rsbil("BillSN"))
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
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))
if Not rssex.eof then Sys_BillFillerMemberID2=trim(rssex("BillMemID2"))
' 讀取違規影像，由Kevin的影像建檔處寫入
strSQL="select * from BillIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
Sys_IisImagePath="":Sys_ImageFileNameA="":Sys_ImageFileNameB=""
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	Sys_ImageFileNameA=trim(rsimage("ImageFileNameA"))
	Sys_ImageFileNameB=trim(rsimage("ImageFileNameB"))
end if

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
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
If Not unit.eof Then
	SysUnit=unit("UnitName")
	SysUnitTel=trim(unit("Tel"))
	SysUnitAddress=trim(unit("Address"))
end if
unit.close


If Sys_RecordMemberID = 6227 or Sys_RecordMemberID = 6607 Then
	sys_title="宜蘭縣政府"
	SysUnit="交通處"
	SysUnitTel="(03)9251000"
	SysUnitAddress="26060 宜蘭市縣政北路1號"
end If 

chkJobID=""

if trim(Session("Unit_ID"))="TP00" Then
	'chkJobID="304,360"
	chkJobID="303,360"

elseif trim(Session("Unit_ID"))="TN00" then
	chkJobID="303,360,361"

elseif trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
	'if trim(Session("Unit_ID"))="TP00" or trim(Session("Unit_ID"))="TO00" or trim(Session("Unit_ID"))="TM00" then	'羅東、礁溪、宜蘭抓所長
		chkJobID="303,304,359,360,361"
	'else
	'	chkJobID="303,304,314"
	'end if
else
	'chkJobID="304"
	chkJobID="99999"
end if

Sys_jobName="":Sys_MainChName=""

strSQL="select a.ChName,b.Content,b.ID,b.showorder from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content,showorder from Code where TypeID=4 ) b where a.JobID=b.ID order by showorder,ID"
'response.write strSQL
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

if trim(Sys_UnitLevelID)="1" and Sys_jobName="" then
	Sys_jobName="分隊長"
	Sys_MainChName="張容苡"
end if

strSQL="select Value from ApConfigUre where ID=40"
set City=conn.execute(strSQL)
Sys_City=City("Value")
City.close

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
end If 

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
if trim(Sys_CarColorID(y))<>"" and not isnull(Sys_CarColorID(y)) then
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	if not rscolor.eof then
		Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	end if
	rscolor.close
end if
next

strSQL="update billmailhistory set mailchknumber='"&Sys_MailNumber&" 950000 17' where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
conn.execute(strSQL)

if trim(Sys_BillTypeID)="1" then

	DelphiASPObj.GenBillPrintBarCode1 PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95000017","220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode1 PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95000017","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if

if Sys_DriverHomeZip="001" then Sys_DriverHomeZip=""
if Sys_OwnerZip="001" then Sys_OwnerZip=""

Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo&"-"&Sys_CarNo

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close

firstBacrCode=right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&"D"&Sys_StationID

pageleft=0
pagetop=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer01" class="style3" style="position:absolute; left:<%=75+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write sys_title&SysUnit
%>
</div>
<div id="Layer66" class="style3" style="position:absolute; left:<%=345+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write SysUnitAddress&"(郵戳請勿蓋在條碼上)"
%>
</div>
<div id="Layer000" style="position:absolute; left:30px; top:0px; z-index:1"><%

'	If Sys_RecordMemberID = 6227 Then
'
'		If trim(Request("Sys_UnitLabelKind")) = "2" Then
'
'			strURL="il01_city.jpg"
'		else
'
'			strURL="il01_legal_city.jpg"
'		End if 
'
'	else
'
'		If trim(Request("Sys_UnitLabelKind")) = "2" Then
'
'			strURL="il01.jpg"
'		else
'
'			strURL="il01_legal.jpg"
'		End if 
'	
'	End if 
	strURL="il03_legal_city.jpg"
	Response.Write "<img src=""..\legal_Img\"&strURL&""" width=""715"" height=""1290"">"
	
	%>
</div>
<div id="Layer01" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=15+pagetop%>px; z-index:10"><%
	response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
%>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=60+pagetop%>px; z-index:11"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:<%=180+pageleft%>px; top:<%=50+pagetop%>px; z-index:10"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"" width=""180"" height=""25"">"%>
</div>
<!--
<div id="Layer05" class="style3" style="position:absolute; left:<%=470+pageleft%>px; top:<%=287+pagetop%>px; z-index:10">
　　<b>第<%=Sys_MailNumber%>號</b><br>
　　<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>><br>
　<b><%="　"&Sys_MAILCHKNUMBER%></b>
</div>
-->
<div id="Layer03" class="style3" style="position:absolute; left:<%=405+pageleft%>px; top:<%=270+pagetop%>px; z-index:10"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:<%=110+pageleft%>px; top:<%=290+pagetop%>px; z-index:10"><%
	if trim(Sys_BillTypeID)="1" then
		response.write Sys_Driver
	elseif trim(Sys_BillTypeID)="2" then
		response.write funcCheckFont(Sys_Owner,16,1)
	end if%>　台啟&nbsp&nbsp&nbsp&nbsp
	<%
	Response.Write "<span class=""style4""><B>郵遞區號："
	if trim(Sys_BillTypeID)="1" then
		response.write Orz_DriverHomeZip
	elseif trim(Sys_BillTypeID)="2" then
		response.write Sys_OwnerZip
	end If 
	Response.Write "</B></span>"
	%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:<%=340+pageleft%>px; top:<%=330+pagetop%>px; z-index:10"><%
	Response.Write "<img  src=""..\BarCodeImage\"&Sys_BillNo&"_1.jpg"">"%>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:<%=100+pageleft%>px; top:<%=310+pagetop%>px; width:340px; z-index:10"><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			if trim(Sys_BillTypeID)="1" then
				'response.write Orz_DriverHomeZip&"　"
				response.write replace(Orz_DriverHomeZipName&Orz_DriverHomeAddress,Orz_DriverHomeZipName&Orz_DriverHomeZipName,Orz_DriverHomeZipName)
			elseif trim(Sys_BillTypeID)="2" then
				'response.write Sys_OwnerZip&" "
				response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
			end if
			response.write "<br><br>"
	%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:<%=110+pageleft%>px; top:<%=330+pagetop%>px; z-index:10"><%
	Response.Write sys_title&SysUnit
%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:<%=70+pageleft%>px; top:<%=360+pagetop%>px; z-index:10"><%
	'Response.Write Sys_CarNo&"　"&Sys_A_Name&"　"&Sys_CarColor&"　"&Sys_STATIONNAME
	Response.Write Sys_CarNo&"　"&Sys_STATIONNAME
%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:<%=380+pageleft%>px; top:<%=360+pagetop%>px; z-index:10"><%
	Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>
<!-------------------------- 判斷 billillegalimage 有沒有這些圖檔 ------------------------------>
<!---------- todo 依據法條判斷, 如果是闖紅燈, 要把 a 檔的 xxxxx_a.jpg 換成 b檔的xxxxxx_b.jpg ---------------------------------------->
<%if trim(Sys_ImageFileNameA)<>"" then%>
	<div id="Layer09" style="position:absolute; left:38px; top:480px; z-index:5"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	%></DIV>
<%end if%>
<!--
<div id="Layer09" style="position:absolute; left:<%=38+pageleft%>px; top:<%=485+pagetop%>px; z-index:10"><%
	response.write "<img src=""d:\0001.jpg"" width=""390"" height=""280"">"
%></DIV>
-->
<%
'response.write Sys_Rule1 & "_" 
'response.write Sys_ImageFileNameB
' ssmith 20091015 紅燈月線不顯示B圖
%>
<%if trim(Sys_ImageFileNameB)<>"" then%>
	<div id="Layer10" style="position:absolute; left:430px; top:480px; z-index:1"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:<%=50+pageleft%>px; top:<%=810+pagetop%>px; width:202px; height:36px; z-index:10">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:<%=50+pageleft%>px; top:<%=845+pagetop%>px; width:202px; height:36px; z-index:10">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:<%=170+pageleft%>px; top:<%=815+pagetop%>px; width:202px; height:36px; z-index:10">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:<%=170+pageleft%>px; top:<%=830+pagetop%>px; width:202px; height:36px; z-index:10">v</div>
<%end if%>

<div id="Layer9" class="style3" style="position:absolute; left:<%=45+pageleft%>px; top:<%=865+pagetop%>px; z-index:10"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"" width=""200"" height=""35"">"
	end if
%></div>

<div id="Layer26" class="style3" style="position:absolute; left:<%=115+pageleft%>px; top:<%=897+pagetop%>px; z-index:10"><%
	if showBarCode then
		response.write firstBacrCode
	end if
%></div>

<div id="Layer10" style="position:absolute; left:<%=510+pageleft%>px; top:<%=845+pagetop%>px; z-index:10"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"" width=""200"" height=""30"""%>></div>

<div id="Layer11" class="style3" style="position:absolute; left:<%=540+pageleft%>px; top:<%=890+pagetop%>px; z-index:10"><%="宜警交"&"　　　　"&Sys_BillNo%></div>

<div id="Layer12" class="style7" style="position:absolute; left:<%=110+pageleft%>px; top:<%=920+pagetop%>px; width:150px; height:11px; z-index:10"><%
		If not ifnull(Sys_Jurgeday(0)) Then			
			response.write "民眾檢舉舉發&nbsp;"&Sys_A_Name&"<br>"
		else
			response.write "逕行舉發&nbsp;"&Sys_A_Name&"<br>"
		end If 

		if int(Sys_Rule1)<>4340003 and int(Sys_Rule1)<>5620001 and int(Sys_Rule1)<>5630001 then response.write "附採證照片"

		response.write "&nbsp;"&Sys_CarColor
%>
</div>

<div id="Layer13" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=915+pagetop%>px; width:28px; height:11px; z-index:10"><%=Sys_Sex%></div>
<div id="Layer14" class="style3" style="position:absolute; left:<%=370+pageleft%>px; top:<%=915+pagetop%>px; width:324px; height:10px; z-index:10"><% if showBarCode then Response.Write "*本單可至郵局或委託代收之超商繳納"%></div>

<div id="Layer15" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=925+pagetop%>px; width:100px; height:10px; z-index:10"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></div>
<div id="Layer16" class="style3" style="position:absolute; left:<%=435+pageleft%>px; top:<%=935+pagetop%>px; width:106px; height:13px; z-index:10"><%=Sys_DriverID%></div>
<div id="Layer17" class="style3" style="position:absolute; left:<%=620+pageleft%>px; top:<%=925+pagetop%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=965+pagetop%>px; width:100px; height:14px; z-index:10"><%=Sys_CarNo%></div>
<div id="Layer19" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=965+pagetop%>px; width:130px; height:20px; z-index:10"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style3" style="position:absolute; left:<%=500+pageleft%>px; top:<%=965+pagetop%>px; width:300px; height:17px; z-index:10"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=990+pagetop%>px; width:507px; height:13px; z-index:10"><%=Orz_DriverHomeZip&" "&funcCheckFont(Orz_DriverHomeZipName&Orz_DriverHomeAddress,16,1)%></div>

<div id="Layer22" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:13px; z-index:10"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style3" style="position:absolute; left:<%=170+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:17px; z-index:10"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style3" style="position:absolute; left:<%=220+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:16px; z-index:10"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style3" style="position:absolute; left:<%=270+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:16px; z-index:10"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" class="style3" style="position:absolute; left:<%=320+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:13px; z-index:10"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style5" style="position:absolute; left:<%=390+pageleft%>px; top:<%=1015+pagetop%>px; width:340px; height:31px; z-index:10"><%
	
	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if (left(trim(Sys_Rule1),1)="3" or left(trim(Sys_Rule1),1)="4") and trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then

		response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
	else

		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"

		response.write Sys_IllegalRule1
	end if

	if trim(Sys_Rule4)<>"" then response.write "("&Sys_Rule4&")"

	if trim(Sys_Rule2)<>"" and trim(Sys_Rule2)>"0" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"

		response.write "<br>"&Sys_IllegalRule2

	end if

	if instr(Sys_BillNo,"QZ")>0 then response.write Sys_Note 

	
	If left(trim(Sys_Rule1),2)<>"55" and left(trim(Sys_Rule1),2)<>"56" and left(trim(Sys_Rule2),2)<>"55" and left(trim(Sys_Rule2),2)<>"56" Then
		
		if trim(point)<>"0" and trim(point)<>"" then response.write "(記"&point&"點)"
	End if 

%></div>
<div id="Layer28" class="style3" style="position:absolute; left:<%=115+pageleft%>px; top:<%=1030+pagetop%>px; width:220px; height:15px; z-index:10"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=1065+pagetop%>px; width:50px; height:11px; z-index:10"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style3" style="position:absolute; left:<%=210+pageleft%>px; top:<%=1065+pagetop%>px; width:35px; height:13px; z-index:10"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style3" style="position:absolute; left:<%=280+pageleft%>px; top:<%=1065+pagetop%>px; width:32px; height:15px; z-index:10"><%=Sys_DealLineDate(2)%></div>

<div id="Layer32" class="style5" style="position:absolute; left:<%=420+pageleft%>px; top:<%=1100+pagetop%>px; z-index:10"><%
	response.write left(trim(Sys_Rule1),2)
	'if len(trim(Sys_Rule1))>7 then response.write "-"&right(trim(Sys_Rule1),1)
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)
		'if len(trim(Sys_Rule2))>7 then response.write "-"&right(trim(Sys_Rule2),1)
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　"&Sys_Level2
	end if

%></div>

<div id="Layer34" class="style3" style="position:absolute; left:<%=390+pageleft%>px; top:<%=1135+pagetop%>px; z-index:10"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"" width=""200"" height=""25"">"
%></div>

<div id="Layer33" class="style7" style="position:absolute; left:<%=600+pageleft%>px; top:<%=1135+pagetop%>px; z-index:11"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer35" style="position:absolute; left:<%=400+pageleft%>px; top:<%=1175+pagetop%>px; width:100px; height:49px; z-index:10"><%
	if instr(Sys_BillNo,"QZ")>0 then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"">宜蘭縣政府警察局<br>交通隊停管場</td></tr>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"">TEL(03)9255919</td></tr>"
		response.write "</table>"

	'elseif trim(Session("Unit_ID"))="TN00" then

	else
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"" nowrap>宜蘭縣政府警察局<br>"&SysUnit&"</td></tr>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"">TEL"&SysUnitTel&"</td></tr>"
		response.write "</table>"
	end if%></div>
<div id="Layer36" style="position:absolute; left:<%=610+pageleft%>px; top:<%=1210+pagetop%>px; width:100px; height:43px; z-index:10"><%
'	if instr(Sys_BillNo,"QZ")>0 then
'			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>小隊長&nbsp;林添福</span></td></tr>"
'			response.write "</table>"
'	elseif Sys_UnitID="TO00" then
'			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>組長&nbsp;莊松杰</span></td></tr>"
'			response.write "</table>" 

	'elseif trim(Session("Unit_ID"))="TN00" then
			
		
	'elseif trim(Sys_UnitLevelID)="1" then
'	else

'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'		response.write "</table>"
'	end if
%></div>
<div id="Layer37" style="position:absolute; left:<%=610+pageleft%>px; top:<%=1190+pagetop%>px; width:200px; height:46px; z-index:10"><%
		if instr(Sys_BillNo,"QZ")>0 then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">警員&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		elseif  trim(Session("Unit_ID"))="TG01" and ifnull(Sys_Jurgeday(0)) then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">警務佐&nbsp;梁建泰</span></td></tr>"
			response.write "</table>"

		elseif trim(Session("Unit_ID"))="TP00" then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		end if%></div>
<div id="Layer38" class="style3" style="position:absolute; left:<%=210+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:10px; z-index:10"><%=sys_Date(0)%></div>
<div id="Layer39" class="style3" style="position:absolute; left:<%=365+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:13px; z-index:10"><%=sys_Date(1)%></div>
<div id="Layer40" class="style3" style="position:absolute; left:<%=515+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:11px; z-index:10"><%=sys_Date(2)%></div>
<div id="Layer41" class="style3" style="position:absolute; left:<%=690+pageleft%>px; top:<%=1255+pagetop%>px; width:80px; height:12px; z-index:10"><%=Sys_BillFillerMemberID%></div>
<div id="Layer43" class="style3" style="position:absolute; left:<%=300+pageleft%>px; top:<%=1285+pagetop%>px; width:250px; height:12px; z-index:10"><%=Sys_DCIRETURNCARTYPE%></div>
</div>

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