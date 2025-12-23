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
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style11 {font-size: 14px}
.style15 {font-size: 15px}
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style13 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
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
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

strCity="select value from Apconfigure where id=35"
set rsCity=conn.execute(strCity)
sys_title=trim(rsCity("value"))
rsCity.close

for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
'===初始化(8/21)==
Sys_BillNo=""
Sys_CarNo=""
Sys_DriverHomeZip=""
Sys_Driver=""
Sys_Owner=""
Sys_OwnerAddress="":Sys_OwnerZip="":Sys_OwnerZipName=""
Orz_Driver="":Orz_DriverHomeAddress="":Orz_DriverHomeZip=""
'================
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Orz_Driver=trim(rs("Owner"))
if Not rs.eof then Orz_DriverHomeAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Orz_DriverHomeZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))

Sys_ILLEGALADDRESS=replace(Sys_ILLEGALADDRESS&"","台","臺")

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

	Orz_DriverHomeAddress=trim(rsfound("OwnerAddress"))
	Orz_DriverHomeZip=trim(rsfound("OwnerZip"))

	If not ifnull(Sys_OwnerAddress) Then

		strSQL="Update Billbase set Owner='"&rsfound("Owner")&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&chkaddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"

		conn.execute(strSQL)
	end If 
end if

if not ifnull(Sys_OwnerAddress) then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"  "," ")," ","　")
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_OwnerZipName=replace(Sys_OwnerZipName&"","台","臺")

Sys_OwnerAddress=replace(Sys_OwnerAddress&"",Sys_OwnerZipName,"")

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

if not ifnull(Orz_DriverHomeZipName) then 
	Orz_DriverHomeZipName=replace(Orz_DriverHomeZipName&"","台","臺")
	Orz_DriverHomeAddress=replace(Orz_DriverHomeAddress&"",Orz_DriverHomeZipName,"")
end if


'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
end if

Sys_DCIReturnStation=0
Sum_Level=0
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
strSql="select distinct BillFillerMemberID,BillMemID2,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID,imagefilename from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
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
if Not rssex.eof then Billbaseimagefilename=trim(rssex("imagefilename"))
Sys_IisImagePath="":Sys_ImageFileNameA="":Sys_ImageFileNameB=""
' 讀取違規影像，由Kevin的影像建檔處寫入
strSQL="select * from BillIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	Sys_ImageFileNameA=trim(rsimage("ImageFileNameA"))
	Sys_ImageFileNameB=trim(rsimage("ImageFileNameB"))
end If 

stytleColor=""

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

stytleColor=""

if Sys_UnitTypeID = "09A7" Then

	stytleColor="style=""width:700px; height:275px; border-left:0px solid #000000; border-right:0px solid #000000;border-bottom:0px solid 000000;"""

elseIf Sys_UnitTypeID = "9A00" Then

	stytleColor="style=""width:700px; height:275px; border-left:5px solid #f0a908; border-right:5px solid #f0a908;border-bottom:5px solid #f0a908;"""

elseif Sys_UnitTypeID = "9B00" Then

	stytleColor="style=""width:700px; height:275px; border-left:5px solid #80f509; border-right:5px solid #80f509;border-bottom:5px solid #80f509;"""

elseif Sys_UnitTypeID = "9C00" Then

	stytleColor="style=""width:700px; height:275px; border-left:5px solid #09f7ae; border-right:5px solid #09f7ae;border-bottom:5px solid #09f7ae;"""

elseif Sys_UnitTypeID = "9D00" Then

	stytleColor="style=""width:700px; height:275px; border-left:5px solid #0909f3; border-right:5px solid #0909f3;border-bottom:5px		solid #0909f3;"""

End if 


if Not ifnull(Sys_BillFillerMemberID2) then
	strSql="select ChName,ImageFilename as MemberFileName from MemberData where MemberID="&Sys_BillFillerMemberID2
	set mem=conn.execute(strsql)
	if Not mem.eof then Sys_MemberFileName2=trim(mem("MemberFileName"))
	if Not mem.eof then Sys_ChName2=trim(mem("ChName"))
	mem.close
end if

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

if trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
	chkJobID="302,304,305,320" 

	strSQL="select a.ChName,b.Content,b.ID,b.showorder from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitTypeID&"' and JobID in("&chkJobID&")) a,(select ID,showorder,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by b.showorder,b.id"

elseif trim(Sys_UnitLevelID)="1" then
	chkJobID="303,306"

	strSQL="select a.ChName,b.Content,b.ID,b.showorder from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,showorder,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by b.showorder,b.id"
end If

Sys_jobName="":Sys_MainChName=""


set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close


Stop_IllegalDate_h="":Stop_IllegalDate_m=""
Stop_DealLineDate=split("","\")
Stop_IllegalDate=split("","\")

If Sys_Rule1 = "5620001" or Sys_Rule1 = "5630001" Then
	If not ifnull(Billbaseimagefilename) Then
		tmpStopFile=split(Billbaseimagefilename,"\")

		strSQL="select DealLineDate,IllegalDate from billbase where CarNo='"&Sys_CarNo&"' and billno is null and ImageFileNameB is not null and imagepathname='"&replace(tmpStopFile(1),".jpg","")&"' and recordstateid=0"

		set rsstop=conn.execute(strSQL)
		if Not rsstop.eof then
			Stop_IllegalDate=split(gArrDT(trim(rsstop("IllegalDate"))),"-")
			Stop_IllegalDate_h=hour(trim(rsstop("IllegalDate")))
			Stop_IllegalDate_m=minute(trim(rsstop("IllegalDate")))

			Stop_DealLineDate=split(gArrDT(trim(rsstop("DealLineDate"))),"-")

			Sys_IllegalDate=split(gArrDT(DateAdd("d",1,rsstop("DealLineDate"))),"-")
			Sys_IllegalDate_h="00"
			Sys_IllegalDate_m="00"
		end if
		rsstop.close
	End if	
End if

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

strSQL="update billmailhistory set mailchknumber='"&Sys_MailNumber&" 951000 17' where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
conn.execute(strSQL)

if trim(Sys_BillTypeID)="1" then

	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95100017","220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95100017","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if

if Sys_DriverHomeZip="001" then Sys_DriverHomeZip=""
if Sys_OwnerZip="001" then Sys_OwnerZip=""

Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo&"-"&Sys_CarNo

firstBacrCode=right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&"D"&Sys_StationID

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer000" style="position:absolute; left:30px; top:0px; z-index:5"><%
	Response.Write "<img src=""d:\舉發單樣本\TaiTung01.gif"" width=""715"" height=""1290"">"
	%>
</div>

<div id="Layer001" style="position:absolute; left:30px; top:0px; z-index:5">
	<table <%=stytleColor%>>
		<tr><td>&nbsp;</td></tr>
	</table>
</div>

<div id="Layer070" style="position:absolute; left:430px; top:10px; z-index:8"><%
	Response.Write "<img src=""../Image/BillNoPage.gif"" width=""80"">"
	%>
</div>

<div id="Layer071" style="position:absolute; left:445px; top:43px; font-size: 12px; z-index:9"><%
	Response.Write replace(gArrDT(date),"-",".")
	%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:<%=75+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write sys_title&replace(SysUnit,sys_title,"")
%>
</div>

<div id="Layer66" class="style3" style="position:absolute; left:<%=345+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write SysUnitAddress
%>
</div>

<div id="Layer01" class="style3" style="position:absolute; left:120px; top:16px; z-index:8"><%
	response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
%>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:125px; top:60px; z-index:8"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:155px; top:45px; z-index:4"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:560px; top:287px; z-index:4">
	　　　<b>第<%=Sys_MailNumber%>號</b><br>
	<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>><br>
    　<b><%="　"&Sys_MAILCHKNUMBER%></b>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:405px; top:270px; z-index:10"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:130px; top:285px; z-index:8"><%
	if trim(Sys_BillTypeID)="1" then
		response.write Sys_Driver
	elseif trim(Sys_BillTypeID)="2" then
		response.write funcCheckFont(Sys_Owner,16,1)
	end if%>　台啟
</div>

<div id="Layer05" class="style3" style="position:absolute; left:350px; top:330px; z-index:4"><%
	Response.Write "<img  src=""../BarCodeImage/"&Sys_BillNo&"_1.jpg"">"%>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:110px; top:309px; width:330px; z-index:8"><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			if trim(Sys_BillTypeID)="1" then
				response.write Orz_DriverHomeZip&"　"
				response.write replace(Orz_DriverHomeZipName&Orz_DriverHomeAddress,Orz_DriverHomeZipName&Orz_DriverHomeZipName,Orz_DriverHomeZipName)
			elseif trim(Sys_BillTypeID)="2" then
				response.write Sys_OwnerZip&" "
				response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
			end if
			response.write "<br><br>"
	%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:110px; top:360px; z-index:8"><%
	'Response.Write Sys_CarNo&"　"&Sys_A_Name&"　"&Sys_CarColor&"　"&Sys_STATIONNAME
	Response.Write Sys_CarNo&"　"&Sys_STATIONNAME
%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:450px; top:360px; z-index:8"><%
	Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>

<div id="Layer091" class="style3" style="position:absolute; left:110px; top:400px; z-index:8"><%
	'Response.Write Sys_CarNo&"　"&Sys_A_Name&"　"&Sys_CarColor&"　"&Sys_STATIONNAME
	response.Write sys_title&replace(SysUnit,sys_title,"")
%>
</div>

<!-------------------------- 判斷 billillegalimage 有沒有這些圖檔 ------------------------------>
<!---------- todo 依據法條判斷, 如果是闖紅燈, 要把 a 檔的 xxxxx_a.jpg 換成 b檔的xxxxxx_b.jpg ---------------------------------------->
<%if trim(Sys_ImageFileNameA)<>"" then%>
	<div id="Layer09" style="position:absolute; left:38px; top:480px; z-index:8"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	%></DIV>
<%
elseIf (trim(Sys_Rule1)="5620001" or trim(Sys_Rule1)="5630001") and trim(Billbaseimagefilename)<>"" Then%>
	<div id="Layer09" style="position:absolute; left:38px; top:480px; z-index:8"><%
		response.write "<img src=""../StopCarPicture/"&Billbaseimagefilename&""" width=""390"" height=""280"">"
	%></DIV>
<%End if
'response.write Sys_Rule1 & "_" 
'response.write Sys_ImageFileNameB
' ssmith 20091015 紅燈月線不顯示B圖
%>
<%if trim(Sys_ImageFileNameB)<>"" then%>
	<div id="Layer10" style="position:absolute; left:430px; top:485px; z-index:8"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:50px; top:805px; width:202px; height:36px; z-index:8">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:50px; top:840px; width:202px; height:36px; z-index:8">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:170px; top:810px; width:202px; height:36px; z-index:8">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:825px; width:202px; height:36px; z-index:8">v</div>
<%end if%>

<div id="Layer9" style="position:absolute; left:20px; top:865px; width:202px; height:36px; z-index:4"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>

<div id="Layer26" class="style3" style="position:absolute; left:<%=115+pageleft%>px; top:<%=902+pagetop%>px; z-index:10"><%
	if showBarCode then
		response.write firstBacrCode
	end if
%></div>

<div id="Layer11" class="style3" style="position:absolute; left:<%=520+pageleft%>px; top:<%=890+pagetop%>px; z-index:10"><%="東警交"&"　　　　"&Sys_BillNo%></div>

<div id="Layer10" style="position:absolute; left:490px; top:840px; width:233px; height:32px; z-index:4"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<!--
<div id="Layer12" style="position:absolute; left:115px; top:920px; width:150px; height:11px; z-index:8"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>
-->
<!--20130509 交通隊施志林說「附採證照片」拿掉，以免廠牌太長時蓋到車號 -->
<div id="Layer12" style="position:absolute; left:115px; top:920px; width:150px; height:11px; z-index:8"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if (int(Sys_Rule1)<>5620001 and int(Sys_Rule1)<>5630001) then response.write ""%>　<%=Sys_CarColor%></span></div>


<div id="Layer13" style="position:absolute; left:255px; top:910px; width:28px; height:11px; z-index:8"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:365px; top:910px; width:324px; height:10px; z-index:8"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>

<div id="Layer15" style="position:absolute; left:255px; top:920px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:420px; top:920px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:615px; top:920px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:120px; top:955px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:255px; top:955px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:495px; top:955px; width:300px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" style="position:absolute; left:120px; top:980px; width:507px; height:13px; z-index:14"><%=Orz_DriverHomeZip&" "&funcCheckFont(Orz_DriverHomeZipName&Orz_DriverHomeAddress,16,1)%></div>

<div id="Layer22" style="position:absolute; left:105px; top:1000px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:165px; top:1000px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:215px; top:1000px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:255px; top:1000px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:305px; top:1000px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:390px; top:1005px; width:610px; height:31px; z-index:20"><span class="style3"><%
	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"

		if len(Sys_IllegalRule1)<26 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	end if
	if trim(Sys_Rule2)<>"0" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end If
	
	If Ubound(Stop_DealLineDate)>0 Then

		Response.Write "<br>停車時間："&Stop_IllegalDate(0)&"年"&Stop_IllegalDate(1)&"月"&Stop_IllegalDate(2)&"日"&Stop_IllegalDate_h&"時"&Stop_IllegalDate_m&"分"
				
		Response.Write "<br>違規成立日即繳費期限"&Stop_DealLineDate(0)&"年"&Stop_DealLineDate(1)&"月"&Stop_DealLineDate(2)&"日次日零時認定"
		
	end if

'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:110px; top:1025px; width:220px; height:15px; z-index:21"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:115px; top:1055px; width:50px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:205px; top:1055px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:275px; top:1055px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style3" style="position:absolute; left:395px; top:1075px; width:400px; height:49px; z-index:29"><%
	response.write left(trim(Sys_Rule1),2)&"　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if

%></div>

<div id="Layer34" style="position:absolute; left:380px; top:1120px; width:400px; height:30px; z-index:4"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" style="position:absolute; left:635px; top:1125px; width:100px; height:40px; z-index:28"><span class="style7"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>

<div id="Layer35" style="position:absolute; left:395px; top:1165px; width:100px; height:49px; z-index:29"><%
	If Sys_UnitLevelID=1 Then
		response.write "<table border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td class=""style3"">台東縣警察局<br>交通警察隊</td></tr>"
		response.write "<tr><td class=""style3"">TEL(089)328344</td></tr>"
		response.write "</table>"
	elseIf Sys_UnitLevelID=2 Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" width=""100"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td width=""100"" class=""style10"">台東縣警察局<br>"&Sys_UnitName&"</td></tr>"
		response.write "<tr><td width=""100"" class=""style10"">TEL"&Sys_UnitTel&"</td></tr>"
		response.write "</table>"
	elseIf Sys_UnitLevelID=3 Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" width=""100"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td width=""100"" class=""style10"">台東縣警察局<br>"&right(Sys_UnitName,len(Sys_UnitName)-instr(Sys_UnitName,"局"))&"</td></tr>"
		response.write "<tr><td width=""100"" class=""style10"">TEL"&Sys_UnitTel&"</td></tr>"
		response.write "</table>"
	end if
%></div>
<div id="Layer36" style="position:absolute; left:595px; top:1195px; width:100px; height:43px; z-index:30"><%
'	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">舉發單專用章<br>"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'	response.write "</table>"
%></div>
<div id="Layer37" style="position:absolute; left:595px; top:1185px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""60"" height=""20""><br>"
	end if
	'response.Write "<font size=2>"&Sys_ChName&"</font>"

	response.write "<table style=""border-bottom:0; border-top:0 ;border-left:0 ; border-right:0 ;border-color:#ff0000;border-style:solid;"" border=""0"" cellspacing=0 cellpadding=0 width=""100"">"
	response.write "<tr>"

	Response.Write "<td style=""border-color:#ff0000;border-style:solid;border-width:1px;"" width=""50%"" height=25 align=""center"" nowrap><span class=""style13"">"&Sys_BillJobName&"</span>&nbsp;<span class=""style13"">"&Sys_ChName&"</span></td>"

	Response.Write "</tr>"
	response.write "</table><br><br><br>"
'	if trim(Sys_MemberFilename2)<>"" then
'		response.write "<br><img src=""../Member/Picture/"&Sys_MemberFilename2&""" width=""60"" height=""20""><br>"
'	end if
'	response.Write "<font size=2>"&Sys_ChName2&"</font>"
%></div>
<div id="Layer38" style="position:absolute; left:205px; top:1240px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:360px; top:1240px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:510px; top:1240px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:670px; top:1235px; width:80px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>
<div id="Layer43" style="position:absolute; left:300px; top:1265px; width:250px; height:12px; z-index:36"><%=Sys_DCIRETURNCARTYPE%></div>
</div>

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
	//window.print();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>