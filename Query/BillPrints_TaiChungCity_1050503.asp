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
.style2 {font-family:"標楷體";font-size: 12px}
.style3 {font-family:"標楷體";font-size: 16px}
.style4 {font-family:"標楷體";font-size: 12px}
.style5 {font-family:"標楷體";font-size: 14px;}
.style6 {font-family:"標楷體";font-size: 22px; line-height:2;}
.style8 {font-size: 36px}
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style11 {font-family:"標楷體"; font-size: 8px; color:#ff0000; }
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
codebase="..\smsx.cab#Version=6,5,439,72">
</object>
<%
on Error Resume Next
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
for i=0 to Ubound(PBillSN)
sumCnt=sumCnt+1
if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

'===初始化(8/21)==
Sys_A_Name=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Driver=""
Sys_Owner=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_OwnerZip=""
Sys_DciReturnCarColor=""
Sys_STATIONNAME=""
Sys_StationTel=""
Sys_Rule1=""
Sys_Rule2=""
Sys_Level1=""
Sys_Level2=""
Sys_CarSimpleID=""

'================
strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Jurgeday=""

if Not rs.eof then
	Sys_CarSimpleID=trim(rs("CarSimpleID"))
	Sys_BillTypeID=trim(rs("BillTypeID"))
	
	Sys_Driver=trim(rs("Driver"))
	Sys_DriverID=trim(rs("DriverID"))
	Sys_DriverHomeAddress=trim(rs("DriverAddress"))
	Sys_DriverHomeZip=trim(rs("DriverZip"))

	Sys_Owner=trim(rs("Owner"))
	Sys_OwnerZip=trim(rs("OwnerZip"))
	Sys_OwnerAddress=trim(rs("OwnerAddress"))

	Sys_INSURANCE=trim(rs("INSURANCE"))
	Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
	Sys_RuleVer=trim(rs("RuleVer"))
	Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
	Sys_RuleSpeed=trim(rs("RuleSpeed"))
	Sys_Note=trim(rs("Note"))
	Sys_Rule4=trim(rs("Rule4"))
end If 

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

if Not rs.eof then Sys_ProjectID=trim(rs("ProjectID"))

rs.close

sys_ProjectName=""

If Sys_ProjectID="209" Then
	strSQL="select Name from Project where ProjectID='"&Sys_ProjectID&"'"

	set rspj=conn.execute(strSQL)
	If not rspj.eof Then sys_ProjectName=trim(rspj("Name"))
	rspj.close
End if 




strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

If ifnull(Sys_OwnerAddress) Then	

	strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
	rszip.close


	If Not ifnull(trim(rsfound("Driver"))) Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsfound.eof then
			Sys_Owner=trim(rsfound("Owner"))
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))			
		end if
	end If
	
End If 

If instr(Sys_OwnerAddress,"信箱") >0 or instr(Sys_OwnerAddress,"信相") > 0 Then
	errBillNo=errBillNo&rsbil("BillNo")&","&Sys_OwnerAddress&"「為郵政地址請確認」\n"
End if

If not ifnull(Sys_OwnerAddress) Then
	strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"

	conn.execute(strSQL)
End if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

if Not ifnull(Sys_OwnerZipName) then
	Sys_OwnerZipName=replace(trim(Sys_OwnerZipName),"臺","台")
	Sys_OwnerAddress=replace(trim(Sys_OwnerAddress),"臺","台")
	Sys_OwnerAddress=replace(trim(Sys_OwnerAddress),Sys_OwnerZipName,"")
end if

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"「姓名請確認」\n"
end if

If isnumeric(left(Sys_OwnerAddress,3)) Then
	Sys_OwnerZip=left(Sys_OwnerAddress,3)
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZip,"")
End if

If instr(Sys_OwnerAddress,"中縣") > 0 or instr(Sys_OwnerAddress,"雄縣") > 0 or instr(Sys_OwnerAddress,"南縣") > 0 or instr(Sys_OwnerAddress,"北縣") > 0 Then
	Sys_OwnerZipName=""
end if

Sum_Level=0:Sys_Level1=0:Sys_Level2=0
Sys_DCIReturnStation=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex="":SysBillMemID2=""
strSql="select * from BillBase where SN="&trim(PBillSN(i))
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
if Not rssex.eof then SysBillMemID2=trim(rssex("BillMemID2"))

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

If not ifnull(SysBillMemID2) Then
	strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.Tel,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&SysBillMemID2
	set mem=conn.execute(strsql)
	if Not mem.eof then Sys_MemLoginID2=trim(mem("LoginID"))
	if Not mem.eof then Sys_BillJobName2=trim(mem("Content"))
	if Not mem.eof then Sys_UnitFilename2=trim(mem("ImageFilename"))
	if Not mem.eof then Sys_ChName2=trim(mem("ChName"))
	if Not mem.eof then Sys_MemberFileName2=trim(mem("MemberFileName"))
	mem.close
End if 

chkJobID=""

	if trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
		chkJobID="303,304,305,307,314,318,1815,1838,1936,1937,1935,1938,1947,1948,1949" 

	elseif trim(Sys_UnitLevelID)="1" then
		chkJobID="303,304,318,307,1947,1948,1949,1838"
	end If
	If Sys_UnitID="0469" Or Sys_UnitID="0561" Then
		Sys_UnitID="0463"
		Sys_UnitName="交通警察大隊第一中隊"
		Sys_UnitTel="(04)23274655"
	ElseIf  Sys_UnitID="046A" Then
		Sys_UnitID="0464"
		Sys_UnitName="交通警察大隊第二中隊"
		Sys_UnitTel="(04)24210612"
	End if
	
	If Sys_UnitID="0463" and trim(Request("chktelunit"))<>"" Then
		strSQL="select Tel from Unitinfo where UnitID='"&trim(Request("chktelunit"))&"'"
		set rstel=conn.execute(strSQL)
		Sys_UnitTel=trim(rstel("Tel"))
		rstel.close
	End if

	If Sys_UnitID="0464" Then
		If instr(Sys_BillNo,"GP")>0 or instr(Sys_BillNo,"GQ")>0 Then
			strSQL="select Tel from Unitinfo where UnitID='0464'"
			set rstel=conn.execute(strSQL)
			Sys_UnitTel=trim(rstel("Tel"))
			rstel.close
		elseif instr(Sys_BillNo,"GR")>0 then
			Sys_UnitTel="(04)25277587"

		End if
	End if
	

	If Sys_UnitID="0463" Or Sys_UnitID="0464" Or Sys_UnitID="0465" Then
		chkJobID="303,304,1838"
	End If 
	'都抓不到 預設 許明義
	'Sys_jobName="隊長":Sys_MainChName="許明義"

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
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	rscolor.close
next

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,295,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,295,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",400,451,36"
	'response.end
end if

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
<%
If instr(Request("Sys_Batchnumber"),"WT")>0 then showBarCode=false
leftpx=0
toptpx=-2
If not ifnull(Request("Sys_Print")) Then leftpx=-15

If trim(Sys_OwnerZip)="001" then Sys_OwnerZip=""
%>
<div id="L78" style="position:relative;">

<div id="Layer50" class="style6" style="position:absolute; left:<%=(70+leftpx)%>px; top:<%=100+toptpx%>px; width:800px; z-index:5">
	<b>本&nbsp;&nbsp;&nbsp;&nbsp;郵&nbsp;&nbsp;&nbsp;&nbsp;件&nbsp;&nbsp;&nbsp;&nbsp;採&nbsp;&nbsp;&nbsp;&nbsp;單&nbsp;&nbsp;&nbsp;&nbsp;掛&nbsp;&nbsp;&nbsp;&nbsp;號&nbsp;&nbsp;&nbsp;&nbsp;郵&nbsp;&nbsp;&nbsp;&nbsp;寄<br>
	此&nbsp;&nbsp;&nbsp;&nbsp;收&nbsp;&nbsp;&nbsp;&nbsp;件&nbsp;&nbsp;&nbsp;&nbsp;回&nbsp;&nbsp;&nbsp;&nbsp;執&nbsp;&nbsp;&nbsp;&nbsp;不	&nbsp;&nbsp;&nbsp;需&nbsp;&nbsp;&nbsp;&nbsp;理&nbsp;&nbsp;&nbsp;&nbsp;會</b>
</div>

<div id="Layer42" class="style5" style="position:absolute; left:<%=(300+leftpx)%>px; top:<%=260+toptpx%>px; width:320px; height:36px; z-index:5">
	<%if trim(Sys_BillTypeID)="1" then
		Response.Write Sys_DriverHomeZip&"&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.write Sys_DriverZipName&Sys_DriverHomeAddress&"<br>"
		Response.Write funcCheckFont(Sys_Driver,16,1)
		Response.Write "&nbsp;&nbsp;台啟"

	elseif trim(Sys_BillTypeID)="2" then
		Response.Write Sys_OwnerZip&"&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,16,1)

		If instr(Sys_OwnerAddress,"中縣") > 0 or instr(Sys_OwnerAddress,"雄縣") > 0 or instr(Sys_OwnerAddress,"南縣") > 0 or instr(Sys_OwnerAddress,"北縣") > 0 Then
			Response.Write "(原登記地址)"
		End If 
		
		Response.Write "<br>" 
		Response.Write funcCheckFont(Sys_Owner,16,1)
		Response.Write "&nbsp;&nbsp;台啟"
		
	end If 
	If instr(Request("Sys_Batchnumber"),"WT")>0 then Response.Write "<font color=""red"">　　拖吊已繳費</font>" 
	%>
</div>

<div id="Layer43" class="style4" style="position:absolute; left:<%=(30+leftpx)%>px; top:<%=297+toptpx%>px; width:270px; height:36px; z-index:5">
	應到案處所：<%=Sys_STATIONNAME%><br>應到案處所電話：<%=Sys_StationTel%><br><%=Sys_UnitName%>
</div>

<div id="Layer45" class="style4" style="position:absolute; left:<%=(40+leftpx)%>px; top:<%=337+toptpx%>px; width:250px; height:36px; z-index:5">
	<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_1.jpg"""%>><br><span class="style2">　　<%=Sys_FirstBarCode%></span>
</div>


<div id="Layer44" class="style4" style="position:absolute; left:<%=(270+leftpx)%>px; top:<%=320+toptpx%>px; width:350px; height:36px; z-index:5">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"" width=""150"" height=""30"""%>><br>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%=Sys_MAILCHKNUMBER%>
</div>

<div id="Layer44" class="style2" style="position:absolute; left:<%=(455+leftpx)%>px; top:<%=325+toptpx%>px; width:250px; height:36px; z-index:5">
	大宗郵資已付掛號函件<br>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	第<%=Sys_MailNumber%>號
</div>


<%if trim(Sys_IMAGEFILENAMEB)<>"" then%>
<div style="position:absolute; left:30px; top:485px;"><img src="<%=Sys_IMAGEPATHNAME&Sys_IMAGEFILENAMEB%>" width="365" height="265"></DIV>
<%end if%>
<%if trim(Sys_IMAGEFILENAME)<>"" then%>
<div style="position:absolute; left:360px; top:485px;"><img src="<%=Sys_IMAGEPATHNAME&Sys_IMAGEFILENAME%>" width="365" height="265"></DIV>
<%end if%>
<%If instr(Request("Sys_Batchnumber"),"WT")>0 then%>
	<div id="Layer49" style="position:absolute; left:<%=(60+leftpx)%>px; top:685px; width:800px; height:36px; z-index:5">
	▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋
	<br>
	▋▋▋▋▋▋▋▋▋
	<br>
	▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋▋
	</div>
	<div id="Layer41" style="position:absolute; left:<%=(70+leftpx)%>px; top:<%=630+toptpx%>px; width:600px; height:36px; z-index:5">
		<font color="red" class="style6">拖吊保管場已代收交通罰鍰，請勿再行繳納。</font>
	</div>
	<%

else

	if showBarCode then%>
	<div id="Layer1" style="position:absolute; left:<%=(50+leftpx)%>px; top:<%=800+toptpx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
	<div id="Layer2" style="position:absolute; left:<%=(50+leftpx)%>px; top:<%=830+toptpx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end If 

End if %>
<!--
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:170px; top:800px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer4" style="position:absolute; left:170px; top:815px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
-->

<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer6" style="position:absolute; left:<%=(635+leftpx)%>px; top:<%=805+toptpx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer7" style="position:absolute; left:<%=(635+leftpx)%>px; top:<%=820+toptpx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer8" style="position:absolute; left:<%=(635+leftpx)%>px; top:<%=835+toptpx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>
<div id="Layer9" style="position:absolute; left:50px; top:<%=850+toptpx%>px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>

<div id="Layer10" style="position:absolute; left:500px; top:<%=850+toptpx%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:590px; top:<%=895+toptpx%>px; width:230px; height:12px; z-index:7"><font size=1><%=Sys_BillNo%></font></div>-->

<div id="Layer144" class="style4" style="position:absolute; left:<%=(120+leftpx)%>px; top:<%=910+toptpx%>px; width:250px; height:11px; z-index:3"><%
	
	If instr(Sys_BillNo,"GP")>0 then
		response.write "<font size=2>拖吊逕行舉發&nbsp;"
		response.write "<br>"
		response.write "　　　　　　"

	elseif instr(Sys_BillNo,"GR")>0 then
		response.write "<font size=2>拖吊逕行舉發&nbsp;"
		response.write "<br>"
		response.write "　　　　　　"

	elseif instr(Sys_BillNo,"GQ")>0 then
		response.write "<font size=2>拖吊逕行舉發&nbsp;"
		response.write "<br>"
		response.write "　　　　　　"

	elseif instr(Sys_BillNo,"G")>0 and instr(Sys_BillNo,"H")>0 then
		response.write "<font size=2>逕行舉發&nbsp;"
		response.write "<br>"
		response.write "附採證照片&nbsp;"
	end if
	response.write Sys_CarColor&"</font>"%>
</div>
<% uppx=4 %>
<div id="Layer17" style="position:absolute; left:<%=(630+leftpx)%>px; top:<%=935+toptpx-uppx%>px; width:99px; height:12px; z-index:10"><span class="style3"><%=fastring%></span></div>

<div id="Layer18" style="position:absolute; left:<%=(120+leftpx)%>px; top:<%=960+toptpx-uppx%>px; width:100px; height:14px; z-index:11"><span class="style3"><%=Sys_CarNo%></span></div>
<div id="Layer19" style="position:absolute; left:<%=(290+leftpx)%>px; top:<%=960+toptpx-uppx%>px; width:200px; height:20px; z-index:12"><span class="style3"><%=Sys_DCIRETURNCARTYPE%></span></div>
<div id="Layer20" style="position:absolute; left:<%=(550+leftpx)%>px; top:<%=955+toptpx-uppx%>px; width:201px; height:17px; z-index:13"><span class="style3"><%=funcCheckFont(Sys_Owner,14,1)%></span></div>
<div id="Layer21" style="position:absolute; left:<%=(180+leftpx)%>px; top:<%=985+toptpx-uppx%>px; width:750px; height:13px; z-index:14"><span class="style3"><%
	Response.Write Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,14,1)

	If instr(Sys_OwnerAddress,"中縣") > 0 or instr(Sys_OwnerAddress,"雄縣") > 0 or instr(Sys_OwnerAddress,"南縣") > 0 or instr(Sys_OwnerAddress,"北縣") > 0 Then
			Response.Write "(原登記地址)"
	End if
%></span></div>

<div id="Layer22" style="position:absolute; left:<%=(110+leftpx)%>px; top:<%=1007+toptpx-uppx%>px; width:40px; height:13px; z-index:15"><span class="style5"><%=Sys_IllegalDate(0)%></span></div>

<div id="Layer23" style="position:absolute; left:<%=(160+leftpx)%>px; top:<%=1007+toptpx-uppx%>px; width:40px; height:17px; z-index:16"><span class="style5"><%=Sys_IllegalDate(1)%></span></div>

<div id="Layer24" style="position:absolute; left:<%=(210+leftpx)%>px; top:<%=1007+toptpx-uppx%>px; width:40px; height:16px; z-index:17"><span class="style5"><%=Sys_IllegalDate(2)%></div>

<div id="Layer25" style="position:absolute; left:<%=(260+leftpx)%>px; top:<%=1007+toptpx-uppx%>px; width:40px; height:16px; z-index:18"><span class="style5"><%=right("00"&Sys_IllegalDate_h,2)%></span></div>

<div id="Layer26" style="position:absolute; left:<%=(310+leftpx)%>px; top:<%=1007+toptpx-uppx%>px; width:40px; height:13px; z-index:19"><span class="style5"><%=right("00"&Sys_IllegalDate_m,2)%></span></div>

<div id="Layer27" class="style5" style="position:absolute; left:<%=(390+leftpx)%>px; top:<%=1010+toptpx-uppx%>px; width:315px; height:31px; z-index:20"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	If not ifnull(sys_ProjectName) Then Response.Write sys_ProjectName&"<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里，經檢定合格儀器測照，時速"&Sys_IllegalSpeed&"公里，超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "(100以上)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "(80以上未滿100)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "(60以上未滿80)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "(40以上未滿60)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "(20以上未滿40)。"
			else
				response.write "(未滿20公里)。"
			end if
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		response.write Sys_IllegalRule1
		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"	
	end if
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		response.write "<br>"&Sys_IllegalRule2
	end if
	'If (Sys_UnitID="046A" Or Sys_UnitID="0464") and instr(Sys_BillNo,"G5H")=0 and instr(Sys_BillNo,"GP")=0 and instr(Sys_BillNo,"GQ")=0 then response.write " (經科學儀器採證)"

	'If Sys_ProjectID="210" Then Response.Write "(佔用公車格)"
	
%></div>
<div id="Layer28" style="position:absolute; left:<%=(110+leftpx)%>px; top:<%=1025+toptpx-uppx%>px; width:240px; height:15px; z-index:21"><span class="style5"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:<%=(120+leftpx)%>px; top:<%=1057+toptpx-uppx%>px; width:34px; height:11px; z-index:22"><span class="style3"><%=Sys_DealLineDate(0)%></span></div>
<div id="Layer30" style="position:absolute; left:<%=(190+leftpx)%>px; top:<%=1057+toptpx-uppx%>px; width:35px; height:13px; z-index:23"><span class="style3"><%=Sys_DealLineDate(1)%></span></div>
<div id="Layer31" style="position:absolute; left:<%=(265+leftpx)%>px; top:<%=1057+toptpx-uppx%>px; width:32px; height:15px; z-index:24"><span class="style3"><%=Sys_DealLineDate(2)%></span></div>

<div id="Layer46" style="position:absolute; left:<%=(430+leftpx)%>px; top:<%=1065+toptpx-uppx%>px; width:32px; height:15px; z-index:24"><span class="style2"><%
	'if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
%></span></div>

<div id="Layer47" style="position:absolute; left:<%=(430+leftpx)%>px; top:<%=1080+toptpx-uppx%>px; width:32px; height:15px; z-index:24"><span class="style2"><%
	'if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
%></span></div>

<div id="Layer32" style="position:absolute; left:<%=(395+leftpx)%>px; top:<%=1100+toptpx-uppx%>px; width:400px; height:49px; z-index:29"><span class="style5"><%

	if instr(Request("Sys_Batchnumber"),"WT")>0 then
		Sys_Level1="":Sys_Level2=""
	end if

	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)

	If len(trim(Sys_Rule1)) > 7 Then 
		Response.Write "-"&right(trim(Sys_Rule1),1)
		Response.Write "　"
	else
		Response.Write "　　"
	end if

	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)
		If len(trim(Sys_Rule2)) > 7 Then 
			Response.Write "-"&right(trim(Sys_Rule2),1)
			Response.Write "　"
		else
			Response.Write "　　"
		end if

		'if len(trim(Sys_Rule2))>7 then response.write right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　"&Sys_Level2
	end if

%></span></div>
<div id="Layer33" style="position:absolute; left:<%=(380+leftpx)%>px; top:<%=1135+toptpx-uppx%>px; width:400px; height:30px; z-index:1"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
	if trim(Sys_Rule4)<>"" and instr(Request("Sys_Batchnumber"),"WT")>0 then response.write "(已由領車人繳納. 收據字號 "&Sys_Rule4&")"
%></div>
<div id="Layer34" style="position:absolute; left:<%=(625+leftpx)%>px; top:<%=1145+toptpx-uppx%>px; width:90px; height:30px; z-index:28"><span class="style2"><%=Sys_STATIONNAME%></span></div>


<div id="Layer35" style="position:absolute; left:<%=(400+leftpx)%>px; top:<%=1195+toptpx-uppx%>px; width:150px; height:49px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"

	If instr(Sys_UnitName,"分局") = 0 Then
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=""15"" align=""center""><span class=""style10"">"

	'	If right(Sys_UnitName,1)="所" Then
	'		strSQL="select UnitName from Unitinfo where unitid='"&Sys_UnitTypeID&"'"
	'		set showuit=conn.execute(strSQL)
	'		Response.Write showuit("UnitName")
	'		Sys_UnitName=replace(Sys_UnitName,showuit("UnitName"),"")
	'		showuit.close
	'	else
		If Sys_UnitTypeID = "0406" Then
			Response.Write "保安警察大隊"
		else
			Response.Write "交通警察大隊"
		End if 
			
	'	End if
		
		Response.Write "</span></td></tr>"
	end if

	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=""40"" align=""center""><span class=""style10"">"

	If instr("5E01,5E02,5E03,5E04",Sys_UnitID)>0 Then
		Response.Write replace(Sys_UnitName,"交通警察大隊","烏日分隊-")

	elseif instr("4D14,4D15,4D16",Sys_UnitID)>0 Then
		Response.Write replace(Sys_UnitName,"交通警察大隊","清水分隊-")

	elseif instr("4FA9",Sys_UnitID)>0 Then
		Response.Write replace(Sys_UnitName,"交通警察大隊","大甲分隊-")

	elseif instr("4H06",Sys_UnitID)>0 Then
		Response.Write replace(Sys_UnitName,"交通警察大隊","太平分隊-")

	elseif instr("4BA7",Sys_UnitID)>0 Then
		Response.Write replace(Sys_UnitName,"交通警察大隊","東勢分隊-")

	elseIf instr(Sys_UnitName,"分局") > 0 Then
		Response.Write replace(replace(Sys_UnitName,"分局","分局<br>"),"交通警察大隊","")
	elseIf Sys_UnitTypeID = "0406" Then
		Response.Write InstrAdd(replace(Sys_UnitName,"保安警察大隊",""),6)
	else
		Response.Write InstrAdd(replace(Sys_UnitName,"交通警察大隊",""),6)
	End if
	
	Response.Write "</span><br><span class=""style10"">"&Sys_UnitTEL&"</span></td></tr>"
	response.write "</table>"
%></div>

<div id="Layer37" style="position:absolute; left:<%=(575+leftpx)%>px; top:<%=1200+toptpx-uppx%>px; width:200px; height:46px; z-index:31"><%
	If ifnull(SysBillMemID2) Then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""120"" align=""center""><span class=""style9"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
			response.write "</table><font size=1>　　"&Sys_BillFillerMemberID&"</font>"
		end If 
	else
		response.Write "<table border=""0"" cellspacing=1 cellpadding=0>"
		Response.Write "<tr><td>"
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""60"" align=""center""><span class=""style11"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
			response.write "</table><font size=1>　　"&Sys_BillFillerMemberID&"</font>"
		end If 
		Response.Write "</td><td>"

		if trim(Sys_MemberFileName2)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFileName2&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""60"" align=""center""><span class=""style11"">"&Sys_BillJobName2&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName2&"</span></td></tr>"
			response.write "</table><font size=1>　　"&Sys_MemLoginID2&"</font>"
		end If 
		Response.Write "</td></tr></table>"
	End if 
	
%></div>
<div id="Layer38" style="position:absolute; left:160px; top:<%=1242+toptpx-uppx%>px; width:60px; height:10px; z-index:32"><span class="style3"><%=sys_Date(0)%></span></div>
<div id="Layer39" style="position:absolute; left:210px; top:<%=1242+toptpx-uppx%>px; width:60px; height:13px; z-index:33"><span class="style3"><%=sys_Date(1)%></span></div>
<div id="Layer40" style="position:absolute; left:260px; top:<%=1242+toptpx-uppx%>px; width:60px; height:11px; z-index:34"><span class="style3"><%=sys_Date(2)%></span></div>
<div id="Layer145" style="position:absolute; left:400px; top:<%=1270+toptpx-uppx%>px; width:150px; height:11px; z-index:37"><font size=2><%=Sys_CarColor%></font></div>
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
	printWindow(true,5.08,5.08,5.08,5.08);
</script>