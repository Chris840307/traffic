<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印--A4 size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style2 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style3 {font-family:"標楷體";font-size: 12px}
.style4 {font-family:"標楷體";font-size: 15px}
.style5 {font-family:"標楷體";font-size: 10px}
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style11 {font-family:"標楷體"; font-size: 8px; color:#ff0000; }
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
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
on Error Resume Next
if trim(request("printStyle"))<>"" then
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

Sys_Jurgeday=""

strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
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

rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
'if Not rsfound.eof then Sys_Driver=trim(rsfound("Driver"))
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
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))
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
Sum_Level=0:Sys_Level1=0:Sys_Level2=0
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
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
Sys_Sex="":SysBillMemID2=""
strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
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
if Not rssex.eof then SysBillMemID2=trim(rssex("BillMemID2"))

strSql="select a.LoginID,c.Content,a.ChName,b.UnitTypeID,b.UnitLevelID,b.UnitID,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
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

If cdbl(Sys_IllegalDate(0)&Sys_IllegalDate(1)&Sys_IllegalDate(2)) < 991225 Then
	Sys_BillFillerMemberID=""
	Sys_UnitName=replace(Sys_UnitName,"警察大","")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第一分隊","第一分局交通分隊")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第二分隊","第二分局交通分隊")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第三分隊","第三分局交通分隊")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第四分隊","第四分局交通分隊")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第五分隊","第五分局交通分隊")
	Sys_UnitName=replace(Sys_UnitName,"交通隊第六分隊","第六分局交通分隊")
	If cdbl(Sys_OwnerZip)> 408 and cdbl(Sys_OwnerZip)< 440 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台中市","台中縣")

		If Sys_OwnerZip=411 or Sys_OwnerZip=412 or Sys_OwnerZip=420 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=423 or Sys_OwnerZip=433 or Sys_OwnerZip=435 or Sys_OwnerZip=436 or Sys_OwnerZip=437 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if
	elseIf cdbl(Sys_OwnerZip)> 709 and cdbl(Sys_OwnerZip)< 746 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台南市","台南縣")

		If Sys_OwnerZip=710 or Sys_OwnerZip=730 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=712 or Sys_OwnerZip=721 or Sys_OwnerZip=722 or Sys_OwnerZip=726 or Sys_OwnerZip=732 or Sys_OwnerZip=737 or Sys_OwnerZip=741 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if

	elseIf cdbl(Sys_OwnerZip)> 813 and cdbl(Sys_OwnerZip)< 853 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"高雄市","高雄縣")

		If Sys_OwnerZip=830 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=820 or Sys_OwnerZip=842 or Sys_OwnerZip=843 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if

	elseIf cdbl(Sys_OwnerZip)> 206 and cdbl(Sys_OwnerZip)< 254 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"新北市","台北縣")

		If Sys_OwnerZip=220 or Sys_OwnerZip=221 or Sys_OwnerZip=231 or Sys_OwnerZip=234 or Sys_OwnerZip=235 or Sys_OwnerZip=236 or Sys_OwnerZip=238 or Sys_OwnerZip=241 or Sys_OwnerZip=242 or Sys_OwnerZip=247 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=224 or Sys_OwnerZip=237 or Sys_OwnerZip=239 or Sys_OwnerZip=251 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if
		
	end if
	
End if

chkJobID=""

	'Levle3 派出所 抓 所長jobid , 警備隊抓 隊長jobid , 交通分隊 抓 分隊長 jobid
	'       公園所 抓 巡佐
	if trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
		'交通分隊 抓 分隊長 jobid
		if trim(Sys_UnitID)="041S" or trim(Sys_UnitID)="042S"  or trim(Sys_UnitID)="044S" or trim(Sys_UnitID)="045S" or trim(Sys_UnitID)="048S" then	
			chkJobID="304"
		' 警備隊抓 隊長jobid
		elseif trim(Sys_UnitID)="0417" or trim(Sys_UnitID)="0426"  or trim(Sys_UnitID)="0447" or trim(Sys_UnitID)="0456" or trim(Sys_UnitID)="0486" or  trim(Sys_UnitID)="0437" then 
			chkJobID="303"
		' 分局組長抓 組長 jobid
		elseif trim(Sys_UnitID)="0411" then 
			chkJobID="318"
		'       公園所 抓 巡佐
		elseif trim(Sys_UnitID)="042T" then 
			chkJobID="317"
		'       六分局交通分隊抓 警務員
		elseif trim(Sys_UnitID)="048S" then 
			chkJobID="307"
		' 派出所 抓 所長jobid 
		else
			chkJobID="314"  	
		end if
	'level2 交通隊交安組(組長),交通隊拖吊場(警務員),交通隊直屬分隊(分隊長),
	'level2 交通隊一組(組長),交通隊二組(組長)
	elseif trim(Sys_UnitLevelID)="1" then
		'交通隊一組(組長)	'交通隊二組(組長)	交通隊交安組(組長)
		if trim(Sys_UnitID)="0461" or trim(Sys_UnitID)="0462" or  trim(Sys_UnitID)="0468" then 
			chkJobID="318"
		'交通隊拖吊場(警務員)
		elseif trim(Sys_UnitID)="046A" then 
			chkJobID="307"
		'交通隊直屬分隊(分隊長)
		elseif trim(Sys_UnitID)="0469" then 
			chkJobID="304"
		elseif trim(Sys_UnitID)="0406" then
			chkJobID="303"
		else
			chkJobID="0"  			'不應該出現的狀況
		end if
	end if
	'都抓不到 預設 許明義
	'Sys_jobName="隊長":Sys_MainChName="許明義"
	strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID="&chkJobID&") a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"
	set rsjob=conn.execute(strSQL)
	if Not rsjob.eof then
		Sys_jobName=rsjob("Content")
		Sys_MainChName=rsjob("ChName")
	end if
	rsjob.close

strSQL="select UnitName,Tel from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
Unit.close

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

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

If cdbl(Sys_IllegalDate(0))<102 Then

	If Sys_DCIReturnStation="60" Then
		Sys_STATIONNAME="臺中區監理所"
	ElseIf Sys_DCIReturnStation="61" Then
		Sys_STATIONNAME="臺中市監理站"
	ElseIf Sys_DCIReturnStation="63" Then
		Sys_STATIONNAME="豐原監理站"
	end if

end if


strSql="select MailNumber,MailTypeID,MailDate,MailChkNumber from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_MailCheckNumber=trim(rs("MailChkNumber"))
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
	strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
	set rscolor=conn.execute(strColor)
	Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
	rscolor.close
next
If ifnull(Sys_MailNumber) Then Sys_MailNumber=0
'smith 1041223 判斷大宗碼是451 or 295
'mailchknumber  511164 400295 36 40800 1
mypos = Instr(1, Sys_MailCheckNumber,"400451",1)
SYS_mailstationid = 295
'response.write Sys_MailCheckNumber
'response.write "<br>"
'response.write mypos
'response.end
if mypos>0 then SYS_mailstationid = 451

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,SYS_mailstationid,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,SYS_mailstationid,36
end if
strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" class="pageprint" style="position:relative;"><%
if showBarCode then%>
	<div id="Layer42" style="position:absolute; left:65px; top:545px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer43" style="position:absolute; left:65px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:185px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:185px; top:560px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:35px; top:595px; width:202px; height:36px; z-index:5"><%
if showBarCode then
	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"

end if
%></div>
<div id="Layer51" style="position:absolute; left:520px; top:590px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer52" style="position:absolute; left:515px; top:620px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer53" class="style4" style="position:absolute; left:125px; top:655px; width:250px; height:30px; z-index:3"><%
	If Sys_UnitID="046A" and instr(Sys_BillNo,"G1H")=0 then
		response.write "<font size=2>拖吊逕行舉發&nbsp;"
		response.write "<br>"
		response.write "　　　　　　"
	else
		response.write "<font size=2>逕行舉發&nbsp;"
		response.write "<br>"
		response.write "附採證照片&nbsp;"
	end if
	response.write Sys_CarColor&"</font>"%>
</div>
<div id="Layer54" style="position:absolute; left:295px; top:650px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:385px; top:650px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; left:290px; top:660px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; left:465px; top:675px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:650px; top:675px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" class="style4" style="position:absolute; left:125px; top:710px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" class="style4" style="position:absolute; left:285px; top:710px; width:200px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" class="style4" style="position:absolute; left:520px; top:705px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" class="style4" style="position:absolute; left:125px; top:738px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" class="style4" style="position:absolute; left:140px; top:765px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer64" class="style4" style="position:absolute; left:190px; top:765px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer65" class="style4" style="position:absolute; left:240px; top:765px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer66" class="style4" style="position:absolute; left:290px; top:765px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer67" class="style4" style="position:absolute; left:340px; top:765px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer68" class="style4" style="position:absolute; left:440px; top:760px; width:320px; height:31px; z-index:20"><%

	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里，經檢定合格儀器測照，時速"&Sys_IllegalSpeed&"公里，超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "(100以上)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "(80以上未滿100)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "(60以上未滿80)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "(40以上未滿60)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
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
	If Sys_UnitID="046A" and instr(Sys_BillNo,"G0H")=0 then response.write " (經科學儀器採證)"
%></div>
<div id="Layer69" class="style4" style="position:absolute; left:125px; top:785px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" class="style4" style="position:absolute; left:170px; top:810px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer71" class="style4" style="position:absolute; left:240px; top:810px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer72" class="style4" style="position:absolute; left:300px; top:810px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer73" class="style4" style="position:absolute; left:455px; top:815px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	'if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　 　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		'if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　 　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer75" class="style3" style="position:absolute; left:435px; top:855px; width:100px; height:30px; z-index:28"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></div>

<div id="Layer74" style="position:absolute; left:515px; top:850px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>

<div id="Layer76" style="position:absolute; left:415px; top:910px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"

	If instr(Sys_UnitName,"分局") = 0 Then
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=""15"" align=""center""><span class=""style2"">"

	'	If right(Sys_UnitName,1)="所" Then
	'		strSQL="select UnitName from Unitinfo where unitid='"&Sys_UnitTypeID&"'"
	'		set showuit=conn.execute(strSQL)
	'		Response.Write showuit("UnitName")
	'		Sys_UnitName=replace(Sys_UnitName,showuit("UnitName"),"")
	'		showuit.close
	'	else
			Response.Write "交通警察大隊"
	'	End if
		
		Response.Write "</span></td></tr>"

	end if

	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=""40"" align=""center""><span class=""style2"">"

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
	else
		Response.Write InstrAdd(replace(Sys_UnitName,"交通警察大隊",""),6)
	End if
	
	Response.Write "</span><br><span class=""style2"">"&Sys_UnitTEL&"</span></td></tr>"
	response.write "</table>"
%></div>
<div id="Layer77" style="position:absolute; left:600px; top:940px; width:200px; height:46px; z-index:31"><%
	If ifnull(SysBillMemID2) Then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""120"" align=""center""><span class=""style9"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
			response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
		end If 
	else
		response.Write "<table border=""0"" cellspacing=1 cellpadding=0>"
		Response.Write "<tr><td>"
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""60"" align=""center""><span class=""style11"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
			response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
		end If 
		Response.Write "</td><td>"

		if trim(Sys_MemberFileName2)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFileName2&""" height=""45"">"
		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;""  width=""60"" align=""center""><span class=""style11"">"&Sys_BillJobName2&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName2&"</span></td></tr>"
			response.write "</table><font size=2>　　"&Sys_MemLoginID2&"</font>"
		end If 
		Response.Write "</td></tr></table>"
	End if 
%></div>
<div id="Layer78" class="style4" style="position:absolute; left:205px; top:995px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer79" class="style4" style="position:absolute; left:360px; top:995px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer80" class="style4" style="position:absolute; left:520px; top:995px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer81" class="style4" style="position:absolute; left:580px; top:995px; width:120px; height:12px; z-index:36"></div>
</div>
<%
	if (i mod 100)=0 then response.flush
next
end if
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,8.08,5.08,5.08,5.08);
</s