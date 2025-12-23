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
.style1 {font-family:"標楷體"; font-size: 14px;}
.style2 {font-family:"標楷體"; font-size: 16px;}
.style3 {font-family:"標楷體"; font-size: 18px;}
.style4 {font-family:"標楷體"; font-size: 26px; line-height:2;}
.style5 {font-family:"標楷體"; font-size: 12px;}
.style6 {font-family:"標楷體"; font-size: 14px; color:#ff0000;}
.style7 {font-family:"標楷體"; font-size: 20px;}
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
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
Sys_Print=Ubound(PBillSN)
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

strUInfo="select * from Apconfigure where ID=31"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then theCity=trim(rsUInfo("value"))
rsUInfo.close

strUInfo="select * from Apconfigure where ID=52"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thePollic=trim(rsUInfo("value"))
rsUInfo.close

for i=0 to Sys_Print
if i<>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Rule4=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

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
Sum_Level=0
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
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
set rssex=conn.execute(strSql)
if trim(Sys_BillTypeID)="1" then
	if Not rssex.eof then
		If not ifnull(Trim(rssex("DriverID"))) Then
			If Mid(Trim(rssex("DriverID")),2,1)="1" Then
				Sys_Sex="男"
			elseif Mid(Trim(rssex("DriverID")),2,1)="2" Then
				Sys_Sex="女"
			End if
		End if
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
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_BillFillerMemberID=trim(rssex("BillFillerMemberID"))
rssex.close

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
elseif Sys_UnitLevelID=2 and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
end if
set unit=conn.Execute(strSQL)
theUnitID=trim(unit("UnitID"))
theUnitName=trim(unit("UnitName"))
theSubUnitSecBossName=trim(unit("SecondManagerName"))
theBigUnitBossName=trim(unit("ManageMemberName"))
theContactTel=trim(unit("Tel"))
theBankAccount=trim(unit("BankAccount"))
theBankName=trim(unit("BankName"))
theUnitAddress=trim(unit("Address"))
unit.close

if trim(Sys_UnitLevelID)="3" then
	chkJobID="303,314"
elseif trim(Sys_UnitLevelID)="1" then
	chkJobID="303"
end if

strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"

set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

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

strSQL="select * from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set mem=conn.execute(strSQL)
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then theUnitAddress=trim(mem("Address"))
mem.close

strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))

rs.close
strSql="select b.Content from BillFastenerDetail a,DCICode b where a.FastenerTypeID=b.ID and b.TypeID=6 and a.BillSN="&trim(rsbil("BillSN"))&" and a.CarNo='"&trim(rsbil("CarNo"))&"'"
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

if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,36
else
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,36
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

showBarCode=true

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
If showBarCode=true Then
	if left(trim(Sys_Rule2),2)="40" or (int(Sys_Rule2)>3310101 and int(Sys_Rule2)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			if Sys_IllegalSpeed-Sys_RuleSpeed>=60 then showBarCode=false
		end if
	elseif left(trim(Sys_Rule2),5)="12101" or left(trim(Sys_Rule2),5)="12102" or left(trim(Sys_Rule2),5)="12105" or left(trim(Sys_Rule2),5)="12106" or left(trim(Sys_Rule2),5)="12107" or left(trim(Sys_Rule2),5)="12109" or left(trim(Sys_Rule2),5)="12110" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="16105" or left(trim(Sys_Rule2),5)="25103" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="25103" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="294" or left(trim(Sys_Rule2),3)="295" or left(trim(Sys_Rule2),3)="303" or left(trim(Sys_Rule2),3)="314" or left(trim(Sys_Rule2),3)="362" or left(trim(Sys_Rule2),3)="363" or left(trim(Sys_Rule2),3)="621" or left(trim(Sys_Rule2),3)="624" or left(trim(Sys_Rule2),3)="625" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="13" or left(trim(Sys_Rule2),2)="18" or left(trim(Sys_Rule2),2)="20" or left(trim(Sys_Rule2),2)="21" or left(trim(Sys_Rule2),2)="23" or left(trim(Sys_Rule2),2)="24" or left(trim(Sys_Rule2),2)="26" or left(trim(Sys_Rule2),2)="27" or left(trim(Sys_Rule2),2)="35" or left(trim(Sys_Rule2),2)="37" or left(trim(Sys_Rule2),2)="43" or left(trim(Sys_Rule2),2)="54" or left(trim(Sys_Rule2),2)="61" then
		showBarCode=false
	elseif int(Sys_Rule2)=29300012 or int(Sys_Rule2)=29300022 or int(Sys_Rule2)=3400003 or int(Sys_Rule2)=3400004 then
		showBarCode=false
	end if
End if
%>
<div id="L78" style="position:relative;">
<div id="Layer1" class="style3" style="position:absolute; left:80px; top:10px; width:250px; height:30px; z-index:5"><%=theUnitAddress%></div>

<div id="Layer2" class="style3" style="position:absolute; left:80px; top:30px; width:250px; height:30px; z-index:5"><%=thenPasserCity&"("&replace(theUnitName,trim(thenPasserCity),"")&")"%></div>

<div id="Layer3" class="style3" style="position:absolute; left:80px; top:50px; width:250px; height:30px; z-index:5">電話：<%=theContactTel%></div>

<div id="Layer4" class="style2" style="position:absolute; left:60px; top:90px; width:300px; height:30px; z-index:5"><%=Sys_OwnerZip&"<br>"&Sys_OwnerZipName&Sys_OwnerAddress%></div>

<div id="Layer5" class="style2" style="position:absolute; left:60px; top:150px; width:100px; height:30px; z-index:5"><%=Sys_Owner%>　台啟</div>

<div id="Layer6" class="style1" style="position:absolute; left:60px; top:200px; width:250px; height:30px; z-index:5">應到案處所：<%=Sys_STATIONNAME%><br>應到案處所電話：<%=Sys_StationTel%></div>

<div id="Layer7" class="style1" style="position:absolute; left:360px; top:70px; width:200px; height:30px; z-index:5" align="center">大宗郵資已付掛號函件<br>第<%=Sys_MailNumber%>號</div>

<div id="Layer8" class="style1" style="position:absolute; left:360px; top:100px; width:200px; height:30px; z-index:5" align="center">
	<%="<img src=""..\BarCodeImage\"&Sys_BillNo&"_2.jpg"">"%><br><%=Sys_MAILCHKNUMBER%>
</div>

<div id="Layer9" class="style5" style="position:absolute; left:620px; top:100px; width:100px; height:30px; z-index:5" align="center">
	<%=left(theCity,2)%>郵局許可號碼<br><%=left(theCity,2)%>字第107號
</div>

<div id="Layer10" class="style4" style="position:absolute; left:270px; top:404px; width:200px; height:30px; z-index:5" align="center">
	<%=replace(thenPasserCity,"警察局","")%>
</div>

<%if showBarCode then%>
	<div id="Layer11" style="position:absolute; left:55px; top:435px; width:20px; height:30px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer12" style="position:absolute; left:55px; top:465px; width:20px; height:30px; z-index:5">Ｖ</div>
<%end if%>

<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer13" style="position:absolute; left:170px; top:440px; width:20px; height:30px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer14" style="position:absolute; left:170px; top:455px; width:20px; height:30px; z-index:5">Ｖ</div>
<%end if%>

<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer15" style="position:absolute; left:635px; top:415px; width:202px; height:30px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer16" style="position:absolute; left:635px; top:430px; width:202px; height:30px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer17" style="position:absolute; left:635px; top:445px; width:202px; height:30px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>

<div id="Layer18" style="position:absolute; left:10px; top:495px; width:200px; height:30px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_3.jpg"">"
	else
		response.write theUnitName
	end if
%></div>

<div id="Layer19" style="position:absolute; left:500px; top:480px; width:200px; height:30px; z-index:6"><%="<img src=""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"">"%></div>

<div id="Layer20" class="style1" style="position:absolute; left:500px; top:520px; width:200px; height:30px; z-index:6"><%=thePollic%>字第</div>

<div id="Layer21" class="style2" style="position:absolute; left:115px; top:545px; width:300px; height:30px; z-index:10"><%
	If Sys_UnitID<>"046A" then
		response.write "逕行舉發&nbsp;"
	else
		response.write "拖吊逕行舉發&nbsp;"
	end if
	response.write Sys_A_Name&"<br>"
	If Sys_UnitID<>"046A" then
		response.write "附採證照片&nbsp;"
	else
		response.write "　　　　　　"
	end if
%></span></div>

<div id="Layer22" class="style2" style="position:absolute; left:265px; top:540px; width:28px; height:30px; z-index:3"><%=Sys_Sex%></div>

<div id="Layer23" class="style1" style="position:absolute; left:360px; top:540px; width:400px; height:30px; z-index:4">＊本單可至全省統一、全家、萊爾富、OK等超商繳納。</div>

<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer24" class="style2" style="position:absolute; left:265px; top:565px; width:100px; height:30px; z-index:8"><%=Sys_DriverBirth(0)%></div>

<div id="Layer25" class="style2" style="position:absolute; left:295px; top:565px; width:100px; height:30px; z-index:8"><%=Sys_DriverBirth(1)%></div>

<div id="Layer26" class="style2" style="position:absolute; left:325px; top:565px; width:100px; height:30px; z-index:8"><%=Sys_DriverBirth(2)%></div>

<div id="Layer27" class="style2" style="position:absolute; left:430px; top:565px; width:106px; height:30px; z-index:9"><%=Sys_DriverID%></div>

<div id="Layer28" class="style2" style="position:absolute; left:630px; top:565px; width:99px; height:30px; z-index:10"><%=fastring%></div>
<%end if%>
<div id="Layer29" class="style2" style="position:absolute; left:115px; top:590px; width:100px; height:30px; z-index:11"><%=Sys_CarNo%></div>

<div id="Layer30" class="style2" style="position:absolute; left:265px; top:590px; width:120px; height:30px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>

<div id="Layer31" class="style2" style="position:absolute; left:500px; top:590px; width:150px; height:30px; z-index:13"><%=chstr(Sys_Owner)%></div>

<div id="Layer32" class="style2" style="position:absolute; left:115px; top:615px; width:500px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&chstr(Sys_OwnerZipName&Sys_OwnerAddress)%></div>

<div id="Layer33" class="style2" style="position:absolute; left:120px; top:635px; width:20px; height:30px; z-index:15"><%=Sys_IllegalDate(0)%></div>

<div id="Layer34" class="style2" style="position:absolute; left:170px; top:635px; width:20px; height:30px; z-index:16"><%=Sys_IllegalDate(1)%></div>

<div id="Layer35" class="style2" style="position:absolute; left:220px; top:635px; width:20px; height:30px; z-index:17"><%=Sys_IllegalDate(2)%></div>

<div id="Layer36" class="style2" style="position:absolute; left:265px; top:635px; width:20px; height:30px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>

<div id="Layer37" class="style2" style="position:absolute; left:315px; top:635px; width:20px; height:30px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>

<div id="Layer38" class="style2" style="position:absolute; left:395px; top:640px; width:300px; height:31px; z-index:20"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "該路段限速"&Sys_RuleSpeed&"公里、經雷達測速為"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if
		end if
	else
		If not ifnull(Sys_Rule4) Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")		
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	end if
	if trim(Sys_Rule2)<>"" then
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end if
			
	if not ifnull(Sys_Note) then response.write "("&Sys_Note&")"
%></div>
<div id="Layer39" class="style2" style="position:absolute; left:115px; top:660px; width:220px; height:20px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>

<div id="Layer40" class="style2" style="position:absolute; left:120px; top:690px; width:30px; height:20px; z-index:22"><%=Sys_DealLineDate(0)%></div>

<div id="Layer41" class="style2" style="position:absolute; left:210px; top:690px; width:30px; height:20px; z-index:23"><%=Sys_DealLineDate(1)%></div>

<div id="Layer42" class="style2" style="position:absolute; left:290px; top:690px; width:30px; height:20px; z-index:24"><%=Sys_DealLineDate(2)%></div>

<div id="Layer43" class="style1" style="position:absolute; left:405px; top:699px; width:400px; height:10px; z-index:29"><%
	response.write left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　"&Sys_Level1
%></div>

<div id="Layer44" class="style1" style="position:absolute; left:405px; top:700px; width:400px; height:10px; z-index:29"><%
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　0"&Sys_Level2
	end if
%></div>

<div id="Layer45" class="style1" style="position:absolute; left:395px; top:740px; width:90px; height:30px; z-index:28"><%=Sys_STATIONNAME%></div>

<div id="Layer46" style="position:absolute; left:480px; top:735px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""..\BarCodeImage\"&Sys_BillNo&"_5.jpg"">"
%></div>

<div id="Layer47" class="style6" style="position:absolute; left:420px; top:810px; width:100px; height:30px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	if Session("Unit_ID") <>"0207" then 
		response.write "<tr><td  class=""style6"" style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center"">&nbsp;"&Sys_UnitName&"&nbsp;<br>&nbsp;"&Sys_UnitTEL&"&nbsp;</td></tr>"
	end if
	response.write "</table>"
%></div>

<div id="Layer48" class="style6" style="position:absolute; left:600px; top:790px; width:100px; height:30px; z-index:30"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=2>"
	if Session("Unit_ID") <>"0207" then 
		response.write "<tr><td nowrap class=""style6"" style=""border-color:#ff0000;border-style:solid;"" align=""center"">主管職名章<br>"&Sys_JobName&"&nbsp;"&Sys_MainChName&"</td></tr>"
	end if
	response.write "</table>"
%></div>

<div id="Layer49" class="style6" style="position:absolute; left:600px; top:840px; width:100px; height:30px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""..\Member\Picture\"&Sys_MemberFilename&""" width=""90"" height=""30"">"
	else
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=2>"
		response.write "<tr><td class=""style6"" style=""border-color:#ff0000;border-style:solid;"" width=80 align=""center"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</td></tr>"
		response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
	end if
%></div>

<div id="Layer50" class="style2" style="position:absolute; left:240px; top:875px; width:60px; height:30px; z-index:32"><%=sys_Date(0)%></div>

<div id="Layer51" class="style2" style="position:absolute; left:375px; top:875px; width:60px; height:30px; z-index:33"><%=sys_Date(1)%></div>

<div id="Layer52" class="style2" style="position:absolute; left:525px; top:875px; width:60px; height:30px; z-index:34"><%=sys_Date(2)%></div>

<div id="Layer53" class="style2" style="position:absolute; left:685px; top:875px; width:80px; height:30px; z-index:36"><%=Sys_BillFillerMemberID%></div>

<div id="Layer54" class="style2" style="position:absolute; left:60px; top:1030px; width:200px; height:20px; z-index:36"><%=thenPasserCity%></div>

<div id="Layer55" class="style1" style="position:absolute; left:350px; top:1030px; width:200px; height:20px; z-index:36"><%=thenPasserCity&replace(theUnitName,trim(thenPasserCity),"")%></div>

<div id="Layer56" class="style1" style="position:absolute; left:570px; top:1030px; width:250px; height:20px; z-index:5"><%=theUnitAddress%></div>

<div id="Layer57" class="style5" style="position:absolute; left:130px; top:1050px; width:190px; height:80px; z-index:56"><%=chstr(Sys_Owner)&"<br>"&Sys_OwnerZip&" "&chstr(Sys_OwnerZipName&Sys_OwnerAddress)%></div>

<!--<div id="Layer58" class="style5" style="position:absolute; left:130px; top:1090px; width:280px; height:20px; z-index:46"><%=Sys_BillNo%></div>-->

<div id="Layer59" style="position:absolute; left:75px; top:1085px; width:200px; height:30px; z-index:36"><%="<img src=""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"">"%></div>
</div>
<%
	if (i mod 100)=0 then response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>