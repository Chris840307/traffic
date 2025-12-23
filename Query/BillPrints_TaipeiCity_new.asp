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
.style5 {font-family:"標楷體";font-size: 15px;}
.style6 {font-family:"標楷體";font-size: 22px; line-height:2;}
.style8 {font-size: 36px}
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
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
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=6000
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close
PrintSum=request("hd_PrintSum"):Sys_Print=request("Sys_Print")
PrintSum=PrintSum-Sys_Print
if trim(Sys_Print)="" then
	PrintSum=0:Sys_Print=Ubound(PBillSN)+1
elseif Cint(request("hd_PrintSum"))-1>Ubound(PBillSN) then
	Sys_Print=Ubound(PBillSN)-(Cint(request("hd_PrintSum"))-Sys_Print)
end if
for i=0 to Cint(Sys_Print-1)
sumCnt=sumCnt+1
if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+PrintSum)&" and ExchangetypeID='W'"
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
'================
strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
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
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
'response.write strSql
'response.end
set rsfound=conn.execute(strSql)
'if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver"))
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

Sum_Level=0
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
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where SN="&trim(PBillSN(i+PrintSum))
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

chkJobID=""

	'Levle3 派出所 抓 所長jobid , 警備隊抓 隊長jobid , 交通分隊 抓 分隊長 jobid
	'       公園所 抓 巡佐
	if trim(Sys_UnitLevelID)="3" or trim(Sys_UnitLevelID)="2" then
		'交通分隊 抓 分隊長 jobid
		if trim(Sys_UnitID)="041S" or trim(Sys_UnitID)="042S"  or trim(Sys_UnitID)="044S" or trim(Sys_UnitID)="045S" or trim(Sys_UnitID)="048S" or trim(Sys_UnitID)="043S" then	
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
			chkJobID="304 and memberid=8144"
		'交通隊直屬分隊(分隊長)
		elseif trim(Sys_UnitID)="0469" then 
			chkJobID="304 "
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
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,36
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
<div id="L78" style="position:relative;">
<div style="position:absolute; left:0px; top:80px;">
<table width="645" height="393" border="0">
  <tr>
    <td width="141" height="69" valign="top">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
    <td rowspan="2" align="right" valign="top" nowrap><span class="style6"></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td height="41" align="left" valign="top">　　　　<img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_1.jpg"""%> hspace="0" vspace="0" align="top"><br>　　　　<span class="style3"><%=Sys_FirstBarCode%></span>
	</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <td colspan="2" align="left" valign="top" nowrap><span class="style3"><%=Sys_DriverHomeZip%><br>
    <%if len(Sys_DriverZipName&Sys_DriverHomeAddress)<21 then
			response.write Sys_DriverZipName&Sys_DriverHomeAddress
		else
			response.write left(Sys_DriverZipName&Sys_DriverHomeAddress,20)&"<br>"&mid(Sys_DriverZipName&Sys_DriverHomeAddress,21,len(Sys_DriverZipName&Sys_DriverHomeAddress))
		end if%></span></td>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<td align="left" valign="top" nowrap><span class="style3"><%=Sys_OwnerZip%><br>
    <%if len(Sys_OwnerZipName&Sys_OwnerAddress)<21 then
			response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,25,1)
		else
			response.write funcCheckFont(left(Sys_OwnerZipName&Sys_OwnerAddress,20)&"<br>"&mid(Sys_OwnerZipName&Sys_OwnerAddress,21,len(Sys_OwnerZipName&Sys_OwnerAddress)),25,1)
		end if%></span></td>
	<%end if%>
  </tr>
  <tr>
    <td>&nbsp;</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <td colspan="2"><span class="style3"><%=chstr(Sys_Driver)%>　台啟</span></td>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<td width="222"><span class="style3"><%=funcCheckFont(Sys_Owner,25,1)%>　台啟　　　　<%=Sys_CarNo%></span></td>
	<%end if%>
    <td width="92" align="right">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="145" align="center"><p class="style4">&nbsp;
   	    </p>
      <p class="style4">大宗郵資已付掛號函件<br>
    第<%=Sys_MailNumber%>號  </p>    </td>
    <td width="23" align="center">&nbsp;</td>
    <td align="right" nowrap>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"><div align="left"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_2.jpg"""%>><br>
        <%=Sys_MAILCHKNUMBER%></div></td>
    <td align="center">&nbsp;</td>
    <td align="right"><p>&nbsp;</p>
    <p class="style8"></p></td>
  </tr>
  <tr>
    <td height="98" valign="top" nowrap><p>　　　　<span class="style5">應到案處所：<%=Sys_STATIONNAME%></span><br>
   	　　　　<span class="style5">應到案處所電話：<%=Sys_StationTel%></span><br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;
		<span class="style4"><%=Sys_UnitName%></span></p>
    <p>&nbsp;</p></td>
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</div>
<!--
<%'if trim(Sys_IMAGEFILENAMEB)<>"" then%>
<div style="position:absolute; left:60px; top:600px;"><img src="img.JPG" width="550"></DIV>
<%'end if%>
-->
<%if trim(Sys_IMAGEFILENAME)<>"" then%>
<div style="position:absolute; left:360px; top:485px;"><img src="<%=Sys_IMAGEPATHNAME&Sys_IMAGEFILENAME%>" width="365" height="265"></DIV>
<%end if%>
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
	printWindow(true,5.08,5.08,5.08,5.08);
</script>