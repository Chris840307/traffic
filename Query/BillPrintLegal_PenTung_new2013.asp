<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title>舉發單列印-屏東縣 使用 Legal Size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體";font-size: 14px; color:#ff0000;}
.style2 {font-family:"標楷體";font-size: 20px; line-height:1;}
.style3 {font-family:"標楷體";font-size: 14px}
.style33{font-family:"標楷體";font-size: 14px}
.style4 {font-family:"標楷體";font-size: 12px}
.style7 {font-family:"標楷體";font-size: 12px}
.style8 {font-family:"標楷體";font-size: 12px}
.style10 {font-family:"標楷體";font-size: 14px; color:#ff0000; }
.style11 {font-family:"標楷體";font-size: 16px}
.style15 {font-family:"標楷體";font-size: 16px; line-height:1;}
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
Server.ScriptTimeout=6000

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select Value from ApConfigUre where ID=40"
set City=conn.execute(strSQL)
Sys_CityUnit=City("Value")
City.close

PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo,DCIerrorCarData from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_DCIerrorCarData=""
if Not rsbil.eof then Sys_DCIerrorCarData=trim(rsbil("DCIerrorCarData"))


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
strSql="select * from BillBase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverHomeAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverHomeZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
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
			chkaddress="(車)"
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))&chkaddress
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		End if

	else
		If ifnull(Sys_OwnerAddress) Then
			chkaddress="(戶)"
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverAddress"))&chkaddress
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverZip"))
		End if
	end if

	If ifnull(Sys_OwnerAddress) Then
		chkaddress="(戶)"
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverAddress"))&chkaddress
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverZip"))
	End if

	If not ifnull(Sys_OwnerAddress) Then
		strSQL="Update Billbase set Owner='"&rsfound("Owner")&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"',DriverZip='"&trim(rsfound("DriverZip"))&"',DriverAddress='"&trim(rsfound("DriverAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
		conn.execute(strSQL)
	end if
end if

if not ifnull(Sys_OwnerAddress) then
	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"  "," ")," ","　")
end If 

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

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
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,RECORDMEMBERID from BillBase where SN="&trim(rsbil("BillSN"))
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

strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFillerTel=trim(mem("Tel"))
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

If Sys_Rule1=5630001 Then
	sys_CityUnit=""
	SysUnit="屏東縣政府公有收費停車管理組"
	SysUnitAddress="屏東市自由路527號"
	SysUnitTel="08-7329767"
End if

if trim(Sys_UnitLevelID)="3" then
	chkJobID="303,304,314"
else
	chkJobID="303,304"
end if

strSQL="select a.ChName,b.Content,b.ID from (select UnitID,ChName,JobID from MemberData where AccountStateID=0 and RecordStateID=0 and UnitID='"&Sys_UnitID&"' and JobID in("&chkJobID&")) a,(select ID,Content from Code where TypeID=4 ) b where a.JobID=b.ID order by ID"
'response.write strSQL
set rsjob=conn.execute(strSQL)
if Not rsjob.eof then
	Sys_jobName=rsjob("Content")
	Sys_MainChName=rsjob("ChName")
end if
rsjob.close

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
end If 

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
if ifnull(Sys_MailNumber) then Sys_MailNumber=0

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode_PT PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,"904","018","17"
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	delphiASPObj.GenBillPrintBarCode_PT PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,"904","018","17"
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
pagepx=60
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer42" style="position:absolute; left:20px; top:27px;">
<table width="710" height="160" border="0" cellspacing=0 cellpadding=0>
	<!---------------------------------------- start  縣市抬頭, 地址, 電話. --------------------------------------------->
	<tr>
		<td>&nbsp;</td>
		<td class="style15"><b><%=SysUnitAddress&"<br>"&SysUnit& "  " & SysUnitTel%></b></td>
		<td>&nbsp;</td>
	</tr>
	<!---------------------------------------- 放大宗掛號    --------------------------------------------->
	
 <tr >
    <td>&nbsp;</td>
    <td  width="530" align="center"><br><%
		If Sys_UnitLevelID < 2 Then
			Response.Write "<p class=""style4""><font size=""2"">大宗郵資已付掛號函件<br>  第"&right("00000000" & trim(Sys_MailNumber),6)&"號  </font></p>"
		end if%>    </td>
    <td >&nbsp;</td>
    
  </tr>

  <tr>

	
    <td >&nbsp;</td>
    <td width="530" align="center"><%
		If Sys_UnitLevelID < 2 Then
			Response.Write "<div align=""center""><img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""></img><br><font size=""2"">"&Sys_MAILCHKNUMBER&"</font></div>"
		end if
	%></td>
	
     <!---------------------------------------- 放 許可證    --------------------------------------------->
     <!--
    		<td class="style2" >
			<span class="style4">
			
			雲林郵局許可號碼<br>
			雲林字第１０７號
			</span>      
		</td>
         <!---許可證的位置用Br控制高低位置-->		
  </tr>
  <tr>
	<td>
	</td>
  </tr>
	<!----------------------------------------  收件人資料. --------------------------------------------->
	<tr>
		<td width="110" height="2">&nbsp;</td>
		<td width="510" valign="TOP" class="style2"><b><font size="4"><%
			if trim(Sys_BillTypeID)="1" then
				response.write Sys_DriverHomeZip&"<br>"
				response.write funcCheckFont(Sys_DriverZipName&Sys_DriverHomeAddress,20,1)
			elseif trim(Sys_BillTypeID)="2" then
				response.write Sys_OwnerZip&"<br>"
				response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,20,1)
				'response.write "台北縣三重市承德里１１２鄰三民路１２３巷１２３號１２１之一"
			end if
			response.write "<br>"
			if trim(Sys_BillTypeID)="1" then
				response.write funcCheckFont(Sys_Driver,20,1)
			elseif trim(Sys_BillTypeID)="2" then
				response.write funcCheckFont(Sys_Owner,20,1)
			end if%>	　敬啟<br><br><br>
			
		</font></b></td> 
		<!--  監理站代碼
				
		<td><p class="style6"><%=Sys_StationID%></p></td> 
		
		-->

	</tr>
</table>
</div>
<!---------------------------------------- start 列印紅單紅色區域內容 --------------------------------------------->
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:48px; top:<%=345+pagepx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:48px; top:<%=360+pagepx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:165px; top:<%=350+pagepx%>px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:165px; top:<%=360+pagepx%>px; width:202px; height:36px; z-index:5">v</div>
<%end if%>
<div id="Layer9" style="position:absolute; left:20px; top:<%=410+pagepx%>px; width:210px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:<%=393+pagepx%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer42" style="position:absolute; left:535px; top:<%=460+pagepx%>px; width:233px; height:12px; z-index:36"><font size=2></font></div>

<div id="Layer12" style="position:absolute; left:105px; top:<%=485+pagepx%>px; width:300px; height:11px; z-index:20"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>4340003 and int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>
<div id="Layer13" style="position:absolute; left:270px; top:<%=490+pagepx%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:405px; top:<%=470+pagepx%>px; width:304px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div><%'=Sys_DriverHomeZip&" "&Sys_DriverZipName&Sys_DriverHomeAddress%>
<div id="Layer15" style="position:absolute; left:270px; top:<%=500+pagepx%>px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:455px; top:<%=490+pagepx%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:650px; top:<%=490+pagepx%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:105px; top:<%=530+pagepx%>px; width:100px; height:14px; z-index:11"><b><%=Sys_CarNo%></b></div>
<div id="Layer19" style="position:absolute; left:305px; top:<%=530+pagepx%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:540px; top:<%=523+pagepx%>px; width:201px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,20,1)%></div>
<div id="Layer21" style="position:absolute; left:175px; top:<%=555+pagepx%>px; width:520px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,20,1)%></div>

<div id="Layer22" style="position:absolute; left:110px; top:<%=585+pagepx%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:170px; top:<%=585+pagepx%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:225px; top:<%=585+pagepx%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:280px; top:<%=585+pagepx%>px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:335px; top:<%=585+pagepx%>px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:420px; top:<%=585+pagepx%>px; width:280px; height:31px; z-index:20"><span class="style33"><%
	If not ifnull(Sys_Jurgeday(0)) Then Response.Write "民眾檢舉案件，檢舉時間 "&Sys_Jurgeday(0)&"年"&Sys_Jurgeday(1)&"月"&Sys_Jurgeday(2)&"日<br>"

	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310219) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"
			if Sys_Level1<>"0" then response.write "(罰鍰新台幣"&Sys_Level1 &"元)"
			
			'if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
			'	response.write "<br>100以上"
			'elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
			'	response.write "<br>80以上未滿100"
			'elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
			'	response.write "<br>60以上未滿80"
			'elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
			'	response.write "<br>40以上未滿60"
			'elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
			'	response.write "<br>20以上未滿40"
			'else
			'	response.write "<br>未滿20公里"
			'end if
		end if
	else
		If trim(Sys_Rule4)<>"" Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if int(Sys_Rule1)=5620001 then	Sys_IllegalRule1=replace(Sys_IllegalRule1,"經催繳","")
		response.write Sys_IllegalRule1 
		if Sys_Level1<>"0" then response.write "(罰鍰新台幣"&Sys_Level1 &"元)"

	end if

	if trim(Sys_Rule2)<>"0" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		
		response.write "<br>"&Sys_IllegalRule2 
		if Sys_Level2<>"0" then response.write "(罰鍰新台幣"&Sys_Level2&"元)"
	end if
	if int(Sys_Rule1)=5620001 then response.write "("&Sys_Note&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:110px; top:<%=605+pagepx%>px; width:220px; height:15px; z-index:21"><span class="style33"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:140px; top:<%=630+pagepx%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:220px; top:<%=630+pagepx%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:300px; top:<%=630+pagepx%>px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<!-----------------------------------------法條編號 --------------------------------------------->
<div id="Layer32" style="position:absolute; left:520px; top:<%=672+pagepx%>px; width:400px; height:49px; z-index:29"><%
	response.write "<span class=""style8"">"&left(trim(Sys_Rule1),2)&"　　　　　"
	if len(trim(Sys_Rule1))>7 then response.write "　　　　　　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　　　　"&Mid(trim(Sys_Rule1),4,2)
	<!--response.write "　　　　　　　　　　　　　　"&Sys_Level1 ----->
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
	<!--	response.write "　　　　　　　　　　　　　"&Sys_Level2----->
	end if
	response.write "</span>"
%></div>

<div id="Layer33" style="position:absolute; left:420px; top:<%=710+pagepx%>px; width:70px; height:40px; z-index:28"><span class="style4"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></div>

<!Smith 2013------------監理站barcode 要在這個高度才刷得出來--->
<div id="Layer34" style="position:absolute; left:473px; top:<%=711+pagepx%>px; width:400px; height:30px; z-index:2"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>

<!----'smith                     舉發單位章	--->
<div id="Layer35" style="position:absolute; left:420px; top:<%=760+pagepx%>px; width:100px; height:89px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style3""><font color='red'>"
	If Sys_UnitLevelID > 2 Then
		Response.Write SysUnit&Sys_UnitName
	else
		Response.Write SysUnit
	End if
	Response.Write "</font></span><br><span class=""style3""><font color='red'>"
	If Sys_UnitLevelID > 2 Then
		Response.Write Sys_UnitFillerTel
	else
		Response.Write SysUnitTel
	End if	
	Response.Write "</font></span></td></tr>"
	response.write "</table>"
	If trim(Sys_DCIerrorCarData)="F" Then response.write "<B>繳註銷後案</B>"
%></div>
	
<div id="Layer36" style="position:absolute; left:580px; top:<%=775+pagepx%>px; width:100px; height:43px; z-index:30"><%
	'if instr(Sys_BillNo,"QZ")>0 then
	'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>分隊長&nbsp;</span></td></tr>"
	'		response.write "</table>"
	'elseif trim(Session("Unit_ID"))="TO00" then
		'礁溪分局 警備隊的話要警備隊隊長 其他就用組長
	'	if Sys_UnitID="TOUD" then
	'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
	'		response.write "</table>"
	'	else
	'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>組長&nbsp;</span></td></tr>"
	'		response.write "</table>" 
	'	end if
	''else
	'	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	'	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
	'	response.write "</table>"
	'end if
	%></div>
<div id="Layer37" style="position:absolute; left:610px; top:<%=780+pagepx%>px; width:200px; height:46px; z-index:31"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
	response.write "</table>"		
%></div>

<!-----------------------------------填單日---------------------------------->
<div id="Layer38" style="position:absolute; left:130px; top:<%=810+pagepx%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:185px; top:<%=810+pagepx%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:245px; top:<%=810+pagepx%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<!-----------------------------------檢查用車號---------------------------------->
<div id="Layer41" style="position:absolute; left:355px; top:<%=930+pagepx%>px; width:90px; height:11px; z-index:34"><%=Sys_CarNo%></div>
<!-----------------------------------送達證書 繳回單位資料---------------------------------->
<div id="Layer42" class="style3" style="position:absolute; left:28px; width:180px; top:<%=956+pagepx%>px; z-index:1"><%=Sys_CityUnit&replace(SysUnit,Sys_CityUnit,"")%></div>

<div id="Layer45" class="style3" style="position:absolute; left:340px; width:180px; top:<%=956+pagepx%>px; z-index:1"><%=Sys_CityUnit&replace(SysUnit,Sys_CityUnit,"")%></div>

<div id="Layer46" class="style3" style="position:absolute; left:555px; width:250px; top:<%=959+pagepx%>px; z-index:1"><%=SysUnitAddress%></div>

<!-----------------------------------送達證書 收件人資料---------------------------------->
<div id="Layer47" class="style3" style="position:absolute; left:115px; top:<%=975+pagepx%>px; z-index:1"><%
	response.write funcCheckFont(Sys_Owner,20,1)&"<br>"
	If instr(Sys_OwnerAddress,"@") >0 Then
		response.write funcCheckFont(Sys_OwnerZip&Sys_OwnerZipName&Sys_OwnerAddress,20,1)
	else
		response.write InstrAdd(Sys_OwnerZip&Sys_OwnerZipName&Sys_OwnerAddress,14)
	End if
%>
</div>
<!-----------------------------------送達證書 文 號 ---------------------------------->
<div id="Layer48" class="style3" style="position:absolute; left:225px; top:<%=1020+pagepx%>px; z-index:1"><%=Sys_BillNo%>
</div>

<!--
<div id="Layer48" style="position:absolute; left:400px; top:<%=1030+pagepx%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
-->

<div id="Layer49" class="style3" style="position:absolute; left:510px; top:<%=1035+pagepx%>px; z-index:1"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
</div>
</div><%
	if (i mod 50)=0 then response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<!-----------------------------------------------------------  設定印表機邊界 ---------------------------------------------------------------------------->
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,5.08,5.08,5.08,5.08);
</script>