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
.style5 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style6 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style7 {font-size: 13px}
.style8 {font-size: 36px}
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
<!--
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,5,439,72">
</object>
-->
<%
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,ILLEGALADDRESS,IllegalSpeed,RuleSpeed,INSURANCE,RuleVer,Note,Rule4,BillFillDate from BillBase where SN="&trim(rsbil("BillSN"))
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
if Not rs.eof then Sys_Rule4=trim(rs("Rule4"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerZip="":Sys_OwnerAddress="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))

strSql="select DriverHomeZip,DriverHomeAddress,OwnerNotifyAddress from BillbaseDCIReturn where Carno=(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"

set rs1=conn.execute(strSQL)

if not rs1.eof then


	If not ifnull(rs1("OwnerNotifyAddress")) Then
		if Not rs1.eof then Sys_OwnerAddress=trim(rs1("OwnerNotifyAddress"))

	elseIf Not ifnull(trim(rsfound("OwnerAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	
	else
		if Not rs1.eof then Sys_OwnerAddress=trim(rs1("DriverHomeAddress"))
		if Not rs1.eof then Sys_OwnerZip=trim(rs1("DriverHomeZip"))

	End if
else
	If Not ifnull(trim(rsfound("OwnerAddress"))) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

	else
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	End if
end if

rs1.close

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If not ifnull(Sys_OwnerZipName) Then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
End if

If not ifnull(Sys_OwnerAddress) Then
		strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"
		conn.execute(strSQL)
end if

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&rsbil("BillNo")&","&Sys_Owner&"\n"
end if

Sys_DCIReturnStation=0
Sum_Level=0
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
If Sys_BillFillerMemberID = "3480" Then
	Stop_IllegalDate=split(gArrDT(trim(rssex("IllegalDate"))),"-")
	Stop_IllegalDate_h=hour(trim(rssex("IllegalDate")))
	Stop_IllegalDate_m=minute(trim(rssex("IllegalDate")))

	strSQL="select DealLineDate from BillBase where CarNo='"&Sys_CarNo&"' and IllegalDate=to_date('"&gOutDT(gInitDT(rssex("IllegalDate")))&" "&hour(trim(rssex("IllegalDate")))&":"&minute(trim(rssex("IllegalDate")))&":00','YYYY/MM/DD/HH24:MI/SS') and ImageFileNameB is not null"

	set rsstop=conn.execute(strSQL)
	if Not rsstop.eof then Sys_IllegalDate=split(gArrDT(DateAdd("d",1,rsstop("DealLineDate"))),"-")
	if Not rsstop.eof then Sys_IllegalDate_h="00"
	if Not rsstop.eof then Sys_IllegalDate_m="00"
	rsstop.close
End if

strSql="select a.LoginID,a.ChName,b.UnitName,decode(b.Unitid,'Z000','A000',b.Unitid) Unitid,b.UnitTypeID,decode(b.Unitid,'Z000','1',b.UnitLevelID) UnitLevelID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerLogInID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
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

if trim(Sys_BillTypeID)="1" then
	If ifnull(Sys_MailNumber) Then Sys_MailNumber="0"
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,0,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	DelphiASPObj.CreateBarCode Sys_MailNumber&"97000718",128,35,260
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	If ifnull(Sys_MailNumber) Then Sys_MailNumber=0	
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"97000718","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

	DelphiASPObj.CreateBarCode Sys_MailNumber&"97000718",128,60,150
end if

Sys_FirstBarCode=Sys_BillNo

strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
Sys_MAILCHKNUMBER=Sys_MailNumber&"97000718"
If Sys_OwnerZip="001" then Sys_OwnerZip=""
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer42" style="position:absolute; left:40px; top:1px;">
<table width="710" height="160" border="0">

  <tr height="12">
		<td>&nbsp;</td>
		<td rowspan=3 align="center">
			<%If Sys_MailNumber <> 0 Then%>
				<%="<img src=""../BarCodeImage/"&Sys_MailNumber&"97000718"&".jpg"">"%><br>
				<%="　"&Sys_MAILCHKNUMBER%>
			<%end if%>
		</td>
		<td rowspan=3>
			<br><br><br><br><%if int(Sys_Rule1)=5620001 then Response.Write "(郵簡)"%>
			<br>
			<%if trim(Request("Sys_BatchNumber"))="99W17526" then Response.Write "ＴＧ"%>
		</td>
  </tr>
  <tr>
    <td><img  src=<%="""..\BarCodeImage\"&Sys_BillNo&"_1.jpg"""%>><br>　　<%=Sys_FirstBarCode%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="160" height="2">&nbsp;</td>
    <td valign="TOP"><%
			if trim(Sys_BillTypeID)="1" then
				response.write Sys_DriverHomeZip&"<br><br>"
				response.write Sys_DriverZipName&Sys_DriverHomeAddress
			elseif trim(Sys_BillTypeID)="2" then
				response.write Sys_OwnerZip&"<br><br>"
				response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,1)
			end if
			response.write "<br><br>"
			if trim(Sys_BillTypeID)="1" then
				response.write Sys_Driver
			elseif trim(Sys_BillTypeID)="2" then
				response.write funcCheckFont(Sys_Owner,20,1)
			end if%>	　台啟
	 </td>
	 <td valign="bottom"><%=Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%></td>
  </tr>
	

</table>
</div>

<!--<div id="Layer6" style="position:absolute; left:40px; top:540px; width:400px; height:36px; z-index:5"><span class="style7">查詢電話：<%=DB_UnitTel%><%'=Sys_UnitTel%>（<%=DB_UnitName%><%'=Sys_UnitName%>）</span></div>-->
<div id="Layer44" class="style1" style="position:absolute; left:30px; top:525px; width:320px; height:36px; z-index:5">
	自103年3月31日起，不服舉發者，應於接獲本單30日內，向處罰機關(機應到案處所)陳述；
	受處罰人於自動繳納後，若不服舉發事實者，仍得於繳納罰鍰30日內向處罰機關陳述。
</div>
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:40px; top:575px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:40px; top:610px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>

<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:160px; top:580px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:160px; top:595px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<!--<div id="Layer5" style="position:absolute; left:165px; top:845px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer6" style="position:absolute; left:625px; top:610px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer7" style="position:absolute; left:625px; top:625px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer8" style="position:absolute; left:625px; top:640px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:10px; top:630px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		'response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:490px; top:625px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer12" style="position:absolute; left:110px; top:680px; width:300px; height:11px; z-index:10"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>4340003 and int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>

<div id="Layer13" style="position:absolute; left:260px; top:675px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:370px; top:675px; width:324px; height:10px; z-index:10"><%'=Sys_DriverHomeZip&" "&Sys_DriverZipName&Sys_DriverHomeAddress%><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>
<div id="Layer15" style="position:absolute; left:260px; top:690px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:425px; top:690px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:620px; top:690px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:125px; top:725px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:260px; top:725px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:500px; top:725px; width:201px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,20,1)%></div>
<div id="Layer21" style="position:absolute; left:125px; top:750px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,20,1)%></div>

<div id="Layer22" style="position:absolute; left:120px; top:770px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:170px; top:770px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:220px; top:770px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:260px; top:770px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:320px; top:770px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:390px; top:775px; width:610px; height:31px; z-index:20"><span class="style3"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if
		end if
	else
		
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1
		if int(Sys_Rule1)=5620001 then	Sys_IllegalRule1=Sys_IllegalRule1&"(掛號催繳通知補繳逾7日期限)"
		If trim(Sys_Rule4)<>"" Then Sys_IllegalRule1=trim(Sys_IllegalRule1&"("&Sys_Rule4&")")
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	end if
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end if

'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:115px; top:795px; width:220px; height:15px; z-index:21"><span class="style3"><%
	Response.Write Sys_ILLEGALADDRESS
	If Sys_BillFillerMemberID = "3480" Then
		Response.Write "<br>"
		Response.Write "("&Stop_IllegalDate(0)&"年"&Stop_IllegalDate(1)&"月"&Stop_IllegalDate(2)&"日"
		Response.Write right("00"&Stop_IllegalDate_h,2)&"時"&right("00"&Stop_IllegalDate_m,2)&"分)"
	end if
%></span></div>
<div id="Layer29" style="position:absolute; left:130px; top:825px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:210px; top:825px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:280px; top:825px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" style="position:absolute; left:395px; top:835px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:370px; top:875px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer33" style="position:absolute; left:620px; top:880px; width:100px; height:40px; z-index:28"><span class="style7"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>
<div id="Layer35" style="position:absolute; left:400px; top:940px; width:100px; height:49px; z-index:29"><%
	if billprintuseimage=1 then
		response.write "<table border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td class=""style1"">花蓮縣警察局<br>"&sysunit&"</td></tr>"
		response.write "<tr><td class=""style1"">TEL"&Sys_UnitTel&"</td></tr>"
		response.write "</table>"

		'if trim(Sys_UnitFilename)<>"" then	response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""110"" height=""80"">"
		'response.write "<img src=""unit3.jpg"">"
	end if%></div>
<div id="Layer36" style="position:absolute; left:560px; top:970px; width:100px; height:43px; z-index:30"><%

%></div>
<div id="Layer37" style="position:absolute; left:670px; top:935px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""..\Member\Picture\"&Sys_MemberFilename&""" width=""60"" height=""20"">"
	end if
	response.write Sys_BillFillerLogInID
%></div>
<div id="Layer42" style="position:absolute; left:560px; top:1020px; width:200px; height:46px; z-index:31"><%

%></div>
<div id="Layer38" style="position:absolute; left:210px; top:1015px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:365px; top:1015px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:515px; top:1015px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:500px; top:1040px; width:200px; height:12px; z-index:36"><B><font size=4><%=Sys_CarNo%></font></B></div>
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
	window.print();
	//printWindow(true,5.08,5.08,5.08,5.08);
</script>