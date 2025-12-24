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
<style type="text/css">
<!--
.style1 {font-family:"標楷體";font-size: 14px; color:#ff0000;}
.style2 {font-size: 12px}
.style3 {font-size: 14px}
.style4 {font-family:"標楷體";font-size: 10px; color:#ff0000;}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style12 {font-family:"標楷體"; font-size: 8px;}
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style13 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style14 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style15 {font-family:"標楷體"; font-size: 20px;}
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsxie8.cab#Version=6,5,439,50">
</object>
<%
'on Error Resume Next
Server.ScriptTimeout=6000

PBillSN=split(trim(request("BillSN")),",")

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

for i=0 to Ubound(PBillSN)

if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"

strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

strSQL="select count(1) cnt from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))
set rsSend=conn.execute(strSQL)

If cdbl(rsSend("cnt")) = 0 Then

	strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,MailDate,SendMailStation,ArriveType,ReturnResonID,Note) values((select nvl(Max(SN),0)+1 as cnt from PassersEndArrived),"&trim(rsbil("BillSN"))&",sysdate,"&Session("User_ID")&","&Session("User_ID")&",sysdate,MailNumber_Sn.NextVal,2,null,null)"

	conn.execute(strSQL)

End if 
rsSend.close

strSql="select * from PasserBase where SN="&trim(rsbil("BillSN"))

set rs=conn.execute(strSql)

Sys_BillUnitID="":Sys_RecordMemberID="":Sys_BillFillerMemberID=""
Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip=""
Sys_Illegaladdress="":Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Note=""
Sys_Rule1="":Sys_Rule2="":Sys_Level1=0:Sys_Level2=0:Sum_Level=0
Sys_RuleVer="":Sys_MailDate=""

if Not rs.eof then 

	Sys_BillUnitID=trim(rs("BillUnitID"))
	Sys_RecordMemberID=trim(rs("RecordMemberID"))
	Sys_BillFillerMemberID=trim(rs("BillFillerMemberID"))
	Sys_Owner=trim(rs("Driver"))
	Sys_OwnerAddress=trim(rs("DriverAddress"))
	Sys_OwnerZip=trim(rs("DriverZip"))
	Sys_Illegaladdress=trim(rs("ILLEGALADDRESS"))
	Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
	Sys_RuleSpeed=trim(rs("RuleSpeed"))
	Sys_Note=trim(rs("Note"))
	Sys_RuleVer=trim(rs("RuleVer"))

	Sys_DCIReturnStation=trim(rs("MemberStation"))

	Sys_BillNo=trim(rs("BillNo"))
	Sys_CarNo=trim(rs("CarNo"))

	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))

	If isnumeric(rs("FORFEIT1")) Then Sys_Level1=rs("FORFEIT1")
	If isnumeric(rs("FORFEIT2")) Then Sys_Level2=rs("FORFEIT2")

	Sum_Level=Sys_Level1+Sys_Level2

	Sys_DCIRETURNCARTYPE="微型電動二輪車"

	Sys_MailDate=trim(rs("BillFillDate"))


end If 

if Not rs.eof then
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_h=""
end If 

if Not rs.eof then
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_m=""
end If 

if Not rs.eof then
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if

if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end If 



rs.close

If ifnull(Sys_OwnerAddress) Then
	strSQL="select Owner,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and Status='S'"
	
	set rsfi=conn.execute(strSql)

	if Not rsfi.eof then
		If Not ifnull(trim(rsfi("Owneraddress"))) Then

			Sys_Owner=trim(rs("Owner"))
			Sys_OwnerAddress=trim(rs("OwnerZip"))
			Sys_OwnerZip=trim(rs("Owneraddress"))
			
			strSQL="update passerbase set" & _
			" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
			",DriverAddress='"&trim(rsfi("Owneraddress"))&"(車)'" & _
			" where SN="&trim(rsbil("BillSN"))

			conn.execute(strSQL)
		end if
	end If 
	rsfi.close
end If 



strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_OwnerAddress=replace(replace(Sys_OwnerAddress&"","臺","台"),Sys_OwnerZipName,"")


strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,c.Content from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.jobid=c.id(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
if Not mem.eof then Sys_BillFillerJobName=mem("Content")
mem.close

strSQL="select a.UnitID,a.UnitName,a.UnitTypeID,a.UnitLevelID,b.chName,c.Content from Unitinfo a,Memberdata b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and b.jobid=c.id(+) and b.MemberID="&trim(Sys_RecordMemberID)
set Unit=conn.execute(strSQL)
Sys_UnitID=Unit("UnitID")
Sys_RedUnitName=Unit("UnitName")
Sys_UnitTypeID=Unit("UnitTypeID")
Sys_UnitLevelID=Unit("UnitLevelID")
Sys_RecordName=Unit("chName")
Sys_RecordJobName=Unit("Content")
Unit.close

If Sys_UnitLevelID=1 or instr(Sys_RedUnitName,"分隊")>0 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysAddress=Unit("Address")
SysUnitTel=Unit("Tel")
Unit.close


Sys_IllegalRule1=""

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then

	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end If 

Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end If 

Sys_STATIONNAME="":Sys_StationTel="":StationID="":theBankAccount="":theBankName=""

strSql="select UnitID,UnitName,Tel,BankName,BankAccount from Unitinfo where unitid='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_STATIONNAME=trim(rs("UnitName"))
if Not rs.eof then Sys_StationTel=trim(rs("Tel"))
if Not rs.eof then Sys_StationID=trim(rs("UnitID"))
if Not rs.eof then theBankAccount=trim(rs("BankAccount"))
if Not rs.eof then theBankName=trim(rs("BankName"))
rs.close

Sys_CityUnit="屏東縣政府警察局"

imgfile="":arrfile=""
Sys_IisImagePath="":Sys_ImageFileNameA="":Sys_ImageFileNameB=""
strSQL="select IisImagePath,ImageFileNameA ImageFileName1,ImageFileNameB ImageFileName2,ImageFileNameC ImageFileName3 from PasserIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	For k=1 To 3
		If Trim(rsimage("ImageFileName"&k))<>"" Then

			If imgfile<>"" Then imgfile=imgfile&"@"
			imgfile=imgfile&rsimage("ImageFileName"&k)
		End if
	Next

	If imgfile<>"" Then

		arrfile=Split(imgfile,"@")

		Sys_ImageFileNameA=trim(arrfile(0))

		If InStr(imgfile,"@") > 0 Then Sys_ImageFileNameB=trim(arrfile(1))

	End if
end If 

rsimage.close

Sys_MailNumber=0:Sys_MAILCHKNUMBER=0

strSQL="select min(SendMailStation) SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))

set rsSend=conn.execute(strSQL)

if Not rsSend.eof then
	Sys_MailNumber=rsSend("SendMailStation")
	Sys_MAILCHKNUMBER=Sys_MailNumber&" 905 018 17"
end If 
rsSend.close

DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,905,018,17

rsbil.close

pageTop=0
pageLeft=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" style="position:relative;">

<div id="Layer44" class="style2" style="position:absolute; left:160px; top:0px; height:12px; z-index:36"><B><%=Sys_CityUnit&replace(SysUnit,Sys_CityUnit,"")&"送達證書&nbsp;&nbsp;請繳回："&SysAddress%></B></div>

<div id="Layer01" class="style3" style="position:absolute; left:125px; top:20px; z-index:3"><%
	'response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp; "&Sys_CarNo&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress
%>
</div>


<div id="Layer02" class="style3" style="position:absolute; left:300px; top:60px; z-index:2"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:95px; top:50px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:390px; top:275px; z-index:1"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:120px; top:290px; z-index:1"><b><%
	response.write funcCheckFont(Sys_Owner,16,1)
	%>　台啟</b>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:120px; top:310px; width:370px; z-index:5"><b><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			response.write Sys_OwnerZip&" "
			response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress
			response.write "<br><br>"
	%></b>
</div>

<div id="Layer45" class="style3" style="position:absolute; left:120px; top:330px; height:12px; z-index:1"><b><%=Sys_CityUnit&replace(SysUnit,Sys_CityUnit,"")%></b></div>

<div id="Layer05" class="style3" style="position:absolute; left:435px; top:287px; z-index:3">
	<%If cdbl(Sys_MailNumber) <> 0 Then%>
		
	　<b>第<%=Sys_MailNumber%>號</b><br>
	　<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>><br>
    　<b><%=Sys_MAILCHKNUMBER%></b>

	<%End if %>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:45px; top:370px; z-index:5"><%="寄件地址："&SysAddress%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:235px; top:345px; z-index:1"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer43" style="position:absolute; left:320px; top:395px; width:250px; height:12px; z-index:36"><%=Sys_DCIRETURNCARTYPE%></div>

<%
if trim(Sys_ImageFileNameA)<>"" then
	Response.Write "<div id=""Layer09"" style=""position:absolute; left:35px; top:480px; z-index:5"">"
	response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	Response.Write "</DIV>"

End if

if trim(Sys_ImageFileNameB)<>"" then
	Response.Write "<div id=""Layer10"" style=""position:absolute; left:425px; top:480px; z-index:1"">"
	response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	Response.Write "</DIV>"
end If 
%>

<div id="Layer2" style="position:absolute; left:40px; top:815px; width:202px; height:36px; z-index:5">Ｖ</div>

<div id="Layer4" style="position:absolute; left:160px; top:800px; width:202px; height:36px; z-index:5">v</div>

<div id="Layer9" style="position:absolute; left:40px; top:845px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&Sys_CityUnit&replace(SysUnit,Sys_CityUnit,"")
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:835px; width:233px; height:32px; z-index:3"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer12" style="position:absolute; left:110px; top:910px; width:250px; height:11px; z-index:6"><span class="style7"><%
		response.write "逕行舉發&nbsp;"
%></span></div>

<div id="Layer13" style="position:absolute; left:260px; top:905px; width:28px; height:11px; z-index:3"><%'=Sys_Sex%></div>
<div id="Layer14" class="style10" style="position:absolute; left:370px; top:905px; width:324px; height:10px; z-index:4"><%
	'if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"
	'Response.Write "＊受通知人收受通知單時應到案日期不足30日或已逾應到案日期者，得於送達生效日後30日內到案。"
%></div>

<div id="Layer15" class="style3" style="position:absolute; left:260px; top:915px; width:100px; height:10px; z-index:8"><%'if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></div>
<div id="Layer16" class="style3" style="position:absolute; left:425px; top:915px; width:106px; height:13px; z-index:9"><%'=Sys_DriverID%></div>
<div id="Layer17" class="style3" style="position:absolute; left:620px; top:915px; width:99px; height:12px; z-index:10"><%'=fastring%></div>
<div id="Layer18" class="style3" style="position:absolute; left:125px; top:965px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" class="style3" style="position:absolute; left:310px; top:965px; width:250px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style3" style="position:absolute; left:550px; top:965px; width:200px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" class="style3" style="position:absolute; left:180px; top:990px; width:610px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)&chkaddress%></div>

<div id="Layer22" class="style3" style="position:absolute; left:115px; top:1020px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style3" style="position:absolute; left:175px; top:1020px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style3" style="position:absolute; left:225px; top:1020px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style3" style="position:absolute; left:285px; top:1020px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" class="style3" style="position:absolute; left:335px; top:1020px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style3" style="position:absolute; left:420px; top:1020px; width:300px; height:31px; z-index:20"><%

	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then

			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"

		end If 
	else
		
		response.write Sys_IllegalRule1
	end If 

	Response.Write "(罰鍰新臺幣"&Sys_Level1&"元整)"

	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
		
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"

			Response.Write "(罰鍰新臺幣"&Sys_Level2&"元整)"

		end If 

	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
		Response.Write "(罰鍰新臺幣"&Sys_Level2&"元整)"
	end If 
%></div>
<div id="Layer28" class="style3" style="position:absolute; left:115px; top:1040px; width:220px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" class="style3" style="position:absolute; left:140px; top:1065px; width:50px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style3" style="position:absolute; left:220px; top:1065px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style3" style="position:absolute; left:300px; top:1065px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style3" style="position:absolute; left:520px; top:1110px; width:400px; height:49px; z-index:29"><%
	response.write left(trim(Sys_Rule1),2)&"　　　　　"
	'if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　　　　"&Mid(trim(Sys_Rule1),4,2)
	'response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　　　　"
		'if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　　　　"&Mid(trim(Sys_Rule2),4,2)
		'response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
%></div>

<div id="Layer33" style="position:absolute; left:450px; top:1155px; width:200px; height:80px; z-index:28"><span class="style7"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></span></div>

<div id="Layer35" style="position:absolute; left:415px; top:1200px; width:200px; height:49px; z-index:29"><%

	response.write "<table style=""border-color:#ff0000;border-style:solid; width:120px;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style3""><font color='red'>"
	If Sys_UnitLevelBillID > 2 Then
		Response.Write SysUnit&"<br>"&Sys_UnitName
	else
		Response.Write SysUnit
	End if
	Response.Write "</font></span><br><span class=""style3""><font color='red'>"
	If Sys_UnitLevelBillID > 2 Then
		Response.Write Sys_UnitTel
	else
		Response.Write SysUnitTel
	End if	
	Response.Write "</font></span></td></tr>"
	response.write "</table>"

%></div>

<div id="Layer37" style="position:absolute; left:605px; top:1220px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
	else
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillFillerJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
		response.write "</table>"	
	end if
%></div>

<div id="Layer38" class="style3" style="position:absolute; left:130px; top:1250px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" class="style3" style="position:absolute; left:195px; top:1250px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" class="style3" style="position:absolute; left:250px; top:1250px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>

</div>

</div>

<%
		if (i mod 10)=0 then response.flush
next
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo
If not ifnull(errBillNo2) Then errBillNo2="下列地址為國外\n"&errBillNo2
%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	<%
	If Not ifnull(errBillNo) Then
		respnse.write "alert("""&errBillNo&""");"
	end if 
	%>
	
	<%
	If Not ifnull(errBillNo2) Then
		respnse.write "alert("""&errBillNo2&""");"
	end if
	%>
	window.print();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>