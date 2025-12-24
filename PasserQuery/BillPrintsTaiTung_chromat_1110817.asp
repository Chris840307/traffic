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
'on Error Resume Next

Server.ScriptTimeout=6000

PBillSN=split(trim(request("BillSN")),",")

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

for i=0 to Ubound(PBillSN)

if cint(i)>0 then response.write "<div class=""PageNext"">　</div>"

strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

strSQL="select count(1) cnt from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))
set rsSend=conn.execute(strSQL)

If cdbl(rsSend("cnt")) = 0 Then

	strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,SendMailStation,ArriveType,ReturnResonID,Note) values((select nvl(Max(SN),0)+1 as cnt from PassersEndArrived),"&trim(rsbil("BillSN"))&",sysdate,"&Session("User_ID")&","&Session("User_ID")&",null,2,null,null)"

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

If ifnull(Sys_OwnerAddress) Then

	strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
	rszip.close

	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"臺","台"),Sys_OwnerZipName,"")
end if

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

sys_title="臺東縣政府警察局"

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

Sys_MailNumber=0

DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95100017","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

firstBacrCode=right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&"D"&Sys_StationID

rsbil.close
%>

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

<div id="Layer02" class="style3" style="position:absolute; left:80px; top:45px; z-index:4"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:290px; top:60px; z-index:8"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:405px; top:270px; z-index:10"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:130px; top:285px; z-index:8"><%
	response.write funcCheckFont(Sys_Owner,16,1)
	%>　台啟
</div>

<div id="Layer05" class="style3" style="position:absolute; left:350px; top:330px; z-index:4"><%
	Response.Write "<img  src=""../BarCodeImage/"&Sys_BillNo&"_1.jpg"">"%>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:110px; top:309px; width:330px; z-index:8"><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			response.write Sys_OwnerZip&" "
			response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
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

<%if trim(Sys_ImageFileNameA)<>"" then%>
	<div id="Layer09" style="position:absolute; left:38px; top:480px; z-index:8"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	%></DIV>
<%End If 

if trim(Sys_ImageFileNameB)<>"" then%>
	<div id="Layer10" style="position:absolute; left:430px; top:485px; z-index:8"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>

<div id="Layer2" style="position:absolute; left:50px; top:840px; width:202px; height:36px; z-index:8">Ｖ</div>

<div id="Layer4" style="position:absolute; left:170px; top:825px; width:202px; height:36px; z-index:8">v</div>

<div id="Layer9" style="position:absolute; left:20px; top:865px; width:202px; height:36px; z-index:4"><%
	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
%></div>

<div id="Layer11" class="style3" style="position:absolute; left:<%=520+pageleft%>px; top:<%=890+pagetop%>px; z-index:10"><%="東警交"&"　　　　"&Sys_BillNo%></div>

<div id="Layer10" style="position:absolute; left:490px; top:840px; width:233px; height:32px; z-index:4"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<!--
<div id="Layer12" style="position:absolute; left:115px; top:920px; width:150px; height:11px; z-index:8"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>
-->
<!--20130509 交通隊施志林說「附採證照片」拿掉，以免廠牌太長時蓋到車號 -->
<div id="Layer12" style="position:absolute; left:115px; top:920px; width:150px; height:11px; z-index:8"><span class="style7">逕行舉發</span></div>

<div id="Layer15" style="position:absolute; left:255px; top:920px; width:100px; height:10px; z-index:8"><font size=2><%'if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" style="position:absolute; left:420px; top:920px; width:106px; height:13px; z-index:9"><%'=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:615px; top:920px; width:99px; height:12px; z-index:10"><%'=fastring%></div>
<div id="Layer18" style="position:absolute; left:120px; top:955px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:255px; top:955px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:495px; top:955px; width:300px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" style="position:absolute; left:120px; top:980px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)%></div>

<div id="Layer22" style="position:absolute; left:105px; top:1000px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:165px; top:1000px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:215px; top:1000px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:255px; top:1000px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:305px; top:1000px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:390px; top:1005px; width:610px; height:31px; z-index:20"><span class="style3"><%

	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
		end If 

	else

		response.write Sys_IllegalRule1
	end If 

	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then

			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"

		end If 
		
	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
	end If 
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
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if

%></div>

<div id="Layer33" style="position:absolute; left:450px; top:1125px; width:300px; height:80px; z-index:28"><span class="style7"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></span></font></div>

<div id="Layer35" style="position:absolute; left:395px; top:1160px; width:100px; height:49px; z-index:29"><%
	If Sys_UnitLevelID=1 Then
		response.write "<table border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td align=""center"" class=""style3"">台東縣警察局<br>交通警察隊</td></tr>"
		response.write "<tr><td align=""center"" class=""style3"">TEL(089)328344</td></tr>"
		response.write "</table>"
	elseIf Sys_UnitLevelID=2 Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" width=""100"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td align=""center"" width=""100"" class=""style10"">台東縣警察局<br>"&Sys_UnitName&"</td></tr>"
		response.write "<tr><td align=""center"" width=""100"" class=""style10"">TEL"&Sys_UnitTel&"</td></tr>"
		response.write "</table>"
	elseIf Sys_UnitLevelID=3 Then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" width=""100"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td align=""center"" width=""100"" class=""style10"">台東縣警察局<br>"&left(Sys_UnitName,instr(Sys_UnitName,"局"))&"<br>"&right(Sys_UnitName,len(Sys_UnitName)-instr(Sys_UnitName,"局"))&"</td></tr>"
		response.write "<tr><td align=""center"" width=""100"" class=""style10"">TEL"&Sys_UnitTel&"</td></tr>"
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
<div id="Layer41" style="position:absolute; left:670px; top:1235px; width:80px; height:12px; z-index:36"><%%></div>
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