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
.style1 {font-family:"標楷體"; font-size: 10px; color:#ff0000; }
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style6 {font-size: 20px; line-height:2;}
.style7 {font-size: 16px}
.style8 {font-size: 36px}
.style9 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style10 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style11 {font-size: 14px}
.style12 {font-family:"標楷體"; font-size: 8px; color:#ff0000; }
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
codebase="..\smsxie8.cab#Version=6,5,439,50">
</object>
<%
'on Error Resume Next
PBillSN=split(trim(request("BillSN")),",")
Server.ScriptTimeout=6000
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 

for i=0 to ubound(PBillSN)

if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
	
strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

if Not rsbil.eof then Sys_BatchNumber=trim(rsbil("BatchNumber"))

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

	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"臺","台"),Sys_OwnerZipName,"")

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

strSQL="select UnitName,Tel,Address from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysAddress=Unit("Address")
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

Sys_MailNumber=0

strSQL="select min(SendMailStation) SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))

set rsSend=conn.execute(strSQL)

if Not rsSend.eof then
	Sys_MailNumber=rsSend("SendMailStation")

end If 
rsSend.close

DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,200,016,36

DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,60,160

Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo

rsbil.close

'if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
'err.Clear
%>
<div id="L78" style="position:relative;">

<div id="Layer1" style="position:absolute; left:60px; top:20px; z-index:5">
<%
Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_1.jpg"" hspace=""0"" vspace=""0"">"
Response.Write "<br><span class=""style7"">"&Sys_FirstBarCode&"</span>"
%>
</div>
<div id="Layer2" class="style3" style="position:absolute; left:250px; top:20px; width:350px; height:36px; z-index:5"><%
If not ifnull(request("Sys_UnitLabelKind")) Then
	response.write "<b>"&SysAddress&"<br>"&SysUnit&"</b>"
End if%>
</div>

<div id="Layer2" style="position:absolute; left:250px; top:70px; width:350px; height:36px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write Sys_OwnerZip&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,1)&"<br>"
		Response.Write funcCheckFont(Sys_Owner,20,1)&"　台啟"
	Response.Write "</span>"%>
</div>

<div id="Layer5" style="position:absolute; left:60px; top:180px; z-index:5"><%
	Response.Write "<span class=""style3"">"
	Response.Write "應到案處所："&Sys_STATIONNAME&"<br>"
   	Response.Write "應到案處所電話："&Sys_StationTel
	Response.Write "</span>"

%>
</div>
<div id="Layer5" style="position:absolute; left:550px; top:160px; z-index:5"><%
	Response.Write "<span class=""style8"">"
	Response.Write Sys_StationID
	Response.Write "</span>"
%>
</div>

<div id="Layer6" style="position:absolute; left:60px; top:320px; width:400px; height:36px; z-index:5"><span class="style7">查詢電話：<%=SysUnitTel%>（<%=SysUnit%>）</span></div>

<div id="Layer9" style="position:absolute; left:20px; top:425px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:415px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer14" style="position:absolute; left:360px; top:485px; width:324px; height:10px; z-index:4"><span class="style1"><%
	'Response.Write Sys_DriverZipName&Sys_DriverHomeAddress
	Response.Write "＊受通知人收受通知單時應到案日期不足30日或已逾應到案日期者，得於送達生效日後30日內到案。"
%></Span></div>
<div id="Layer15" style="position:absolute; left:265px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%'=Sys_DriverBirth(0)%></span></div>
<div id="Layer142" style="position:absolute; left:295px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%'=Sys_DriverBirth(1)%></span></div>
<div id="Layer143" style="position:absolute; left:325px; top:520px; width:100px; height:10px; z-index:8"><span class="style7"><%'=Sys_DriverBirth(2)%></span></div>
<div id="Layer16" style="position:absolute; left:430px; top:520px; width:106px; height:13px; z-index:9"><span class="style7"><%'=Sys_DriverID%></span></div>
<div id="Layer17" style="position:absolute; left:630px; top:520px; width:99px; height:12px; z-index:10"><span class="style7"><%'=fastring%></span></div>
<%'end if%>
<div id="Layer18" style="position:absolute; left:125px; top:545px; width:100px; height:14px; z-index:11"><span class="style7"><%=Sys_CarNo%></span></div>
<div id="Layer19" style="position:absolute; left:310px; top:545px; width:117px; height:20px; z-index:12"><span class="style7"><%=Sys_DCIRETURNCARTYPE%></span></div>
<div id="Layer20" style="position:absolute; left:570px; top:545px; width:201px; height:17px; z-index:13"><span class="style7"><%=funcCheckFont(Sys_Owner,22,1)%></span></div>
<div id="Layer21" style="position:absolute; left:165px; top:570px; width:507px; height:13px; z-index:14"><span class="style7"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></span></div>

<div id="Layer22" style="position:absolute; left:105px; top:600px; width:40px; height:13px; z-index:15"><span class="style7"><%=Sys_IllegalDate(0)%></span></div>
<div id="Layer23" style="position:absolute; left:160px; top:600px; width:40px; height:17px; z-index:16"><span class="style7"><%=Sys_IllegalDate(1)%></span></div>
<div id="Layer24" style="position:absolute; left:220px; top:600px; width:40px; height:16px; z-index:17"><span class="style7"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:270px; top:600px; width:40px; height:16px; z-index:18"><span class="style7"><%=right("00"&Sys_IllegalDate_h,2)%></span></div>
<div id="Layer26" style="position:absolute; left:330px; top:600px; width:40px; height:13px; z-index:19"><span class="style7"><%=right("00"&Sys_IllegalDate_m,2)%></span></div>
<div id="Layer27" style="position:absolute; left:430px; top:600px; width:270px; height:31px; z-index:20"><span class="style4"><%
	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write Sys_IllegalRule1
			response.write "（該路段限速"&Sys_RuleSpeed&"公里、經雷達(射)測速為"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里）"
		end If 


	else

		response.write Sys_IllegalRule1
	
	end If 

	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "<br>"&Sys_IllegalRule1
			response.write "（該路段限速"&Sys_RuleSpeed&"公里、經雷達(射)測速為"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里）。"
		end If 


	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule1
	
	end if
%></span></div>
<div id="Layer28" style="position:absolute; left:110px; top:625px; width:217px; height:15px; z-index:21"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:140px; top:645px; width:34px; height:11px; z-index:22"><span class="style7"><%=Sys_DealLineDate(0)%></span></div>
<div id="Layer30" style="position:absolute; left:220px; top:645px; width:35px; height:13px; z-index:23"><span class="style7"><%=Sys_DealLineDate(1)%></span></div>
<div id="Layer31" style="position:absolute; left:300px; top:645px; width:32px; height:15px; z-index:24"><span class="style7"><%=Sys_DealLineDate(2)%></span></div>
<div id="Layer32" style="position:absolute; left:430px; top:685px; width:400px; height:49px; z-index:29"><span class="style4"><%response.write "第"&left(trim(Sys_Rule1),2)&"條"
			'if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款規定"
				response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"

			if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
				response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
				'if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
				response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款規定"
				response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
			end if
			%></span></div>

<div id="Layer34" style="position:absolute; left:480px; top:730px; width:90px; height:30px; z-index:28"><span class="style3"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></span></div>

<div id="Layer35" style="position:absolute; left:423px; top:785px; width:100px; height:49px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"

		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style10"">&nbsp;"&Sys_UnitName&"&nbsp;</span><br><span class=""style10"">&nbsp;"&Sys_UnitTEL&"&nbsp;</span></td></tr>"

	response.write "</table>"
%></div>
<div id="Layer36" style="position:absolute; left:500px; top:745px; width:100px; height:43px; z-index:30"><%
'	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'	if Session("Unit_ID") <>"0207" then 
'		response.write "<tr><td nowrap style=""border-color:#ff0000;border-style:solid;"" align=""center""><span class=""style9"">主管職名章</span><br><span class=""style10"">"&Sys_JobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'	end if
'	response.write "</table>"
%></div>
<div id="Layer37" style="position:absolute; left:615px; top:760px; width:200px; height:46px; z-index:31"><%
'	if trim(Sys_MemberFilename)<>"" then
'		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""110"" height=""40"">"
'	else
'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=63 height=25 align=""center""><span class=""style9"">"&Sys_BillJobName&"&nbsp;&nbsp;&nbsp;</span><br><span class=""style10"">"&Sys_ChName&"</span></td></tr>"
'		response.write "</table><font size=2>　　"&Sys_BillFillerMemberID&"</font>"
'	end if
%></div>

<div id="Layer47" class="style12" style="position:absolute; left:240px; top:825px; width:200px; height:10px; z-index:32">(自103年3月31日起，前、後段日數均改為30日)</div>
<div id="Layer38" style="position:absolute; left:135px; top:835px; width:60px; height:10px; z-index:32"><span class="style7"><%=sys_Date(0)%></span></div>
<div id="Layer39" style="position:absolute; left:195px; top:835px; width:60px; height:13px; z-index:33"><span class="style7"><%=sys_Date(1)%></span></div>
<div id="Layer40" style="position:absolute; left:255px; top:835px; width:60px; height:11px; z-index:34"><span class="style7"><%=sys_Date(2)%></span></div>
<div id="Layer41" style="position:absolute; left:300px; top:835px; width:80px; height:12px; z-index:36"><span class="style7"><%
	if Session("Unit_ID") <>"0207" then Response.Write Sys_BillFillerMemberID
%></span></div>


<div id="Layer45" style="position:absolute; left:180px; top:1005px; width:100px; height:12px; z-index:36"><span class="style3"><%
	Response.Write Sys_BillNo
%></span></div>

<div id="Layer43" style="position:absolute; left:210px; top:1022px; width:350px; height:12px; z-index:36"><span class="style3"><%'=Sys_MAILCHKNUMBER%></span></div>

<div id="Layer44" style="position:absolute; left:370px; top:1030px; width:350px; height:12px; z-index:10"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>></div>

<div id="Layer45" style="position:absolute; left:185px; top:1050px; width:100px; height:12px; z-index:36"><span class="style3"><%
	Response.Write funcCheckFont(Sys_Owner,22,1)
%></span></div>

<div id="Layer42" style="position:absolute; left:185px; top:1070px; width:230px; height:12px; z-index:36; background-color:#FFFFFF"><span class="style3"><%
	Response.Write Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)
%></span></div>

<div id="Layer46" style="position:absolute; left:450px; top:1140px; width:230px; height:12px; z-index:36; background-color:#FFFFFF"><span class="style3"><%
	Response.Write "<B>□ 本人　□ 代收</B>"
%></span></div>

<div id="Layer42" style="position:absolute; left:190px; top:1285px; width:500px; z-index:36; background-color:#FFFFFF"><span class="style7"><%
If cdbl(sys_UnitLevelAddr) > 1 Then
	response.write "<font color=""red"">請繳回："
	SysAddress=replace(SysAddress,"一","１")
	SysAddress=replace(SysAddress,"二","２")
	SysAddress=replace(SysAddress,"三","３")
	SysAddress=replace(SysAddress,"四","４")
	Response.Write SysAddress
	Response.Write "　　"
	
	SysUnit=replace(SysUnit,"一","１")
	SysUnit=replace(SysUnit,"二","２")
	SysUnit=replace(SysUnit,"三","３")
	SysUnit=replace(SysUnit,"四","４")
	Response.Write SysUnit
	Response.Write "</font>"
End if
%></span></div>

</div>

<%
'	If trim(Sys_BillUnitTypeID) = "0207" Then
'		if (i mod 30)=0 then response.flush
'	else
		response.flush
'	End if 
	
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