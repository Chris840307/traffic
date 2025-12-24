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
.style2 {font-family:"標楷體"; font-size: 16px;}
.style3 {font-family:"標楷體"; font-size: 9px; color:#ff0000; }
.style4 {font-family:"標楷體"; font-size: 10px;}
.style5 {font-family:"標楷體"; font-size: 14px;}
.style6 {font-family:"標楷體"; font-size: 14px;}
.style7 {font-family:"標楷體"; font-size: 12px;}
-->
</style>
</head>

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

if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
	
	strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

If rsbil.eof Then 
	rsbil.close
	Response.Write "<span class=""style8"">案件還沒上傳監理站入案。</span>"
	Response.End
end If 

strSQL="select count(1) cnt from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))
set rsSend=conn.execute(strSQL)

If cdbl(rsSend("cnt")) = 0 Then

	strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,MailDate,SendMailStation,ArriveType,ReturnResonID,Note) values((select nvl(Max(SN),0)+1 as cnt from PassersEndArrived),"&trim(rsbil("BillSN"))&",sysdate,"&Session("User_ID")&","&Session("User_ID")&",sysdate,null,2,null,null)"

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

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.UnitLevelID,b.Tel,c.Content from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.jobid=c.id(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
if Not mem.eof then Sys_BillFillerJobName=mem("Content")
mem.close

If Sys_UnitLevelID=3 Then
	Sys_UnitName=replace(Sys_UnitName,"分局","分局<br>")
end if

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

DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,36

rsbil.close
%>

<div id="L78" style="position:relative;">

<div id="Layer2" style="position:absolute; left:<%=70+leftpx%>px; top:<%=30+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<div id="Layer4" style="position:absolute; left:<%=190+leftpx%>px; top:<%=20+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:675px; top:10px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:675px; top:<%=25+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:675px; top:<%=40+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; left:<%=50+leftpx%>px; top:<%=55+toppx%>px; width:233px; height:36px; z-index:5"><%
	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
%></div>
<div id="Layer10" style="position:absolute; left:<%=540+leftpx%>px; top:<%=50+toppx%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" class="style4" style="position:absolute; left:<%=535+leftpx%>px; top:<%=95+toppx%>px; width:230px; height:12px; z-index:7">　<%=BillPageUnit%>交字第　　　　　　　　　　　號</div>
<div id="Layer12" class="style5" style="position:absolute; left:<%=130+leftpx%>px; top:<%=115+toppx%>px; width:300px; height:11px; z-index:3"><%
	response.write "逕行舉發　"
%></div>

<div id="Layer18" class="style2" style="position:absolute; left:<%=135+leftpx%>px; top:<%=160+toppx%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" class="style2" style="position:absolute; left:<%=295+leftpx%>px; top:<%=160+toppx%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style2" style="position:absolute; left:<%=470+leftpx%>px; top:<%=160+toppx%>px; width:200px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" class="style2" style="position:absolute; left:<%=135+leftpx%>px; top:<%=180+toppx%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" class="style2" style="position:absolute; left:<%=145+leftpx%>px; top:<%=205+toppx%>px; width:200px; height:13px; z-index:15"><%

	Response.Write Sys_IllegalDate(0)&"年"&Sys_IllegalDate(1)&"月"&Sys_IllegalDate(2)&"日"
	Response.Write right("00"&Sys_IllegalDate_h,2)&"時"&right("00"&Sys_IllegalDate_m,2)&"分"
%></div>

<div id="Layer27" class="style5" style="position:absolute; left:<%=440+leftpx%>px; top:<%=210+toppx%>px; width:280px; height:31px; z-index:20"><%	
	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
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
		end If 
	else

		response.write Sys_IllegalRule1
	
	end If 
	
	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
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

		end If 
		
	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
	end If 
%></div>
<div id="Layer28" class="style2" style="position:absolute; left:<%=130+leftpx%>px; top:<%=240+toppx%>px; width:230px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>

<div id="Layer29" class="style2" style="position:absolute; left:<%=145+leftpx%>px; top:<%=270+toppx%>px; width:200px; height:11px; z-index:22"><%

	Response.Write Sys_DealLineDate(0)&"年"&Sys_DealLineDate(1)&"月"&Sys_DealLineDate(2)&"日前"
%></div>

<div id="Layer32" class="style6" style="position:absolute; left:<%=440+leftpx%>px; top:<%=280+toppx%>px; width:300px; height:49px; z-index:29"><%
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
%></div>

<div id="Layer34" class="style7" style="position:absolute; left:<%=430+leftpx%>px; top:<%=335+toppx%>px; width:200px; height:80px; z-index:28"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></div>

<div id="Layer35" class="style2" style="position:absolute; font-size:16px; left:<%=425+leftpx%>px; top:<%=430+toppx%>px; width:150px; height:49px; z-index:29"><%response.write "嘉義市政府警察局<br>"&Sys_UnitName%></div>

<div id="Layer37" style="position:absolute; font-size:10px; left:<%=625+leftpx%>px; top:<%=430+toppx%>px; width:200px; height:46px; z-index:31"><%
	if sys_City="嘉義市" then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillFillerJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
		response.write "</table>"
	end if
	if billprintuseimage=1 then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		end if
	end if
	response.write Sys_BillFillerMemberID
%></div>
<div id="Layer83" class="style3" style="position:absolute; left:210px; top:450px; width:200px; height:10px; z-index:32">(自103年3月31日起，前、後段日數均改為30日)</div>
<div id="Layer38" class="style2" style="position:absolute; left:<%=110+leftpx%>px; top:<%=465+toppx%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer39" class="style2" style="position:absolute; left:<%=160+leftpx%>px; top:<%=465+toppx%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer40" class="style2" style="position:absolute; left:<%=200+leftpx%>px; top:<%=465+toppx%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer41" class="style2" style="position:absolute; left:<%=250+leftpx%>px; top:<%=465+toppx%>px; width:120px; height:12px; z-index:36">填單</div>


<div id="Layer43" style="position:absolute; left:<%=70+leftpx%>px; top:<%=545+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<div id="Layer45" style="position:absolute; left:<%=190+leftpx%>px; top:<%=550+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<!--<div id="Layer46" style="position:absolute; left:<%=180+leftpx%>px; top:<%=565+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:<%=565+toppx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:<%=50+leftpx%>px; top:<%=580+toppx%>px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:<%=540+leftpx%>px; top:<%=580+toppx%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer52" class="style4" style="position:absolute; left:<%=540+leftpx%>px; top:<%=620+toppx%>px; width:230px; height:12px; z-index:7">　<%=BillPageUnit%>交字第　　　　　　　　　　　號</div>
<div id="Layer53" class="style5" style="position:absolute; left:<%=130+leftpx%>px; top:<%=640+toppx%>px; width:300px; height:11px; z-index:3"><%
	response.write "逕行舉發　"
%></div>

<div id="Layer59" class="style2" style="position:absolute; left:<%=135+leftpx%>px; top:<%=685+toppx%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" class="style2" style="position:absolute; left:<%=295+leftpx%>px; top:<%=685+toppx%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" class="style2" style="position:absolute; left:<%=470+leftpx%>px; top:<%=685+toppx%>px; width:200px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" class="style2" style="position:absolute; left:<%=135+leftpx%>px; top:<%=710+toppx%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" class="style2" style="position:absolute; left:<%=145+leftpx%>px; top:<%=735+toppx%>px; width:200px; height:13px; z-index:15"><%

	Response.Write Sys_IllegalDate(0)&"年"&Sys_IllegalDate(1)&"月"&Sys_IllegalDate(2)&"日"
	Response.Write right("00"&Sys_IllegalDate_h,2)&"時"&right("00"&Sys_IllegalDate_m,2)&"分"
%></div>

<div id="Layer68" class="style5" style="position:absolute; left:<%=440+leftpx%>px; top:<%=740+toppx%>px; width:280px; z-index:20"><%	
	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
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
		end If 
	else

		response.write Sys_IllegalRule1
	
	end If 
	
	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
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

		end If 
		
	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
	end If 

%></div>
<div id="Layer69" class="style2" style="position:absolute; left:<%=130+leftpx%>px; top:<%=770+toppx%>px; width:230px; height:15px; z-index:21"><%
'if sys_City="嘉義市" then response.write "嘉義市"
response.write Sys_ILLEGALADDRESS%></div>

<div id="Layer70" class="style2" style="position:absolute; left:<%=145+leftpx%>px; top:<%=800+toppx%>px; width:200px; height:11px; z-index:22"><%

	Response.Write Sys_DealLineDate(0)&"年"&Sys_DealLineDate(1)&"月"&Sys_DealLineDate(2)&"日前"
%></div>

<div id="Layer73" class="style6" style="position:absolute; left:<%=440+leftpx%>px; top:<%=815+toppx%>px; width:300px; height:49px; z-index:29"><%
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
%></div>

<div id="Layer75" class="style7" style="position:absolute; left:<%=430+leftpx%>px; top:<%=870+toppx%>px; width:200px; height:80px; z-index:28"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></div>

<div id="Layer74" style="position:absolute; left:<%=515+leftpx%>px; top:<%=875+toppx%>px; width:400px; height:30px; z-index:24"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>
<div id="Layer76" class="style2" style="position:absolute; font-size:16px; left:<%=425+leftpx%>px; top:<%=955+toppx%>px; width:150px; height:49px; z-index:29"><%if sys_City="嘉義市" then response.write "嘉義市政府警察局<br>"&Sys_UnitName%></div>
<div id="Layer77" style="position:absolute; font-size:10px; left:<%=625+leftpx%>px; top:<%=955+toppx%>px; width:200px; height:46px; z-index:31"><%
	if sys_City="嘉義市" then
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillFillerJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
		response.write "</table>"
	end if
	if billprintuseimage=1 then
		if trim(Sys_MemberFilename)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
		end if
	end if
	response.write Sys_BillFillerMemberID
%></div>
<div id="Layer83" class="style3" style="position:absolute; left:235px; top:975px; width:200px; height:10px; z-index:32">(自103年3月31日起，前、後段日數均改為30日)</div>
<div id="Layer78" class="style2" style="position:absolute; left:<%=110+leftpx%>px; top:<%=995+toppx%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer79" class="style2" style="position:absolute; left:<%=160+leftpx%>px; top:<%=995+toppx%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer80" class="style2" style="position:absolute; left:<%=200+leftpx%>px; top:<%=995+toppx%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer81" class="style2" style="position:absolute; left:<%=250+leftpx%>px; top:<%=995+toppx%>px; width:120px; height:12px; z-index:36">填單</div>
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