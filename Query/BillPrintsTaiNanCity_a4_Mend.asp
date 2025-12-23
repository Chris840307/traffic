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
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
Sys_BillNo=trim(request("BillNo"))
Sys_CarNo=trim(request("CarNo"))
Sys_BillTypeID=trim(request("BillTypeID"))
Sys_DriverID=trim(request("DriverID"))
Sys_IllegalAddress=trim(request("IllegalAddress"))
Sys_IllegalSpeed=trim(request("IllegalSpeed"))
Sys_RuleSpeed=trim(request("RuleSpeed"))
Sys_Note=trim(request("Note"))

if Not ifnull(request("BillFillDate")) then
	sys_Date=split(gArrDT(request("BillFillDate")),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if

If Not ifnull(trim(request("Owner"))) Then
	Sys_Owner=trim(request("Owner"))
	Sys_OwnerAddress=trim(request("OwnerAddress"))
	Sys_OwnerZip=trim(request("OwnerZip"))
else
	Sys_Owner=trim(request("Driver"))
	Sys_OwnerAddress=trim(request("DriverAddress"))
	Sys_OwnerZip=trim(request("DriverZip"))
end if
Sys_OwnerZipName=request("OwnerZipName")


If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&Sys_BillNo&","&Sys_Owner&"\n"
end if

Sys_StationID=0
Sum_Level=0

Sys_Level1=trim(request("FORFEIT1"))
Sys_Level2=trim(request("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
Sys_StationID=trim(request("DCIReturnStation"))
Sys_Rule1=trim(request("Rule1"))
Sys_Rule2=trim(request("Rule2"))
Sys_DCIRETURNCARTYPE=trim(request("DetailCarType"))

Sys_Sex=""

if trim(Sys_BillTypeID)="1" then
	If not ifnull(Trim(request("DriverID"))) Then
		If Mid(Trim(request("DriverID")),2,1)="1" Then
			Sys_Sex="男"
		elseif Mid(Trim(request("DriverID")),2,1)="2" Then
			Sys_Sex="女"
		End if
	End if
end if

Sys_RecordMemberID=trim(request("operat"))

if Not ifnull(request("IllegalDate")) then
	Sys_IllegalDate=split(gArrDT(trim(request("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end if
if Not ifnull(request("IllegalDate")) then
	Sys_IllegalDate_h=hour(trim(request("IllegalDate")))
else
	Sys_IllegalDate_h=""
end if
if Not ifnull(request("IllegalDate")) then
	Sys_IllegalDate_m=minute(trim(request("IllegalDate")))
else
	Sys_IllegalDate_m=""
end if
if Not ifnull(request("DealLineDate")) then
	Sys_DealLineDate=split(gArrMT(trim(request("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrMT(trim("")),"-")
end if

If cdbl(Sys_IllegalDate(0)&Sys_IllegalDate(1)&Sys_IllegalDate(2)) < 991225 Then
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

Sys_UnitName=trim(request("FillUnitName"))
Sys_UnitTel=trim(request("FillUnitTEL"))
Sys_IllegalRule1=trim(request("Rule1txt"))
Sys_IllegalRule2=trim(request("Rule2txt"))

Sys_STATIONNAME=trim(request("STATIONNAME"))
Sys_MailDate=trim(request("MailDate"))

'LEO偷偷改過
fastring=""
If Not ifnull(request("Hold"))  Then
	tempFastring=split(trim(request("Hold")),",")
	For i = 0 to Ubound(tempFastring)
		if trim(tempFastring(i)) <> "" then
		    if trim(fastring)<>"" then 
		        fastring=fastring&","
		    else
	            fastring=fastring&tempFastring(i)
	        end if
        end if  
	Next
End if

if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Sys_A_Name=trim(request("CarMark"))
Sys_CarColor=trim(request("CarColor"))
'Sys_BillJobName=trim(request("ManagerLevel"))
'Sys_ChName=trim(request("Boss"))
Sys_ChName=replace(Request("BillFillerName"),"　","&nbsp;&nbsp;<br>")

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if
strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;"><%
if showBarCode then
%>
<div id="Layer1" style="position:absolute; left:80px; top:0px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:80px; top:20px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:200px; top:0px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:200px; top:15px; width:202px; height:36px; z-index:5">v</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:675px; top:10px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:675px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:675px; top:40px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>>-->
<div id="Layer9" style="position:absolute; left:55px; top:50px; width:233px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:520px; top:45px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" style="position:absolute; left:530px; top:85px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第　　　　　　　　　號</font></div>
<div id="Layer12" style="position:absolute; left:135px; top:110px; width:150px; height:11px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer13" style="position:absolute; left:290px; top:110px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:400px; top:106px; width:370px; height:10px; z-index:4"><font size=2><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納(7-11,全家,萊爾富,OK)"%></font></div>
<div id="Layer15" style="position:absolute; left:285px; top:130px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; left:465px; top:130px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:660px; top:130px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:135px; top:155px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:300px; top:155px; width:140px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:470px; top:155px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" style="position:absolute; left:140px; top:175px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; left:135px; top:200px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer23" style="position:absolute; left:175px; top:200px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer24" style="position:absolute; left:215px; top:200px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer25" style="position:absolute; left:255px; top:200px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer26" style="position:absolute; left:295px; top:200px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>
<div id="Layer27" style="position:absolute; left:435px; top:200px; width:620px; height:31px; z-index:20"><%
	response.write "<font size=2>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "<br>100公里以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "<br>80公里以上未滿100公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "<br>60公里以上未滿80公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "<br>40公里以上未滿60公里"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "<br>20公里以上未滿40公里"
			else
				response.write "<br>未滿20公里"
			end if
			
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if	
	end if

	if not ifnull(Sys_Rule4) then response.write "("&Sys_Rule4&")"

	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end if
	end if
		
	response.write " (經科學儀器採證)"
				
	'	if trim(Sys_Note)<>"" then response.write "<br>("&Sys_Note&")"
	response.write "</font>"
%></div>
<div id="Layer28" style="position:absolute; left:135px; top:235px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:135px; top:265px; width:267px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer30" style="position:absolute; left:175px; top:265px; width:267px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer31" style="position:absolute; left:215px; top:265px; width:267px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer32" style="position:absolute; left:435px; top:260px; width:400px; height:49px; z-index:29"><%
	Response.Write "<font size='2'>"
	Response.Write "　　　　　道 路 交 通 管 理 處 罰 條 例<br>"
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:435px; top:320px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>

<div id="Layer33" style="position:absolute; left:515px; top:315px; width:400px; height:30px; z-index:28"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<div id="Layer35" style="position:absolute; left:430px; top:380px; width:130px; height:49px; z-index:29"><%'="<font size=2>臺南市政府警察局"&SysUnit&"<br>"&SysUnitTel&"</font>"%></div>
<div id="Layer38" style="position:absolute; left:175px; top:460px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer39" style="position:absolute; left:330px; top:460px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer40" style="position:absolute; left:480px; top:460px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer41" style="position:absolute; left:580px; top:460px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>


<%if showBarCode then%>
<div id="Layer42" style="position:absolute; left:80px; top:525px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer43" style="position:absolute; left:80px; top:545px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:200px; top:530px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:200px; top:540px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:670px; top:535px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:670px; top:550px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:670px; top:565px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->

<div id="Layer50" style="position:absolute; left:55px; top:575px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:520px; top:575px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer52" style="position:absolute; left:530px; top:615px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第　　　　　　　　　號</font></div>
<div id="Layer53" style="position:absolute; left:135px; top:640px; width:150px; height:11px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer54" style="position:absolute; left:300px; top:640px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:400px; top:640px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; left:295px; top:665px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer57" style="position:absolute; left:465px; top:665px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:660px; top:665px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:135px; top:680px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:300px; top:680px; width:140px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:480px; top:680px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" style="position:absolute; left:135px; top:705px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" style="position:absolute; left:135px; top:730px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer64" style="position:absolute; left:175px; top:730px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer65" style="position:absolute; left:215px; top:730px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer66" style="position:absolute; left:255px; top:730px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%>時</div>
<div id="Layer67" style="position:absolute; left:295px; top:730px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%>分</div>
<div id="Layer68" style="position:absolute; left:435px; top:730px; width:620px; height:31px; z-index:20"><%
	response.write "<font size=2>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
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
			'response.write "(經科學儀器採證)"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if	
	end if
	if not ifnull(Sys_Rule4) then response.write "("&Sys_Rule4&")"
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
	response.write " (經科學儀器採證)"
	response.write "</font>"
%></div>
<div id="Layer69" style="position:absolute; left:135px; top:765px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:135px; top:795px; width:40px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer71" style="position:absolute; left:175px; top:795px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer72" style="position:absolute; left:215px; top:795px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日前</div>
<div id="Layer73" style="position:absolute; left:435px; top:795px; width:400px; height:49px; z-index:29"><%
	Response.Write "<font size='2'>"
	Response.Write "　　　　　道 路 交 通 管 理 處 罰 條 例<br>"
	response.write "第"&left(trim(Sys_Rule1),2)&"條"
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
	response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
		if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
		response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
	end if
	response.write "</font>"
%></div>
<div id="Layer75" style="position:absolute; left:435px; top:840px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>
<div id="Layer74" style="position:absolute; left:515px; top:845px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>
<div id="Layer76" style="position:absolute; left:430px; top:940px; width:120px; height:49px; z-index:29"><%
	'response.write "<font size=2>臺南市政府警察局"&SysUnit&"<br>"&SysUnitTel&"</font>"
%></div>
<div id="Layer78" style="position:absolute; left:175px; top:990px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer79" style="position:absolute; left:330px; top:990px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer80" style="position:absolute; left:480px; top:990px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer81" style="position:absolute; left:580px; top:990px; width:120px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>
</div>
<%

If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,5.08,11,5.08,5.08);
</script>