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
'on Error Resume Next
strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
Sys_RuleVer=trim(rsCity("value"))
rsCity.close

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

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

If Sys_BillTypeID=2 Then
	If len(trim(Sys_Owner))<3 Then errBillNo=errBillNo&Sys_BillNo&","&Sys_Owner&"\n"
end if

Sys_StationID=0
Sum_Level=0:Sys_Level1=0:Sys_Level2=0

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

BillSN=0:Sys_MailNumber=0
if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode BillSN,Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& BillSN&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode BillSN,Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,451,36

	'response.write "DelphiASPObj.GenBillPrintBarCode"& BillSN&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
end if

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

<div id="Layer50" style="position:absolute; left:35px; top:595px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:520px; top:590px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer52" style="position:absolute; left:515px; top:620px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer53" style="position:absolute; left:125px; top:655px; width:250px; height:30px; z-index:3"><font size=2>逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer54" style="position:absolute; left:295px; top:650px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:385px; top:650px; width:324px; height:10px; z-index:4"></div>
<div id="Layer56" style="position:absolute; left:290px; top:660px; width:100px; height:10px; z-index:8"></div>
<div id="Layer57" style="position:absolute; left:465px; top:675px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:650px; top:675px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:125px; top:710px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:285px; top:710px; width:200px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:520px; top:710px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer62" style="position:absolute; left:125px; top:740px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,22,1)%></div>

<div id="Layer63" style="position:absolute; left:140px; top:765px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer64" style="position:absolute; left:190px; top:765px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer65" style="position:absolute; left:240px; top:765px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer66" style="position:absolute; left:290px; top:765px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer67" style="position:absolute; left:340px; top:765px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer68" class="style4" style="position:absolute; left:440px; top:760px; width:320px; height:31px; z-index:20"><%
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里，經檢定合格儀器測照，時速"&Sys_IllegalSpeed&"公里，超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>100 then
				response.write "(100以上)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>80 then
				response.write "(80以上未滿100)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>60 then
				response.write "(60以上未滿80)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>40 then
				response.write "(40以上未滿60)。"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>20 then
				response.write "(20以上未滿40)。"
			else
				response.write "(未滿20公里)。"
			end if
		else
			Response.Write Sys_IllegalRule1
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
<div id="Layer69" style="position:absolute; left:125px; top:785px; width:267px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:170px; top:810px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer71" style="position:absolute; left:240px; top:810px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer72" style="position:absolute; left:300px; top:810px; width:50px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer73" style="position:absolute; left:455px; top:820px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer75" style="position:absolute; left:435px; top:855px; width:100px; height:30px; z-index:28"><font size=2><%
'2013/05/22
If Sys_BillNo="HD6167346" Then
	response.write "臺中市監理站"
else
	response.write Sys_STATIONNAME
End If 
%></font></div>

<div id="Layer74" style="position:absolute; left:515px; top:850px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>

<div id="Layer76" style="position:absolute; left:425px; top:930px; z-index:29"><%
'	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style2"">"&Sys_UnitName&"</span><br><span class=""style2"">"&Sys_UnitTEL&"</span></td></tr>"
'	response.write "</table>"
%></div>
<div id="Layer77" style="position:absolute; left:600px; top:940px; width:200px; height:46px; z-index:31"><%
'	if trim(Sys_MemberFilename)<>"" then
'		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
'	else
'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=63 height=25 align=""center""><span class=""style2"">"&Sys_ChName&"</span></td></tr>"
'		response.write "</table>"
'		Response.Write "<font size=2>　　"&Sys_BillFillerMemberID&"</font>"
'	end if
%></div>
<div id="Layer78" style="position:absolute; left:205px; top:995px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer79" style="position:absolute; left:360px; top:995px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer80" style="position:absolute; left:520px; top:995px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer81" style="position:absolute; left:580px; top:995px; width:120px; height:12px; z-index:36"></div>
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
	printWindow(true,8.08,5.08,5.08,5.08);
</script>