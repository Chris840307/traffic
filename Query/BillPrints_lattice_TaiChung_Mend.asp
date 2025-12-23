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
.pageprint {
  margin-left: 0mm;
  margin-right: 0mm;
  margin-top: 0mm;
  margin-bottom: 0mm;
}
</style>
</head>

<body>
<!--<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>-->
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
	sys_Date=split(gArrDT(trim(request("BillFillDate"))),"-")
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

Sys_DCIReturnStation=0
Sum_Level=0

Sys_Level1=trim(request("FORFEIT1"))
Sys_Level2=trim(request("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
Sys_DCIReturnStation=trim(request("DCIReturnStation"))
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
	Sys_DealLineDate=split(gArrDT(trim(request("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if

Sys_UnitName=trim(request("FillUnitName"))
Sys_UnitTel=trim(request("FillUnitTEL"))
Sys_IllegalRule1=trim(request("Rule1txt"))
Sys_IllegalRule2=trim(request("Rule2txt"))


strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

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
BillSN=0

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode BillSN,Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	Sys_MailNumber=0
	DelphiASPObj.GenBillPrintBarCode BillSN,Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate
end if
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" class="pageprint" style="position:relative;"><%
if showBarCode then
%>
<div id="Layer1" style="position:absolute; left:0px; top:15px; width:10px; height:20px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:0px; top:40px; width:10px; height:20px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:130px; top:15px; width:202px; height:36px; z-index:5">v</div>
<%else%>
	<div id="Layer4" style="position:absolute; left:130px; top:25px; width:202px; height:36px; z-index:5">v</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:625px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:625px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:625px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:-30px; top:65px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　　　　　"&Sys_UnitName
	end if
%></div>
<!--<div id="Layer42" style="position:absolute; left:210px; top:70px; width:202px; height:36px; z-index:5"><%="<font size=1>"&SysUnit&"<br>("&SysUnitTel&")</font>"%></div>-->
<div id="Layer10" style="position:absolute; left:460px; top:65px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:485px; top:110px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer12" style="position:absolute; left:70px; top:130px; width:150px; height:11px; z-index:3">逕行舉發　<%=Sys_A_Name%><br>附採證照片　<%=Sys_CarColor%></font></div>
<div id="Layer13" style="position:absolute; left:215px; top:130px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:300px; top:130px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div>
<div id="Layer15" style="position:absolute; left:215px; top:155px; width:100px; height:10px; z-index:8"><font size=2></font></div>
<div id="Layer16" style="position:absolute; left:370px; top:155px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:560px; top:155px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:70px; top:180px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:220px; top:180px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:450px; top:180px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" style="position:absolute; left:70px; top:210px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; left:70px; top:235px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:125px; top:235px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:180px; top:235px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:230px; top:235px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:290px; top:235px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:375px; top:235px; width:400px; height:31px; z-index:20"><%
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
			response.write " (經雷達、雷射測速儀器採證)"
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"
		
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
	response.write "</font>"
%></div>
<div id="Layer28" style="position:absolute; left:70px; top:260px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:120px; top:280px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:180px; top:280px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:260px; top:280px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" style="position:absolute; left:385px; top:290px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　"&Sys_Level1
	if not ifnull(Sys_Rule2) then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer33" style="position:absolute; left:370px; top:320px; width:400px; height:30px; z-index:28"><%if showBarCode then response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"%></div>

<div id="Layer34" style="position:absolute; left:610px; top:325px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>


<!--<div id="Layer35" style="position:absolute; left:455px; top:420px; width:100px; height:49px; z-index:29"><%
	'if billprintuseimage=1 then
		'response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""70"" height=""70"">"
	'else
		'response.write Sys_UnitName
	'end if%></div>
<div id="Layer36" style="position:absolute; left:580px; top:420px; width:100px; height:43px; z-index:30">主管</div>
<div id="Layer37" style="position:absolute; left:660px; top:410px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""90"" height=""30"">"
	else
		response.write Sys_ChName
	end if
%></div>-->
<div id="Layer38" style="position:absolute; left:170px; top:470px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:330px; top:470px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:480px; top:470px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:650px; top:470px; width:80px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>
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
	window.print();
	//printWindow(true,5.08,5.08,5.08,5.08);
</script>