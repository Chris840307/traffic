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
.style1 {font-family:"標楷體";font-size: 14px; color:#ff0000;}
.style2 {font-family:"標楷體";font-size: 20px; line-height:1;}
.style3 {font-family:"標楷體";font-size: 14px}
.style33{font-family:"標楷體";font-size: 14px}
.style4 {font-family:"標楷體";font-size: 12px}
.style7 {font-family:"標楷體";font-size: 11px}
.style8 {font-family:"標楷體";font-size: 36px}
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
'on Error Resume Next
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
pagepx=60
Sys_DriverHomeAddress=Trim(Request("DriverHomeAddress"))
Sys_DriverHomeZip=Trim(Request("DriverHomeZip"))
Sys_DriverZipName=Trim(Request("DriverHomeZipName"))
Sys_Owner=Trim(Request("Owner"))
Sys_OwnerAddress=Trim(Request("OwnerAddress"))
Sys_OwnerZip=Trim(Request("OwnerZip"))
Sys_OwnerZipName=Trim(Request("OwnerZipName"))
Sys_IllegalSpeed=Trim(Request("IllegalSpeed"))
Sys_RuleSpeed=Trim(Request("RuleSpeed"))
Sys_Note=Trim(Request("Note"))
Sys_OwnerZipName=Trim(Request("OwnerZipName"))
sys_Date=split(gArrDT(trim(Request("BillFillDate"))),"-")
Sys_Rule1=trim(Request("Rule1"))
Sys_Rule2=trim(Request("Rule2"))
Sys_Level1=trim(Request("FORFEIT1"))
Sys_Level2=trim(Request("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
Sys_BillNo=Trim(Request("BillNo"))
Sys_CarNo=Trim(Request("CarNo"))
Sys_IllegalRule1=Trim(Request("IllegalRule1"))
Sys_IllegalRule2=Trim(Request("IllegalRule2"))
Sys_DCIRETURNCARTYPE=Trim(Request("DCIRETURNCARTYPE"))
Sys_RECORDMEMBERName=Trim(Request("IllegalDate"))
Sys_IllegalDate=split(gArrDT(trim(Request("IllegalDate"))),"-")
Sys_IllegalDate_h=hour(trim(Request("IllegalDate")))
Sys_IllegalDate_m=minute(trim(Request("IllegalDate")))
Sys_BillFillerMemberID=Trim(Request("BillFillerMemberLoginID"))
Sys_UnitName=Trim(Request("UnitName"))
Sys_UnitTEL=Trim(Request("UnitFillerTel"))
SysUnit=Trim(Request("SysUnit"))
SysUnitAddress=Trim(Request("SysUnitAddress"))
Sys_STATIONNAME=Trim(Request("STATIONNAME"))
Sys_StationTel=Trim(Request("StationTel"))
Sys_StationID=Trim(Request("StationID"))
Sys_MailNumber=Trim(Request("MailNumber"))
Sys_MailDate=now
fastring=Trim(Request("Fastring"))
A_Name=Trim(Request("A_Name"))
CarColor=Trim(Request("CarColor"))
Sys_DealLineDate=split(gArrDT(trim(Request("DealLineDate"))),"-")
Sys_MAILCHKNUMBER=Trim(Request("MAILCHKNUMBER"))
Sys_BillTypeID=Trim(Request("BillTypeID"))
Sys_ILLEGALADDRESS=Trim(Request("ILLEGALADDRESS"))
Sys_ChName=Trim(Request("ChName"))

PBillSN=0
If ifnull(Sys_MailNumber) Then Sys_MailNumber=0

DelphiASPObj.GenBillPrintBarCode_PT PBillSN,Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,"900","018","17"


if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
pagesum=530

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
		'If Sys_UnitLevelID < 2 Then
		'	Response.Write "<p class=""style4""><font size=""2"">大宗郵資已付掛號函件<br>  第<"&right("00000000" & trim(Sys_MailNumber),6)&"號  </font></p>"
		'end if%>    </td>
    <td >&nbsp;</td>
    
  </tr>

  <tr>

	
    <td >&nbsp;</td>
    <td width="530" align="center"><%
'		If Sys_UnitLevelID < 2 Then
'			Response.Write "<div align=""center""><img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""></img><br><font size=""2"">"&Sys_MAILCHKNUMBER&"</font></div>"
'		end if
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
				response.write Sys_DriverHomeZip&"<br>"
				response.write Sys_DriverZipName&Sys_DriverHomeAddress&"<br>"
				response.write funcCheckFont(Sys_Owner,20,1)
				%>	　敬啟<br><br><br>
			
		</font></b></td> 
		<!--  監理站代碼
				
		<td><p class="style6"><%=Sys_StationID%></p></td> 
		
		-->

	</tr>
</table>
</div>
<!---------------------------------------- start 列印紅單紅色區域內容 --------------------------------------------->
<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:30px; top:<%=335+pagepx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer2" style="position:absolute; left:30px; top:<%=380+pagepx%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer3" style="position:absolute; left:145px; top:<%=350+pagepx%>px; width:202px; height:36px; z-index:5">v</div>
<%else%>
<div id="Layer4" style="position:absolute; left:145px; top:<%=365+pagepx%>px; width:202px; height:36px; z-index:5">v</div>
<%end if%>
<div id="Layer9" style="position:absolute; left:20px; top:<%=410+pagepx%>px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write "　　"&SysUnit
	end if
%></div>
<div id="Layer10" style="position:absolute; left:500px; top:<%=400+pagepx%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer42" style="position:absolute; left:535px; top:<%=460+pagepx%>px; width:233px; height:12px; z-index:36"><font size=2></font></div>

<div id="Layer12" style="position:absolute; left:117px; top:<%=470+pagepx%>px; width:300px; height:11px; z-index:20"><span class="style7">逕行舉發　<%=Sys_A_Name%><br><%if int(Sys_Rule1)<>4340003 and int(Sys_Rule1)<>5620001 then response.write "附採證照片"%>　<%=Sys_CarColor%></span></div>
<div id="Layer14" style="position:absolute; left:390px; top:<%=465+pagepx%>px; width:324px; height:10px; z-index:4"><%if showBarCode then response.write "*本單可至郵局或委託代收之超商繳納"%></div><%'=Sys_DriverHomeZip&" "&Sys_DriverZipName&Sys_DriverHomeAddress%>
<div id="Layer17" style="position:absolute; left:630px; top:<%=485+pagepx%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:105px; top:<%=525+pagepx%>px; width:100px; height:14px; z-index:11"><b><%=Sys_CarNo%></b></div>
<div id="Layer19" style="position:absolute; left:310px; top:<%=525+pagepx%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:550px; top:<%=525+pagepx%>px; width:201px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,20,1)%></div>
<div id="Layer21" style="position:absolute; left:175px; top:<%=555+pagepx%>px; width:520px; height:13px; z-index:14"><%=Sys_DriverHomeZip&" "&funcCheckFont(Sys_DriverZipName&Sys_DriverHomeAddress,20,1)%></div>

<div id="Layer22" style="position:absolute; left:130px; top:<%=580+pagepx%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:170px; top:<%=580+pagepx%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:220px; top:<%=580+pagepx%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:270px; top:<%=580+pagepx%>px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:330px; top:<%=580+pagepx%>px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:410px; top:<%=580+pagepx%>px; width:610px; height:31px; z-index:20"><span class="style33"><%
	response.write "<font size=3>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"
			
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
		if len(Sys_IllegalRule1)<25 then
			response.write Sys_IllegalRule1
		else
			response.write left(Sys_IllegalRule1,25)&"<br>"&mid(Sys_IllegalRule1,26,len(Sys_IllegalRule1))
		end if
	end If 
	Response.Write "(新臺幣"&Sys_Level1&"元)"
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if int(Sys_Rule2)=4340003 then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		if len(Sys_IllegalRule2)<26 then
			response.write "<br>"&Sys_IllegalRule2
		else
			response.write "<br>"&left(Sys_IllegalRule2,26)&"<br>"&mid(Sys_IllegalRule2,27,len(Sys_IllegalRule2))
		end If 
		Response.Write "(新臺幣"&Sys_Level2&"
	end if

	if int(Sys_Rule1)=5620001 then response.write "("&Sys_Note&")"
	response.write "</font>"
%></span></div>
<div id="Layer28" style="position:absolute; left:105px; top:<%=600+pagepx%>px; width:220px; height:15px; z-index:21"><span class="style33"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:140px; top:<%=625+pagepx%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:220px; top:<%=625+pagepx%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:310px; top:<%=625+pagepx%>px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<!-----------------------------------------法條編號 --------------------------------------------->
<div id="Layer32" style="position:absolute; left:520px; top:<%=670+pagepx%>px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	'response.write "　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		'response.write "　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer33" style="position:absolute; left:410px; top:<%=710+pagepx%>px; width:90px; height:40px; z-index:28"><span class="style3"><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></span></font></div>

<div id="Layer34" style="position:absolute; left:470px; top:<%=700+pagepx%>px; width:400px; height:30px; z-index:2"><%
	if showBarCode then	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"
%></div>
<!----'smith                     舉發單位章	--->
<div id="Layer35" style="position:absolute; left:420px; top:<%=760+pagepx%>px; width:100px; height:89px; z-index:29"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" height=40 align=""center""><span class=""style3""><font color='red'>"
	Response.Write SysUnit&Sys_UnitName
	Response.Write "</font></span><br><span class=""style3""><font color='red'>"
	Response.Write Sys_UnitTEL
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
<div id="Layer37" style="position:absolute; left:600px; top:<%=750+pagepx%>px; width:200px; height:46px; z-index:31"><%
	response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
	response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">警員&nbsp;"&Sys_ChName&"</span></td></tr>"
	response.write "</table>"		
%></div>

<!-----------------------------------填單日---------------------------------->
<div id="Layer38" style="position:absolute; left:130px; top:<%=810+pagepx%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:190px; top:<%=810+pagepx%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:250px; top:<%=810+pagepx%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<!-----------------------------------檢查用車號---------------------------------->
<div id="Layer41" style="position:absolute; left:355px; top:<%=930+pagepx%>px; width:90px; height:11px; z-index:34"><%=Sys_CarNo%></div>
<!-----------------------------------送達證書 繳回單位資料---------------------------------->
<div id="Layer42" class="style3" style="position:absolute; left:50px; top:<%=950+pagepx%>px; z-index:1"><%=SysUnit&Sys_UnitName%></div>

<div id="Layer45" class="style3" style="position:absolute; left:340px; top:<%=950+pagepx%>px; z-index:1"><%=SysUnit&Sys_UnitName%></div>

<div id="Layer46" class="style3" style="position:absolute; left:560px; top:<%=969+pagepx%>px; z-index:1"><%=InstrAdd(SysUnitAddress,14)%></div>

<!-----------------------------------送達證書 收件人資料---------------------------------->
<div id="Layer47" class="style3" style="position:absolute; left:115px; top:<%=980+pagepx%>px; z-index:1"><%
	response.write funcCheckFont(Sys_Owner,20,1)&"<br>"
	If instr(Sys_DriverHomeAddress,"@") >0 Then
		response.write funcCheckFont(Sys_DriverHomeZip&Sys_DriverZipName&Sys_DriverHomeAddress,20,1)
	else
		response.write InstrAdd(Sys_DriverHomeZip&Sys_DriverZipName&Sys_DriverHomeAddress,14)
	End if
%>
</div>
<!-----------------------------------送達證書 文 號 ---------------------------------->
<div id="Layer48" class="style3" style="position:absolute; left:225px; top:<%=1020+pagepx%>px; z-index:1"><%=Sys_BillNo%>
</div>

<!--
<div id="Layer48" style="position:absolute; left:400px; top:<%=1040+pagepx%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
-->

<div id="Layer49" class="style3" style="position:absolute; left:510px; top:<%=1035+pagepx%>px; z-index:1"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
</div>
</div><%
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