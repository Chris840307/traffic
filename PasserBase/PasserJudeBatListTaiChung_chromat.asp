<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'fMnoth=month(now)
'if fMnoth<10 then
'fMnoth="0"&fMnoth
'end if
'fDay=day(now)
'if fDay<10 then
'fDay="0"&fDay
'end if
'fname=year(now)&fMnoth&fDay&"_批次文件.doc"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>批次輸出系統</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	line-height:2;
}
.style2 {font-size: 18px; font-family: "標楷體"; line-height:2;}
.style3 {font-size: 18px; line-height:2;}
.style4 {font-family: "標楷體"; line-height:2;}
.style5 {font-size: 18px; line-height:2;}
.style6 {font-family: "標楷體"; font-size: 18px; line-height:2; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
	line-height:2;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
	line-height:2;
}
.style9 {font-family: "標楷體"; line-height:2;}
.style10 {font-size: 16px; line-height:2;}
.style11 {font-size: 14px; line-height:2;}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
	line-height:2;
}
.style13 {font-size: 14px; font-family: "標楷體"; line-height:2; }
.style14 {
	font-size: 22px;
	font-family: "標楷體";
	line-height:1;
}
.style15 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style16 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style17 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style18 {font-family: "標楷體"; font-size: 20px; line-height:2; }
.style19 {font-size: 24px; line-height:2; }
.style20 {font-size: 36px; line-height:2; }
.style21 {font-size: 18px; line-height:2; }
.style22 {font-family: "標楷體"; font-size: 18px;}
.style23 {font-family: "標楷體"; font-size: 14px;}
.style24 {font-family: "標楷體"; font-size: 12px;}
.style25 {font-family: "標楷體"; font-size: 24px;}
.style26 {font-family: "標楷體"; font-size: 10px;}
-->
</style>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close
set rsUInfo=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

thenPasserUnit=""
strSQL="select UnitID,UnitTypeID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
	Sys_GroupUnitName=trim(rsunit("UnitName"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsunit=conn.Execute(strSQL)
if Not rsunit.eof then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_Tel=trim(rsunit("Tel"))
	thenPasserUnit=trim(rsunit("UnitName"))
end if
rsunit.close

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
else
	Sys_SendBillSN=request("hd_BillSN")
End if
BillSN=Split(Sys_SendBillSN,",")
BillState=""
for i=0 to Ubound(BillSN)
	For k=0 to 2
		If k=0 Then
			if trim(request("Sys_PasserJude"))="1" then '裁決書
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserJudeDeliverTaiChung_chromat.asp"-->
					</Div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserJudeDeliverTaiChung_chromat.asp"-->
					</Div><%
				end if
			end if
		elseif k=1 then
			if trim(request("Sys_PasserUrge"))="1" then '催繳書
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserUrgeDeliverTaiChung_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserUrgeDeliverTaiChung_chromat.asp"-->
					</div><%
				end if
			end if
		elseif k=2 then
			if trim(request("Sys_PasserDeliver"))="1" then '送達證書
				DeliverKind=1
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				end if
			elseif trim(request("Sys_PasserDeliver"))="2" then
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				end if
			end if
		end if
	next
next
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,4.23,4.23,4.23,4.23);
</script>