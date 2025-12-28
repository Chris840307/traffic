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

.style1 {font-family: "標楷體"; font-size: 14px; }
.style2 {font-family: "標楷體"; font-size: 25px; }
.style3 {font-family: "標楷體"; font-size: 16px; }
.style4 {font-family: "標楷體"}
.style5 {font-size: 18px}
.style6 {font-family: "標楷體"; font-size: 18px; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
}
.style9 {font-family: "標楷體"}
.style10 {font-size: 16px}
.style11 {font-size: 14px}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
}
.style13 {font-size: 14px; font-family: "標楷體"; }
.style14 {
	font-size: 30px;
	font-family: "標楷體";
}
.style15 {font-family: "標楷體"; font-size: 28px; }
.style16 {font-family: "標楷體"; font-size: 20px; }
.style17 {font-family: "標楷體"; font-size: 23px; }
.style18 {font-family: "標楷體"; font-size: 24px; }
.style19 {font-size: 24px}
.style20 {font-size: 36px}
.style21 {font-size: 18px}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then
	Sys_SendBillSN=request("hd_BillSN")
else
	Sys_SendBillSN=request("BillSN")
End if
BillSN=Split(Sys_SendBillSN,",")

strCity="select value from Apconfigure where id=52"
set rsCity=conn.execute(strCity)
theBillNumber=trim(rsCity("value"))
rsCity.close

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close
set rsUInfo=nothing

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
'strSQL="select * from UnitInfo where UnitLevelID=1 and UnitName like '%交%隊'"
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
theBigUnitBossName=trim(rsUnit("ManageMemberName"))
theContactTel=trim(rsUnit("Tel"))
theBankAccount=trim(rsUnit("BankAccount"))
rsUnit.close

for i=0 to Ubound(BillSN)
	if cint(i)<>0 then response.write "<div class=""PageNext""></div>"%>
	<div id="L78" class="pageprint" style="position:relative;">
	<!--#include virtual="traffic/PasserBase/PaseBillPrit96_chromat.asp"-->
	</div><%
Next
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.50,5.50,5.50,5.50);
</script>