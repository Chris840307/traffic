<%@ CODEPAGE="65001"%>
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
'fname=year(now)&fMnoth&fDay&"_批次裁決文件.doc"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<title>批次裁決輸出系統</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
.noprint {display:none;}
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
.style11 {font-size: 14px;}
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
.style22 {font-family: "標楷體"; font-size: 16px;}
.style22A {font-family: "標楷體"; font-size: 16px;text-align:justify; text-justify:distribute-all-lines; text-align-last:justify; text-indent:1px;width:93px}
.style22B {font-family: "標楷體"; font-size: 16px;text-align:justify; text-justify:distribute-all-lines; text-align-last:justify; text-indent:1px;width:130px}
.style23 {font-family: "標楷體"; font-size: 15px;}
.style24 {font-family: "標楷體"; font-size: 12px;}
.style25 {font-family: "標楷體"; font-size: 24px;}
.style26 {font-family: "標楷體"; font-size: 10px;}
.style27 {font-size: 12px;}
.style28 {font-family: "標楷體"; font-size: 20px; color:#3333ff;}
.style29 {font-family: "標楷體"; font-size: 45px; color:#3333ff;}
.style30 {font-family: "標楷體"; font-size: 21px; color:#3333ff;}
-->
</style>
</head>
<body>
<%
If ifnull(Session("Unit_ID")) Then Session("Unit_ID")=trim(Request("JpgUnitID"))

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
	end if 
end if 
rsUInfo.close
set rsUInfo=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

UrgeDate="裁決日期"
UrgeNo="字第"
Papertype="裁決書"
If not ifnull(request("BillUrge")) Then
	UrgeNo="交催字第"
	UrgeDate="催繳日期"
	Papertype="催繳書"
end if

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
else
	Sys_SendBillSN=request("hd_BillSN")
End if
strSQL="select sn from PasserBase where sn in("&Sys_SendBillSN&") "&trim(request("orderstr"))
set rs=conn.execute(strSQL)
BillSN=""
While Not rs.eof
	If Not ifnull(BillSN) Then BillSN=BillSN&","
	BillSN=BillSN&rs("sn")
	rs.movenext
Wend
rs.close
BillSN=Split(Sys_SendBillSN,",")

'strSQL="select DriverID,Driver,DriverSex,DriverAddress,count(DriverID) as affair,sum(FORFEIT1) as FORFEIT1 from PasserBase where SN in ("&Sys_SendBillSN&") Group by DriverID,Driver,DriverSex,DriverAddress"
'set rssum=conn.execute(strSQL)
BillState=""
'while Not rssum.eof
for i=0 to Ubound(BillSN)
	strSQL="select DriverID,Driver,DriverSex,DriverZip,DriverAddress,'1' as affair,FORFEIT1,IllegalDate,memberstation from PasserBase where SN ="&BillSN(i)
	set rssum=conn.execute(strSQL)

	strSQL="select WordNum from UnitInfo Where UnitID='"&rssum("memberstation")&"'"
	set rs=conn.execute(strSQL)
	If not rs.eof Then
		If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
	end if
	rs.close

	strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&rssum("memberstation")&"'"
	set rsUnit=conn.execute(strSQL)
	Sys_UnitID=trim(rsUnit("UnitID"))
	Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
	Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
	rsUnit.close

	If Sys_UnitLevelID=1 Then
		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"

		If sys_City="台南市" and trim(Sys_UnitID)="07A7" Then
			strSQL="select * from UnitInfo where UnitID='0707'"
		End if
		
	else
		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
	end if

	set unit=conn.Execute(strSQL)
	If Not unit.eof Then
		theUnitID=trim(unit("UnitID"))
		if trim(unit("UnitName"))<>"" and not isnull(unit("UnitName")) then
			theUnitName=replace(trim(unit("UnitName")),"台","臺")
		end if 
		theSubUnitSecBossName=trim(unit("SecondManagerName"))
		theBigUnitBossName=trim(unit("ManageMemberName"))
		theContactTel=trim(unit("Tel"))
		theBankAccount=trim(unit("BankAccount"))
		theBankName=trim(unit("BankName"))
		theUnitAddress=trim(unit("Address"))
	end if
	unit.close

	if trim(request("Sys_PasserNotify"))="1" then '交辦單
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/Paser_UrgeView.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/Paser_UrgeView.asp"--><%
		end if
	end if
	if trim(request("Sys_PasserUrge"))="1" then '催繳書
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserUrge_BatWord.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserUrge_BatWord.asp"--><%
		end if
	end if
	if trim(request("Sys_PasserSend"))="1" then '寄 存 送 達 通 知 書
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserDeliver.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserDeliver.asp"--><%
		end if
	end if
	if trim(request("Sys_PasserJudeSend"))="1" then '寄 存 送 達 通 知 書
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeSendDeliver.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeSendDeliver.asp"--><%
		end if
	end if
	'rssum.movenext
'wend
rssum.close

'for i=0 to Ubound(BillSN)

if trim(request("Sys_PasserJude"))="1" then '裁決書
	If sys_City = "花蓮縣" Then
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_HuaLien.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_HuaLien.asp"--><%
		end If 
	elseIf sys_City = "台南市" Then
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver.asp"--><%
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliverPaper.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver.asp"--><%
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliverPaper.asp"--><%
		end if	
	else
		if BillState<>"" then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_Pdf_miaoli.asp"--><%
		else
			BillState="1"%>
			<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_Pdf_miaoli.asp"--><%
			'response.write "<br><br>"
		end if
	end if
end if

if trim(request("Sys_PasserJude_Label"))="1" then '裁決通知書(保防版
	if BillState<>"" then
		response.write "<div class=""PageNext"">&nbsp;</div>"%>
		<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_label.asp"--><%
	else
		BillState="1"%>
		<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_label.asp"--><%
		'response.write "<br><br>"
	end if
end if

if trim(request("Sys_PasserSign"))="1" then '簽辦書
	if BillState<>"" then
		response.write "<div class=""PageNext"">&nbsp;</div>"%>
		<!--#include virtual="traffic/PasserBase/Passer_SignView.asp"--><%
	else
		BillState="1"%>
		<!--#include virtual="traffic/PasserBase/Passer_SignView.asp"--><%
		'response.write "<br><br>"
	end if
end if

if trim(request("Sys_PasserDeliver"))="1" then '郵務送達證書
	if BillState<>"" then
		response.write "<div class=""PageNext"">&nbsp;</div>"%>
		<!--#include virtual="traffic/PasserBase/BillBase_Deliver.asp"--><%
	else
		BillState="1"%>
		<!--#include virtual="traffic/PasserBase/BillBase_Deliver.asp"--><%
	end if
end If 

if trim(request("Sys_PasserLabel_miaoli"))="1" then '苗栗保防標籤
	if BillState<>"" then
		response.write "<div class=""PageNext"">&nbsp;</div>"%>
		<!--#include virtual="traffic/PasserBase/PasserLabel_miaoli.asp"--><%
		If i < Ubound(BillSN) Then 
			i=i+1
			Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
		%>
			<!--#include virtual="traffic/PasserBase/PasserLabel_miaoli.asp"--><%
		end if
	else
		BillState="1"%>
		<!--#include virtual="traffic/PasserBase/PasserLabel_miaoli.asp"--><%
		If i < Ubound(BillSN) Then
			i=i+1
			Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
		%>
			<!--#include virtual="traffic/PasserBase/PasserLabel_miaoli.asp"--><%
		end if
	end if
end If 

If i < Ubound(BillSN) Then response.write "<div class=""PageNext"">&nbsp;</div>"
BillState=""
Next
if trim(request("Sys_PasserDeliver"))="1" then '裁決書
	response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<!--#include virtual="traffic/PasserBase/PasserBaseUrgeJudeList_word.asp"-->
<%
elseif trim(request("Sys_PasserSend"))="1" then
	response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<!--#include virtual="traffic/PasserBase/PasserBaseSendList_word.asp"-->
<%
end if
%>

</body>
</html>
