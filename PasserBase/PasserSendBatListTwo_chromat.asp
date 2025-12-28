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
.style2 {font-family: "標楷體"; font-size: 24px; line-height:1.5;}
.style3 {font-family: "標楷體"; font-size: 16px; }
.style4 {font-family: "標楷體"}
.style5 {font-size: 18px}
.style6 {font-family: "標楷體"; font-size: 12px; }
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

'tmpSN=""
'
'strSQL="select a.*,b.DriverID,b.cmt,(b.Forfeit-b.PayAmount) total from (select sn,BillNo,DoubleCheckStatus,Driver,IllegaLDate,Rule1 from PasserBase where sn in("&Sys_SendBillSN&")) a,(select DriverID,min(SN) sn,count(1) cmt,sum(Forfeit1) Forfeit,sum(nvl(PayAmount,0)) PayAmount from (select sn,DriverID,Forfeit1 from PasserBase where sn in("&Sys_SendBillSN&")) pas,(select billsn,sum(nvl(PayAmount,0)) PayAmount from PasserPay where billsn in("&Sys_SendBillSN&") group by billsn) pay where pas.sn=pay.billsn(+) group by DriverID) b where a.sn=b.sn "&Request("orderstr")
'tmp_paserSN="":Sys_Cmt="":Sys_Total=""
'set rsPasser=conn.execute(strSQL)
'While not rsPasser.eof
'	If not ifnull(tmpSN) Then
'		tmpSN=tmpSN&","
'		Sys_Cmt=Sys_Cmt&","
'		Sys_Total=Sys_Total&","
'	end if
'
'	tmpSN=tmpSN&trim(rsPasser("sn"))
'	Sys_Cmt=Sys_Cmt&trim(rsPasser("cmt"))
'	Sys_Total=Sys_Total&trim(rsPasser("total"))
'	
'	rsPasser.movenext
'Wend
'rsPasser.close
'
'Sys_Cmt=Split(Sys_Cmt,",")
'Sys_Total=Split(Sys_Total,",")
''BillSN=Split(Sys_SendBillSN,",")
'BillSN=Split(tmpSN,",")


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

strSQL="select a.chName,b.Content from MemberData a,(select ID,Content from Code where TypeID=4 ) b where chname='"&trim(session("Sys_SendChName"))&"' and accountstateid=0 and a.JobID=b.ID(+)"
set mem=conn.execute(strSQL)
If not mem.eof Then
	chName=mem("chName")
	JobName=mem("Content")
end if
If ifnull(JobName) Then jobName="警員"
mem.close

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
if not rsUInfo.eof then thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
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

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	DB_UnitID=trim(unit("UnitID"))
	if not isnull(unit("UnitName")) and trim(unit("UnitName"))<>"" then
		DB_UnitName=replace(replace(trim(unit("UnitName")),"交通組",""),"台","臺")
	end if 
	DB_Tel=trim(unit("Tel"))
	theSubUnitSecBossName=trim(unit("SecondManagerName"))
	theBigUnitBossName=trim(unit("ManageMemberName"))
	thePasserSendBankAccountName=trim(unit("PasserSendBankAccountName"))
	thePasserVATnumber=trim(unit("VATNUMBER"))
	thePasserSendBankAccount=trim(unit("PasserSendBankAccount"))
	thePasserSendBankName=trim(unit("PasserSendBankName"))
	theBankName=trim(unit("BankName"))
	theBankAccount=trim(unit("BankAccount"))
end if
unit.close

If ifnull(thePasserSendBankAccount) Then
	thePasserSendBankAccount=trim(theBankAccount)
	thePasserSendBankName=trim(theBankName)
End if

for i=0 to Ubound(BillSN)
	if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
	If trim(request("Sys_PasserNotify"))="1" Then%>
		<div id="L78" class="pageprint" style="position:relative;">
		<!--#include virtual="traffic/PasserBase/PaseBillPrit96_not_Two_bat.asp"-->
		</div><%
		If sys_City="台南市" Then
			response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<div id="L78" class="pageprint" style="position:relative;">
			<!--#include virtual="traffic/PasserBase/PaseBillPrit96_not_Two_paper_bat.asp"-->
			</div><%
		End if
	else%>
		<div id="L78" class="pageprint" style="position:relative;">
		<!--#include virtual="traffic/PasserBase/PaseBillPrit96Two_chromat.asp"-->
		</div><%
	End if
Next
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	//printWindow(true,5.50,5.50,5.50,5.50);
</script>