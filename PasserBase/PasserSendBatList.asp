<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

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

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"

strSQL="select sn from PasserBase where Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN) "&trim(request("orderstr"))
set rs=conn.execute(strSQL)
BillSN=""
While Not rs.eof
	If Not ifnull(BillSN) Then BillSN=BillSN&","
	BillSN=BillSN&rs("sn")
	rs.movenext
Wend
rs.close
BillSN=Split(BillSN,",")

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

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

strSQL="select a.chName,b.Content from MemberData a,(select ID,Content from Code where TypeID=4 ) b where chname='"&trim(session("Sys_SendChName"))&"' and accountstateid=0 and a.JobID=b.ID(+)"
set mem=conn.execute(strSQL)
If not mem.eof Then
	chName=mem("chName")
	JobName=mem("Content")
end if
If ifnull(JobName) Then jobName="警員"
mem.close

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

If sys_City="彰化縣" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
end If 

for i=0 to Ubound(BillSN)
	if cint(i)<>0 then response.write "<div class=""PageNext"">&nbsp;</div>"%>
	<!--#include virtual="traffic/PasserBase/PaseBillPrit96_not_bat.asp"-->
	<%
	If sys_City="台南市" or sys_City="屏東縣" Then
		response.write "<div class=""PageNext"">&nbsp;</div>"%>
		<!--#include virtual="traffic/PasserBase/PaseBillPrit96_not_paper_bat.asp"--><%
	End if
Next
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	//window.focus();
	//printWindow(true,25,10,10,10);
</script>