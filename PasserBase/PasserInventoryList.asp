<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=31"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then 
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
	end if
end if 
rsUInfo.close
set rsUInfo=nothing

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

strSQL="select InventoryNo," & _
		"(select (select UnitName from Unitinfo where Unitid=pb.memberstation) UitName from PasserBase pb where sn=PasserCreditor.billsn) UitName," & _
		"(select (select UnitOrder from Unitinfo where Unitid=pb.memberstation) UnitOrder from PasserBase pb where sn=PasserCreditor.billsn) UnitOrder," & _
		"(select BillNo from PasserBase pb where sn=PasserCreditor.BillSN) BillNo," & _
		"(select DriverID from PasserBase pb where sn=PasserCreditor.BillSN) DriverID," & _
		"(select BILLFILLDATE from PasserBase pb where sn=PasserCreditor.BillSN) BILLFILLDATE," & _
		"(select DEALLINEDATE from PasserBase pb where sn=PasserCreditor.BillSN) DEALLINEDATE," & _
		"(select (select illegalRule from Law where version=2 and itemid=pb.Rule1 and rownum=1) Rule1 from PasserBase pb where sn=PasserCreditor.BillSN) Rule1," & _
		"(select (select illegalRule from Law where version=2 and itemid=pb.Rule2 and rownum=1) Rule2 from PasserBase pb where sn=PasserCreditor.BillSN and Rule2 is not null) Rule2," & _
		"(select Driver from PasserBase pb where sn=PasserCreditor.BillSN) Driver," & _
		"(select (forfeit1+nvl(forfeit2,0)) Forfeit1 from PasserBase pb where sn=PasserCreditor.BillSN) Forfeit1," & _
		"nvl((select sum(nvl(PayAmount,0)) as PaySum from PasserPay where BillSN=PasserCreditor.BillSN),0) PaySum," & _
		"(select opengovNumber from PasserSendDetail where sn=PasserCreditor.SendDetailSN) opengovNumber" & _
		" from PasserCreditor where Exists(select 'Y' from "&BasSQL&" where SN=PasserCreditor.billsn) and InventoryNo is not null" & _
		" order by UnitOrder,UitName,Billno"

set rs=conn.execute(strSQL)

If not rs.eof Then sysUit=rs("UitName")
nowpage=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>交付保管品核對清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<%
While not rs.eof
	nowpage=nowpage+1

	chkUit = rs("UitName")

	If nowpage > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><strong>債權憑證與交付保管品核對清冊</strong></td>
			</tr>
			<tr>
				<td align="left">機關名稱：<%=chkUit%></td>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td>保管品收據編號<br>承辦單位</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>罰鍰名稱<br>受處罰人</td>
					<td>罰單餘額</td>
					<td>執行名義文號</td>
				</tr>
				<%
				For i = 1 to 10
					If rs.eof Then exit For 

					If chkUit <> rs("UitName") Then exit For 
					
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write rs("InventoryNo")
					If not ifnull(rs("UitName")) Then Response.Write "<br>"&rs("UitName")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("BillNo")
					If not ifnull(rs("DriverID")) Then Response.Write "<br>"&rs("DriverID")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write gInitDT(rs("BILLFILLDATE"))
					If not ifnull(rs("DEALLINEDATE")) Then Response.Write "<br>"&gInitDT(rs("DEALLINEDATE"))
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("Rule1")
					If not ifnull(rs("Rule2")) Then Response.Write "<br>"&rs("Rule2")
					If not ifnull(rs("Driver")) Then Response.Write "<br>"&rs("Driver")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write (cdbl(rs("Forfeit1"))-cdbl(rs("PaySum")))
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("opengovNumber")
					Response.Write "&nbsp;</td>"


					Response.Write "</tr>"


					rs.movenext
				Next
				%>
			</table>
		</td>
	</tr>
</table>
<%
wend
rs.close

If nowpage = 0 Then

strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsuit=conn.execute(strSQL)

chkUit=rsuit("UnitName")

rsuit.close
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><strong>債權憑證與交付保管品核對清冊</strong></td>
			</tr>
			<tr>
				<td align="left">機關名稱：<%=chkUit%></td>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td>保管品收據編號<br>承辦單位</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>罰鍰名稱<br>受處罰人</td>
					<td>罰單餘額</td>
					<td>執行名義文號</td>
				</tr>
				<%
					
					Response.Write "<tr>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

					Response.Write "<tr>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

					Response.Write "<tr>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

				%>
			</table>
		</td>
	</tr>
</table>
<%
End if 
%>
</body>
</html>
<%
conn.close
set conn=nothing
%>