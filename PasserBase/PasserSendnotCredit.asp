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

strSQL="select (select UnitName from Unitinfo where Unitid=pb.memberstation) UitName," & _
		"(select UnitOrder from Unitinfo where Unitid=pb.memberstation) UnitOrder," & _
		"(select ChName from MemberData where MemberID=pb.RecordMemberID) UitChName," & _
		"(select SendDate from PasserSend where BillSN=pb.SN) SendDate," & _
		"BillNo,DriverID,BILLFILLDATE,DEALLINEDATE,Driver,DriverAddress,Rule1,Rule2,(forfeit1+nvl(forfeit2,0)) Forfeit1" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)" & _
		" and Exists(select 'Y' from passersendDetail where billsn=pb.sn and TRUNC(sysdate-senddate) > 365" & _
		" and not Exists(select 'Y' from PasserCreditor where senddetailsn=passersendDetail.sn))" & _
		" union all " & _
		"select (select UnitName from Unitinfo where Unitid=pb.memberstation) UitName," & _
		"(select UnitOrder from Unitinfo where Unitid=pb.memberstation) UnitOrder," & _
		"(select ChName from MemberData where MemberID=pb.RecordMemberID) UitChName," & _
		"(select SendDate from PasserSend where BillSN=pb.SN) SendDate," & _
		"BillNo,DriverID,BILLFILLDATE,DEALLINEDATE,Driver,DriverAddress,Rule1,Rule2,(forfeit1+nvl(forfeit2,0)) Forfeit1" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)" & _
		" and Exists(select 'Y' from passersend where billsn=pb.sn and TRUNC(sysdate-senddate) > 365" & _
		" and not Exists(select 'Y' from PasserCreditor where billsn=passersend.billsn))" & _
		" and pb.billstatus<>9 and pb.recordstateid=0" & _
		" order by UnitOrder,UitName,SendDate"

set rs=conn.execute(strSQL)

nowpage=0
chkUit=""
Uit_Sum=0
total_Sum=0

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>已移送執行逾特定期間未結案清冊</title>
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
				<td align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td align="center"><strong>行政罰鍰已移送執行逾特定期間未結案清冊</strong></td>
			</tr>
			<tr>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td nowrap>機關單位</td>
					<td>承辦人</td>
					<td nowrap>罰鍰單號<br>身分證號</td>
					<td nowrap>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td nowrap>違反法令<br>移送日期</td>
					<td nowrap>罰單金額</td>
					<td nowrap>處理情況</td>
				</tr>
				<%
				For i = 1 to 10
					If rs.eof Then exit For 

					If chkUit <> rs("UitName") Then exit For 

					Uit_Sum=Uit_Sum+1
					total_Sum=total_Sum+1
					
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write rs("UitName")
					Response.Write "</td>"

					Response.Write "<td nowrap>"
					Response.Write rs("UitChName")
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
					Response.Write rs("Driver")
					If not ifnull(rs("DriverAddress")) Then Response.Write "<br>"&rs("DriverAddress")
					Response.Write "</td>"


					Response.Write "<td>"
					Response.Write rs("Rule1")
					If not ifnull(rs("Rule2")) Then Response.Write "\"&rs("Rule2")
					Response.Write "<br>"&gInitDT(rs("SendDate"))
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("Forfeit1")
					Response.Write "</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"


					rs.movenext
				Next
				%>
			</table>
		</td>
	</tr>
	<%
		If not rs.eof Then
			If chkUit <> rs("UitName") Then
				chkUit=rs("UitName")

				Response.Write "<tr>"
				Response.Write "<td>小計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write Uit_Sum&"件<td>"
				Response.Write "<tr>"

				Uit_Sum=0
			end If 
		else
			Response.Write "<tr>"
			Response.Write "<td>小計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write Uit_Sum&"件<td>"
			Response.Write "<tr>"

			Response.Write "<tr>"
			Response.Write "<td>總計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write total_Sum&"件<td>"
			Response.Write "<tr>"
		End if 
	%>
	<tr>
		<td>
			備註：本表係針對已移送執行署後逾1年以上仍未取得債權憑證案件者，主要是提醒業務單位承辦人員及主管查明仍未取得債權憑證原因。
		</td>
	</tr>
	<tr>
		<td>
			主辦業務人員：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			業務主管：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
				<td align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td align="center"><strong>行政罰鍰已移送執行逾特定期間未結案清冊</strong></td>
			</tr>
			<tr>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td nowrap>機關單位</td>
					<td>承辦人</td>
					<td nowrap>罰鍰單號<br>身分證號</td>
					<td nowrap>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td nowrap>違反法令<br>移送日期</td>
					<td nowrap>罰單金額</td>
					<td nowrap>處理情況</td>
				</tr>
				<%

					
					Response.Write "<tr>"
					Response.Write "<td>"&chkUit&"</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

					Response.Write "<tr>"
					Response.Write "<td>"&chkUit&"</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

					Response.Write "<tr>"
					Response.Write "<td>"&chkUit&"</td>"
					Response.Write "<td>&nbsp;</td>"

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
	<%
		Response.Write "<tr>"
		Response.Write "<td>總計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "0件<td>"
		Response.Write "<tr>"
	%>
	<tr>
		<td>
			備註：本表係針對已移送執行署後逾1年以上仍未取得債權憑證案件者，主要是提醒業務單位承辦人員及主管查明仍未取得債權憑證原因。
		</td>
	</tr>
	<tr>
		<td>
			主辦業務人員：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			業務主管：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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