<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'fMnoth=month(now)
'if fMnoth<10 then fMnoth="0"&fMnoth
'fDay=day(now)
'if fDay<10 then	fDay="0"&fDay
'fname=year(now)&fMnoth&fDay&"_債權憑證清冊.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950"

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


strSQL="select count(1) cnt" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where sn=pb.sn) and billstatus=0 and RecordStateid=0 and not Exists(select 'Y' from PassersEndArrived where PasserSN=pb.sn)"

set rs=conn.execute(strSQL)

pageCnt=fix(cdbl(rs("cnt"))/10+0.99999999999999)
rs.close

strSQL="select (select UnitName from Unitinfo where Unitid=pb.memberstation) UitName," & _
		"(select UnitOrder from Unitinfo where Unitid=pb.memberstation) UnitOrder," & _
		"(select ChName from MemberData where MemberID=pb.RecordMemberID) UitChName," & _
		"BillNo,DriverID,BILLFILLDATE,DEALLINEDATE,Driver,DriverAddress,Rule1,Rule2,(forfeit1+nvl(forfeit2,0)) Forfeit1" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where sn=pb.sn) and billstatus=0 and RecordStateid=0 and not Exists(select 'Y' from PassersEndArrived where PasserSN=pb.sn)" & _
		" order by UnitOrder,UitName"

set rs=conn.execute(strSQL)

nowpage=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>未送達清冊</title>
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

	If nowpage > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td align="center"><strong>行政罰鍰未送達清冊</strong></td>
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
					<td>機關單位</td>
					<td>承辦人</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td>違反法令</td>
					<td>罰單金額</td>
					<td>處理情況</td>
				</tr>
				<%
				For i = 1 to 10
					If rs.eof Then exit For 
					
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write rs("UitName")
					Response.Write "</td>"

					Response.Write "<td>"
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
					If not ifnull(rs("Rule2")) Then Response.Write "<br>"&rs("Rule2")
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
	<tr>
		<td>
			備註：本表產生日期以開立罰鍰繳款書日起第7日，主要是提醒承辦人員，倘仍未送達應即採取補送程序，以免影響本府權益。
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
				<td align="center"><strong>行政罰鍰未送達清冊</strong></td>
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
					<td>機關單位</td>
					<td>承辦人</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td>違反法令</td>
					<td>罰單金額</td>
					<td>處理情況</td>
				</tr>
				<%
				For i = 0 to 2
					
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write chkUit
					Response.Write "</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"

				Next
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			備註：本表產生日期以開立罰鍰繳款書日起第7日，主要是提醒承辦人員，倘仍未送達應即採取補送程序，以免影響本府權益。
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