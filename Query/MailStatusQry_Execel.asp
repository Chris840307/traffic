<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_郵寄狀態查詢.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	
	strSQL=request("TempSQL")

	set rsfound=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>郵寄狀態查詢</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>郵寄狀態查詢</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td align="center">單號</td>
					<td align="center">車號</td>
					<td align="center">郵件號碼</td>
					<td align="center">郵寄日期</td>					
					<td align="center">郵寄退回日期</td>
					<td align="center">退件原因</td>
					<td align="center">處理日期</td>
					<td align="center">郵件狀態</td>
					<td align="center">處理郵局</td>
				</tr>
				<%

					while Not rsfound.eof
						response.write "<tr bgcolor='#FFFFFF' align='center'>"
							response.write "<td class=""font10"">&nbsp;"&rsfound("Billno")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&rsfound("CarNo")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&rsfound("Mailnumber")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDateTime(rsfound("Maildate"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDate(rsfound("MailReturndate"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetResonName(rsfound("ReturnResonID"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDate2(rsfound("ProcDate"))&" "&trim(rsfound("ProcTime"))&"</td>"
							response.write "<td class=""font10"">&nbsp;"&trim(rsfound("MailStatus"))&"</td>"
							response.write "<td class=""font10"">&nbsp;"&trim(rsfound("HandleBrueau"))&"</td>"
						rsfound.movenext
					wend
				%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>