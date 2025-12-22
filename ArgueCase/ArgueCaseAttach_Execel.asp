<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_申訴案件附件.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<%
if request("SQLstr")<>"" then
	set rs=conn.execute(request("SQLstr"))
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>申訴案件-附件資料</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>流水號</td>
					<td>附件物品</td>
					<td>備註</td>
				</tr>
				<%
				if request("SQLstr")<>"" then
					tempSN=0
					while Not rs.eof
						tempSN=tempSN+1
						response.write "<tr>"
						response.write "<td>"&tempSN&"</td>"
						response.write "<td>"&rs("AttachName")&"</td>"
						response.write "<td>"&rs("Note")&"</td>"
						response.write "</tr>"
						rs.movenext
					wend
					rs.close
				end if
				%>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>
<%conn.close%>