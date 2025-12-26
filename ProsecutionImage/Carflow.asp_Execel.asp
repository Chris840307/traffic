<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="..\Common\DB.ini"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_車流統計.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

strSQL="select StartPassingDate,EndPassingDate,Passings,Violations,MinSpeed,MaxSpeed from Passing "&trim(request("SQLstr"))

set rsfound=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>車流統計</title>
<!-- #include file="..\Common\css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td height="26" align="center"><strong>車流統計列表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr align="center">
					<td>開始日期</td>
					<td>結束時間</td>
					<td>通過車輛</td>
					<td>違規車輛</td>
					<td>最低時速</td>
					<td>最高時速</td>
				</tr><%
					while Not rsfound.eof
						response.write "<tr>"
						response.write "<td>"&rsfound("StartPassingDate")&"&nbsp;</td>"
						response.write "<td>"&rsfound("EndPassingDate")&"&nbsp;</td>"
						response.write "<td>"&rsfound("Passings")&"&nbsp;</td>"
						response.write "<td>"&rsfound("Violations")&"&nbsp;</td>"
						response.write "<td>"&rsfound("MinSpeed")&"&nbsp;</td>"
						response.write "<td>"&rsfound("MaxSpeed")&"&nbsp;</td>"
						response.write "</tr>"
						rsfound.movenext
					wend%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>