<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_公告匯入檔.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

	strSQL="select billno,opengovnumber,opengovdate from billmailhistory where billsn in (select billsn from DCILog"&request("strDCISQL")&") order by BillSN"
	
	set rsfound=conn.execute(strSQL)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>公告匯入檔</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>

	<table border="1" cellpadding="4" cellspacing="1">
		<%
			ReturnMarkType=split("3,4,5,Y",",")
			ReturnMarkName=Split("單退,寄存,公示,撤消",",")
			while Not rsfound.eof
				response.write "<tr>"
				response.write "<td>"&rsfound("billno")&"</td>"
				response.write "<td style='mso-number-format:""\@"";'>"&rsfound("opengovnumber")&"</td>"
				response.write "<td style='mso-number-format:""\@"";'>"&gInitDT(rsfound("opengovdate"))&"</td>"
									
				response.write "</tr>"
				rsfound.movenext
			wend

			rsfound.close
		%>
	</table>
</body>
</html>
<%conn.close%>