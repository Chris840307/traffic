<!-- #include file="Common\db.ini" -->
<!-- #include file="Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body>

<table border="1" width="100%" id="table1">
	<tr>
<%
strSQL="select * from Notice where startdate <=sysdate and enddate > sysdate and recordstateid<>-1 order by startdate desc "

set rssysinfo=conn.execute(strSQL)

While Not rssysinfo.Eof
  response.write "<td>"
  response.write rssysinfo("NoticeData")
  response.write "</td>"  
  response.write "<tr>"    
  rssysinfo.MoveNext
Wend
%>
	</tr>
</table>

</body>

</html>
