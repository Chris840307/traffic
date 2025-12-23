<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_"&Trim(request("date1"))&"建檔日_攔停_舉發單資料.xls"

Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
'Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
'response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000
Response.flush
%>
<%


strSql="select x.BillNo,x.CarNo,x.IllegalDate,y.IllegalDate as StopTime" &_
	" from (select carno,illegaldate,imagefilename,imagepathname from billbase where billno is null " &_
	" and rule1='5620001' and recordstateid=0)y, billbase x " &_
	" where x.sn in ("&trim(request("BillSN"))&") and x.rule1='5620001' and x.recordstateid=0 " &_
	" and x.carno=y.carno and x.imagefilename like '%'||y.imagepathname||'%' and y.imagepathname is not null"
set rsfound=conn.execute(strSql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單查詢</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<table width="100%" border="1">
<tr>
	<td colspan="6" align="center">對照表</td>
</tr>
<tr>
<td width="130">項次</td>
<td width="130">車號</td>
<td width="130">違規單號</td>
<td width="130">違規時間</td>
<td width="130">停車時間</td>
<td width="100">備註</td></tr>
<%  cnt=0
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					'Response.flush
					cnt=cnt+1
%>
	<tr>
		<td><%=cnt%></td>
		<td><%=trim(rsfound("CarNo"))%></td>
		<td><%=trim(rsfound("BillNo"))%></td>
		<td><%=year(trim(rsfound("IllegalDate")))-1911&"/"&right("00"&month(trim(rsfound("IllegalDate"))),2)&"/"&right("00"&Day(trim(rsfound("IllegalDate"))),2)&" "&right("00"&hour(trim(rsfound("IllegalDate"))),2)&":"&right("00"&Minute(trim(rsfound("IllegalDate"))),2)%></td>
		<td><%=year(trim(rsfound("StopTime")))-1911&"/"&right("00"&month(trim(rsfound("StopTime"))),2)&"/"&right("00"&Day(trim(rsfound("StopTime"))),2)&" "&right("00"&hour(trim(rsfound("StopTime"))),2)&":"&right("00"&Minute(trim(rsfound("StopTime"))),2)%></td>
		<td>&nbsp;</td>
	</tr>
<%
				rsfound.MoveNext
				Wend
				rsfound.close
				set rsfound=nothing
				%>
				
</body>
</html>
<%
conn.close
set conn=nothing
%>