<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>送達紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
Server.ScriptTimeout=6000

strSQL="select billsn from passerSend where (select count(1) cnt from PasserSendDetail where BillSN=passerSend.BillSN)=0 and (select count(1) cnt from passerbase where recordstateid=0 and billstatus<>9 and sn=passersend.billsn)=1"
set rssn=conn.execute(strSQL)

while Not rssn.eof

	strSQL="select count(1) cnt from PasserSendDetail where BillSN="&trim(rssn("billsn"))

	set rscnt=conn.execute(strSQL)

	If cdbl(rscnt("cnt"))=0 Then
		strSQL="select OpenGovNumber,SendNumber,SendDate from PasserSend where billsn="&trim(rssn("billsn"))
		set rssend=conn.execute(strSQL)

		If not rssend.eof Then

			strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(rssn("billsn"))&",'"&trim(rssend("OpenGovNumber"))&"','"&trim(rssend("SendNumber"))&"',"&funGetDate(rssend("SendDate"),0)&",sysdate,"&Session("User_ID")&")"

			conn.execute(strSQL)
		End if
		rssend.close
	End if
	rscnt.close

	rssn.movenext
wend

rssn.close
%>
<BODY>
<form name=myForm method="post">
已處理
</form>
</BODY>
</HTML>
<%
conn.close
set conn=nothing
%>