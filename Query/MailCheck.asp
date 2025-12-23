<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<% 
Function GetCDateTime(tmpDate)
GetCDateTime=CDbl(Year(tmpDate))-1911 & "-" & Right("0"&Month(tmpDate),2) & "-" & Right("0"&day(tmpDate),2) & " " & Right("0"&hour(tmpDate),2) & ":" & Right("0"&minute(tmpDate),2) & ":" & Right("0"&second(tmpDate),2)
End Function

Function GetCDate(tmpDate)
GetCDate=CDbl(Year(tmpDate))-1911 & "-" & Right("0"&Month(tmpDate),2) & "-" & Right("0"&day(tmpDate),2) 
End function

BillNo=request("BillNo")
sql="select mailDate as tmpDate,'郵寄日' as tmpfield,null as tmpRecordDate from billmailhistory where billno='" & BillNo & "' and mailDate is not null"
sql=sql& " union all "
sql=sql& " select usermarkDate as tmpDate,'最後操作日' as tmpfield,usermarkDate as tmpRecordDate from billmailhistory where billno='" & BillNo & "' and usermarkDate is not null"
sql=sql& " union all "
sql=sql& " select opengovmailreturnDate as tmpDate,'公示日' as tmpfield,OpenGovRecordDate as tmpReCordDate from billmailhistory where billno='" & BillNo & "' and opengovmailreturnDate is not null"
sql=sql& " union all "
sql=sql& " select StoreAndSendMailDate as tmpDate,'寄存日' as tmpfield,StoreAndSendRecordDate as tmpReCordDate from billmailhistory where billno='" & BillNo & "' and StoreAndSendMailDate is not null"
sql=sql& " union all "
sql=sql& " select usermarkreturnDate as tmpDate,'usermarkreturnDate' as tmpfield,null as tmpRecordDate from billmailhistory where billno='" & BillNo & "' and usermarkreturnDate is not null"
sql=sql& " union all "
sql=sql& " select SignDate as tmpDate,'收受日' as tmpfield,null as tmpRecordDate from billmailhistory where billno='" & BillNo & "' and SignDate is not null"
sql=sql& " order by tmpRecordDate"
'response.write sql
Set rs1=conn.execute(sql)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>郵寄時間</title>
</head>

<body onload="window.focus();" onkeydown="if (event.keyCode==13) {event.keyCode=9; return event.keyCode }
		">
<form name=myForm method="post">
<script type="text/javascript" src="../js/date.js"></script>
	<table width='100%' border='1' cellpadding="1" id="table1">
		<tr bgcolor="#FFCC33">
			<td><strong>郵寄時間資料</strong></td>
		</tr>
	</table>

	<table width='100%' border='1' cellpadding="2" id="table2">
	<%While Not rs1.eof%>
	<tr>	
			<td bgcolor="#FFFF99" align="right"><strong><%=rs1("tmpfield")%></strong></td>
			<td align="left">&nbsp;<%=GetCDate(rs1("tmpDate"))%></td>
			<td align="left">&nbsp;<%
			If Not IsNull(rs1("tmpRecordDate")) then
				response.write GetCDateTime(rs1("tmpRecordDate"))
			End if
			%></td>
	</tr>
	<%rs1.MoveNext
	wend%>

			
		
		</table>

</form>

</body>

</html>
<%conn.close()%>