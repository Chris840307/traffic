<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<!--#include virtual="traffic/Common/DCIURL.ini"-->
<% 
Function GetCDate(tmpDate)
GetCDate=CDbl(Mid(tmpDate,1,4))-1911 & "-" & Mid(tmpDate,5,2) & "-" & Mid(tmpDate,7,2) & " " & Mid(tmpDate,9,2) & ":" & Mid(tmpDate,11,2) & ":" & Mid(tmpDate,13,2)
End function
mailnum=Right("000000"&request("MailNum"),6) & "36400017"
mailnum2=Right("000000"&request("MailNum"),6) & "36400018"

sql="Select * from mailresult where mailnum in ('" & mailnum & "','" & mailnum2 & "') order by mailnum,ProcDateTime"

Set rs1=conn.execute(sql)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>郵寄歷程</title>
</head>

<body onload="window.focus();" onkeydown="if (event.keyCode==13) {event.keyCode=9; return event.keyCode }
		">
<form name=myForm method="post">
<script type="text/javascript" src="../js/date.js"></script>
	<table width='100%' border='1' cellpadding="1" id="table1">
		<tr bgcolor="#FFCC33">
			<td><strong>郵寄歷程資料</strong></td>
		</tr>
	</table>

	<table width='100%' border='1' cellpadding="2" id="table2">
	<tr>	
			<td colspan="18" bgcolor="#00FFFF" height="35"><b>郵件號碼：<font size="5" color="red"><%=trim(request("MailNum"))%></font></b>
			&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>郵件號碼：<font size="5" color="red"><%=mailnum%></font></b>
			</td>
			<% While Not rs1.eof
			tmpString=""
			if not isnull(rs1("T14")) Then tmpString="&#13;退件原因：" & rs1("T14") & rs1("T15")
			if not isnull(rs1("T16")) Then tmpString="&#13;退回局說明：" & rs1("T16") & rs1("T17")
			if not isnull(rs1("W14")) Then tmpString="&#13;通知日期時間：" & rs1("W14") 
			if not isnull(rs1("W15")) Then tmpString="&#13;留局原因：" & rs1("W15") & rs1("W16")
			if not isnull(rs1("G14")) Then tmpString="&#13;報值或保價金額：" & rs1("G14")
			if not isnull(rs1("G15")) Then tmpString="&#13;代收貨價金額：" & rs1("G15")
			if not isnull(rs1("G16")) Then tmpString="&#13;信箱號碼：" & rs1("G16")
			if not isnull(rs1("H14")) Then tmpString="&#13;處理方式：" & rs1("H14") & rs1("H15")
			if not isnull(rs1("H16")) Then tmpString="&#13;未妥投原因：" & rs1("H16") & rs1("H17")
			if not isnull(rs1("H18")) Then tmpString="&#13;招領局號：" & rs1("H18")
			if not isnull(rs1("H19")) Then tmpString="&#13;招領局名：" & rs1("H19")
			if not isnull(rs1("H20")) Then tmpString="&#13;招領局電話號碼：" & rs1("H20")
			if not isnull(rs1("N14")) Then tmpString="&#13;收件不成功原因：" & rs1("N14") & rs1("N15")
			if not isnull(rs1("Z14")) Then tmpString="&#13;接收局號：" & rs1("Z14")
			if not isnull(rs1("Z15")) Then tmpString="&#13;接收局名：" & rs1("Z15")
			if not isnull(rs1("Z16")) Then tmpString="&#13;清單號碼：" & rs1("Z16")
			%>
		<tr>
		<!--
			<td bgcolor="#FFFF99" align="right"><strong>交寄日期</strong></td>
			<td align="left">&nbsp;<%
					if not isnull(rs1("MailDate")) Then response.write trim(rs1("MailDate"))
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>郵費單</strong></td>
			<td align="left">&nbsp;<%
					if not isnull(rs1("MailNo")) Then response.write trim(rs1("MailNo"))
			%></td>
-->
			<td bgcolor="#FFFF99" align="right"><strong>處理時間</strong></td>
			<td align="left" title= "交寄日期：<%=rs1("MailDate")%>&#13;郵費單：<%=rs1("MailNo")%>">&nbsp;<%
					if not isnull(rs1("ProcDateTime")) Then response.write GetCDate(rs1("ProcDateTime"))
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>郵件狀態</strong></td>
			<td align="left">&nbsp;<%
					if not isnull(rs1("MailStateName")) Then response.write trim(rs1("MailStateCode"))&"："&trim(rs1("MailStateName"))
			%></td>
			<!--
			<td bgcolor="#FFFF99" align="right"><strong>郵局號</strong></td>
			<td align="left">&nbsp;<%
					if not isnull(rs1("PostNo")) Then response.write trim(rs1("PostNo"))
			%></td>
			-->
			<td bgcolor="#FFFF99" align="right"><strong>郵局名</strong></td>
			<td align="left"  title = "郵局電話：<%=Trim(rs1("PostTel"))%>&#13;郵局傳真：<%=rs1("PostFax")%>&#13;郵局地址:<%=rs1("PostAddr")%>">&nbsp;<%
					if not isnull(rs1("PostNo")) Then response.write trim(rs1("PostName"))
			%></td>

			
			<td bgcolor="#FFFF99" align="right"><strong>檔案名</strong></td>
			<td align="left" colspan="3" title = "<%=tmpString%>">&nbsp;<%
					if not isnull(rs1("FileName")) Then response.write trim(rs1("FileName"))
			%></td>
		</tr>
		<%
		rs1.MoveNext
		wend%>
			
			
		
		</table>

</form>

</body>

</html>
<%conn.close()%>