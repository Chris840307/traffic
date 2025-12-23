<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style4 {
	font-size: 18px
}
-->
</style>
<title>舉發單批號</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
	<table width="100%" border='1' align="left" cellpadding="1">
		<tr>
			<td bgcolor="#FFCC33" align="center" height="25"><strong>批號</strong></td>
			<td><strong><%=trim(request("BatchNo"))%></strong></td>
		</tr>
		<tr>
			<td bgcolor="#FFCC33" align="center" height="25">起始單號</td>
			<td><%=trim(request("FistNo"))%></td>
		</tr>
		<tr>
			<td bgcolor="#FFCC33" align="center" height="25">結束單號</td>
			<td><%=trim(request("LastNo"))%></td>
		</tr>
		<tr>
			<td bgcolor="#FFCC33" align="center" height="25">合計筆數</td>
			<td><%=trim(request("BillCount"))%></td>
		</tr>
<%
	If Trim(request("BatchNoQryCar"))<>"" Then
%>
		<tr>
			<td bgcolor="#FFCC33" align="center" height="25">車籍查詢批號</td>
			<td><%=trim(request("BatchNoQryCar"))%></td>
		</tr>
<%
	End if
%>
		<tr>
			<td colspan="2" align="center">
				<input type="button" value="關閉" onclick="window.close();">
			</td>
		</tr>
	</table>
</form>
</body>
</html>
