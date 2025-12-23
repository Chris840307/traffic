<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 6800
%>
<%
'權限
'AuthorityCheck(234)
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

%>
<%

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--


-->
</style>
<title>有效清冊</title>
</head>
<body>
<form name="myForm" method="post">
	<table width="600" border="1" cellpadding="2" cellspacing="0" align="center">
		<tr>
			<td align="center" colspan="2" bgcolor="#FFCC33">
				<font size="3">有效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFFF99" width="30%" height="40">
				舉發單類型
			</td>
			<td >
				<select name="BillType">
					<option value="0">攔停 + 逕舉</option>
					<option value="1">攔停</option>
					<option value="2">逕舉</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFCC00" colspan="2">
				<input type="button" value="產生清冊" onclick="StopSendList();"> &nbsp;
				<input type="button" value="離開" onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="SQLstr" value="<%=Trim(request("SQLstr"))%>">
			</td>
		</tr>
	</table>
	
</form>
</body>
</html>
<script language="javascript">
function StopSendList(){
	window.open("ValidSendList_Excel_ML.asp?SQLstr=<%=Trim(request("SQLstr"))%>&BillType="+myForm.BillType.value,"inputWin123","width=900,height=500,left=50,top=0,scrollbars=yes,menubar=yes,resizable=yes,status=yes,toolbar=yes");
}
</script>
<%conn.close%>