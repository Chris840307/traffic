<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>舉發單查詢</title>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return funBillQry();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="2">舉發單查詢</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發單號</td>
				<td>
					<input type="text" name="billno" value="" size="12" maxlength="9" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="2">
					<input type="button" name="okbtn" value="確 定" onClick="funBillQry();">
				</td>
			</tr>
		</table>		
	</form>
</body>
<script language="JavaScript">
function funBillQry(){
	if(myForm.billno.value!=''){
		window.opener.funBillNoQuery_Stop(myForm.billno.value);
	}
	self.close();
}
myForm.billno.focus();
</script>
</html>
