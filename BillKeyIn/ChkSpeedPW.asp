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
<title>超速密碼輸入</title>
<%


%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr >
				<td colspan="4">超速 100~150 公里,需輸入密碼才可建檔
				<br>
					<input type="password" name="SpeedPW" value="">
				</td>

			</tr>


			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="chk" value="確定" onclick="ChkPW();">
				<input type="button" name="close" value="重新輸入車速" onclick="SpeedErr();">
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function ChkPW(){
	if (myForm.SpeedPW.value=="0978"){
		opener.SpeedError=0;
		window.close();
	}else{
		alert("密碼輸入錯誤!!");
	}
}
function SpeedErr(){
	opener.myForm.RuleSpeed.value="";
	opener.myForm.IllegalSpeed.value="";
	opener.SpeedError=0;
	window.close();
}
</script>
</html>
