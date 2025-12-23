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
<!--#include virtual="traffic/Common/css.txt"-->
<title>舉發單信封黏貼標籤列印位置設定</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">
		<br><br>
		<table width='100%' border='0' bgcolor="#CCCCCC" cellpadding="1" cellspacing="2">
			<tr>
				<td colspan="2" height="27" bgcolor="#FFCC33" class="pagetitle">
					<strong>舉發單信封黏貼標籤列印位置設定</strong>
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td nowrap>列印位置設定</td>

				<td nowrap>
					從第
					<input name="Sys_RecordDate1" type="text" class="btn1" value="<%=request("Sys_RecordDate1")%>" size="8" maxlength="7">
					<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_RecordDate1');">
					個黏貼標籤開始列印
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#EBFBE3" align="center">
					<input type="button" value="產  生" onclick="funMailListCity();">
					<input type="button" value="關  閉" onclick="self.close();">
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">

function funMailListCity(){
	if (myForm.Sys_RecordDate1.value=='' || myForm.Sys_RecordDate2.value==''){
		alert("請輸入匯入日期!");
	}else{


		UrlStr="StopBillReturnList.asp";
		myForm.action=UrlStr;
		myForm.target="BillReturnList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		self.close();
	}
}
</script>
</html>
