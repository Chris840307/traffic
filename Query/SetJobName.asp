<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單簽章設定</title>
</head>
<body>
<form name=myForm method="post">
<table width="20%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>舉發單簽章設定</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>主管職稱</font></td>
					<td><input name="Sys_BillJobName" type="text" value="" size="8" class="btn1"></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>主管姓名</font></td>
					<td>
						<input name="Sys_MainChName" type="text" value="" size="8" class="btn1">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funAdd(){
	window.opener.myForm.hd_BillJobName.value=myForm.Sys_BillJobName.value;
	window.opener.myForm.hd_MainChName.value=myForm.Sys_MainChName.value;
	alert("設定完成!!");
	window.close();
}
window.resizeTo(250,300);
</script>