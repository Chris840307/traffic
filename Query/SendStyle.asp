<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印</title>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>舉發單列印</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>舉發單郵寄次數</font></td>
					<td><input class="btn1" type="radio" name="chkSend" value="第一次郵寄" checked>第一次郵寄
					<input class="btn1" type="radio" name="chkSend" value="第二次郵寄">第二次郵寄</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>附件</font></td>
					<td>
						<input class="btn1" type="radio" name="chkLabel" value="行政文書">附"送達證書"字樣<br>
						<input class="btn1" type="radio" name="chkLabel" value="更改通知書">附"更改通知書"字樣
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
myForm.chkSend[0].checked=true;
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funAdd(){
	var strLabel=''; var strUpLabel='';
	if(myForm.chkSend[0].checked){
		var strSend=myForm.chkSend[0].value;
	}else{
		var strSend=myForm.chkSend[1].value;
	}
	if(myForm.chkLabel[0].checked){
		strLabel=myForm.chkLabel[0].value;
	}else if(myForm.chkLabel[1].checked){
		strUpLabel=myForm.chkLabel[1].value;
	}
	window.opener.myForm.Sys_SendKind.value=strSend;
	window.opener.myForm.Sys_LabelKind.value=strLabel;
	window.opener.myForm.Sys_LabelUpdate.value=strUpLabel;
	window.opener.funLabelStyleKeelung();
}
</script>