<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>T-SQL線上更新</title>
</head>
<body>
<form name="myForm" method="post" action="UpdateFile.asp" enctype="multipart/form-data" >
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>T-SQL線上更新</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>上傳格式<font color="red">(*.txt)</font></td>
					<td>
						檔案1:<input type="file" name="file1" style="width:400" class="btn1" value="" size="50">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input type="submit" name="Submit" value="上傳">
			<input type="button" name="Submit2" value="關閉" onclick="self.close();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
</form>
</body>
</html>