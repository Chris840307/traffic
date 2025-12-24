<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>欄停點收檔案匯入</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post" action="BillbaseChkUpfile.asp" enctype="multipart/form-data" >
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>欄停點收檔案匯入</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>匯入格式<font color="red">(*.xls)</font></td>
					<td>
						檔案<input type="file" name="file1" style="width:300" class="btn1" value="" size="30">
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
<script type="text/javascript" src="../js/date.js"></script>