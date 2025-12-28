<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/randomString.asp" --> 
<%
if trim(request("SN"))<>"" then
	Session("Map_SN")=request("SN")
	ProgressID = gen_key(10)
else
	Response.write "<script>"
	Response.Write "alert('資料錯誤請重新執行!!');"
	response.write "self.close();"
	Response.write "</script>"
	response.end
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>送達掃描上傳</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
.btn3{
   font-size:14px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
</style>
</head>
<body>
<form name="myForm2" method="post" action="upfile.asp" enctype="multipart/form-data" onSubmit="myOpen(this);">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>送達掃描上傳</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99" rowspan=4><font size=4>上傳格式<br><font color="red">(*.jpg,*.jpeg,*.tif,*.pdf)</font></td>
					<td>
						檔案:<input type="file" name="file1" style="width:400" class="btn1" value="" size="50">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input type="submit" name="Submit" class="btn3" style="width:40px;height:20px;" value="上傳">
			<input type="button" name="Submit2" class="btn3" style="width:40px;height:20px;" value="關閉" onclick="self.close();">
			<input type="Hidden" name="SN" value="<%=request("SN")%>">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
</form>
</body>
</html>
<script> 
	function myOpen(form){ 
		window.open("../Common/fileUpProgress.asp?progressID=<%=ProgressID%>","","width=500,height=200,scrollbars=no,toolbar=no,status=no,resizable=no,menubar=no,location=no"); 
		var url=form.action; 

		if (url.indexOf("?",0)==-1) { 
			form.action = url+"?progressID=<%=ProgressID%>"; 
		}else{ 
			form.action = url+"&progressID=<%=ProgressID%>"; 
		} 
	} 
</script> 