<%
asp="UploadFile.asp?type="&request("type")

 



 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>上傳列表</title>
<body>
<script>
	function SelectUpdateType(){
		myForm.submit();
	}

function DelData(FileName){
     if (confirm("是否刪除資料?\n\n請確認")==true){
		myForm.FileName.value=FileName;
		myForm.kinds.value="Del";
		myForm.submit();
	}
}

</script>
<form name="myForm" method="Post">
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>上傳列表</td>
			</td>
		</table>

		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
		<td bgcolor="#FFCC33" align="center">
		<input type="text" name="type" value="<%=request("type")%>">
		<br>
		<input type="button" value="違規照片上傳" name="btnClear" onclick="SelectUpdateType();"></td>


		</table>

<iframe frameborder="0" width="100%" height="800px"
           src="<%=asp%>">
這裡的文字只會出現在沒支援 iframe 的瀏覽器。
</iframe>
</form>
</html>
