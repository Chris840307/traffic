<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>批次列印</title>
</head>
<script>
function DataSubmit(){
	if(myForm.SQLstr.value!=''){
		UrlStr="StopBillPrints_HuaLien_Top.asp";
		myForm.action=UrlStr;
		myForm.target="topFrame";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
  }
</script>
<form method="POST" action="" target="" name="myForm">
	<input type="Hidden" name="SQLstr" value="<%=request("SQLstr")%>">
</form>

<frameset rows="100%,0%" cols="*" framespacing="0" frameborder="NO" border="0" onload="DataSubmit();">
  <frame src="" name="topFrame" id="topFrame" title="topFrame" />
  <frame src="" name="mainFrame" id="mainFrame" title="mainFrame" />
<noframes>
<body>
</body>
</noframes>

</frameset>
</html>
