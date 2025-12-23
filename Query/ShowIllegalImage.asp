<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違規影像放大</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 

<%


%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">
<strong>放大：在圖片上按滑鼠左鍵 &nbsp; &nbsp; &nbsp; 縮小：按住Ctrl ＋ 在圖片上按滑鼠左鍵</strong>
<br>
	<img src="<%=trim(request("FileName"))%>" name="imgB3" onclick="fn_image(1);" >
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var image = document.getElementById("imgB3");
var PercentSize = (image.width / 10);
var NowPercent=2;
if (PercentSize > 100) {
	image.width=(PercentSize * 6);
}else{
	image.width=(PercentSize * 8);
}
function fn_image(type){
	if (event.ctrlKey == false){
		if (NowPercent<=10){
			image.width += PercentSize;
			NowPercent=NowPercent+1;
			//image.height +=50;
		}
	}else{
		if (NowPercent>=0)
		{
			image.width -= PercentSize;
			NowPercent=NowPercent-1;
			//image.height -=50;
		}
	}
}

</script>
