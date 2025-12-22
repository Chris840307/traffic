<% @EnableSessionState=False%>
<html xmlns:v>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>文件上傳進度指示條</title>
<STYLE>
v\:*{behavior:url(#default#VML);}
*{font-size:12px;}
</STYLE>
<style type="text/css">
<!--
font {
	font-size: 14px;
}
td {
	font-size: 14px;
	color: #333333;
}
b {
	font-size: 14px;
}
span {
	font-size: 14px;
}
a:link {
	color: #333333;
	text-decoration: none;
}
a:hover {
	color: #990000;
	text-decoration: underline;
}
a:visited {
	color: #000000;
	text-decoration: none;
}
-->
</style>
</head>
<BODY topmargin="0" leftmargin="0" onLoad="begin()" bgcolor="#CCCCCC"> 
<p><br> 
<table width="100%" border="0" cellspacing="0" cellpadding="4"> 
	<tr> 
		<td align="center"><b>文件上傳進度指示條</b></td> 
	</tr> 

	<tr> 
		<td>狀態：<span ID="myStatus"></span></td> 
	</tr> 

	<tr> 
		<td width="500"><div style="table-Layout:fixed;width:100%;height:100%;border:1 solid black"><v:RoundRect id="myRect" style="height:20;" name="myRect">  <v:fill type="gradient" id="fill1" color="blue"/> </v:RoundRect></div></td>
	</tr> 

	<tr> 
		<td>已經上傳：<span ID="message"></span></td> 
	</tr> 

	<tr> 
		<td>使用時間：<span ID="time">0</span> 秒 </td> 
	</tr> 

	<tr> 
		<td>平均速度：<span ID="speed">0</span> KB/秒 </td> 
	</tr> 
</table> 
</body>
</html>
<script language="Javascript"> 
	self.moveTo(getTop(200),getLeft(500)); 
	var intBytesTransferred=0; 
	var intTotalBytes=0; 
	var useTime=1; //s 
	var getData; 
	var myWidth=486; 
	var beginUploadFlg=false; 
	fill1.color="rgb("+Math.round(Math.random()*255)+","+Math.round(Math.random()*255)+","+Math.round(Math.random()*255)+")"; 
	myStatus.innerHTML="正在初始化...."; 

	function begin() 
	{ 
		message.innerHTML="開始讀取資料...."; 
		var Doc = new ActiveXObject('Microsoft.XMLDOM'); 
		Doc.async = false; 
		Doc.load("fileUpProgressRead.asp?progressID=<%=Request.QueryString("progressID")%>&aa="+new Date().getTime()); 

		if(Doc.parseError.errorCode != 0) //檢查獲取數據時是否發生錯誤
		{ 
			delete(Doc); 
			if(beginUploadFlg){ 
				intBytesTransferred=intTotalBytes; 
			}else{ 
				message.innerHTML="上傳動作尚未啟動！"; 
			} 
		}else{ 
			var rootNode=Doc.documentElement; 
			if(rootNode.childNodes != null)  
			{
				beginUploadFlg=true; 
				intBytesTransferred=Number(rootNode.childNodes.item(0).childNodes.item(0).text); 
				intTotalBytes=Number(rootNode.childNodes.item(0).childNodes.item(1).text); 
				useTime=Number(rootNode.childNodes.item(0).childNodes.item(2).text); 
				message.innerHTML="讀取訊息成功。"; 
			} 
			delete(rootNode); 
		} 

		delete(Doc); 

		if(intTotalBytes==0){ 
			intBytesTransferred=1; 
			intTotalBytes=100; 
		} 

		display(); 

		if(intTotalBytes>0 && intBytesTransferred<intTotalBytes){ 
			if(beginUploadFlg){ 
				myStatus.innerHTML="正在上傳，請耐心等待...."; 
			} 

			time.innerHTML=useTime; 
			speed.innerHTML=Math.round((intBytesTransferred/useTime)/1024); 
			getData = setTimeout("begin()",1000); 
		}else{ 
			myStatus.innerHTML="數據上傳完畢，3秒後自動關閉。"; 
			setTimeout("self.close()",3000); 
		} 
	} 

	function display(){ 
		myRect.style.width=Math.round(myWidth/(intTotalBytes/intBytesTransferred)); 
		fill1.angle=Math.round(300/(intTotalBytes/intBytesTransferred)); 

		if(beginUploadFlg){ 
			message.innerText=intBytesTransferred+"/"+intTotalBytes+","+Math.round(100/(intTotalBytes/intBytesTransferred))+"%"; 
		} 
	} 
	function getTop(windowHeight){ 
		var top = parseInt((screen.height - windowHeight)/2-15); 
		return top; 
	} 

	function getLeft(windowWidth){ 
		var left = parseInt((screen.width - windowWidth)/2-5); 
		return left; 
	} 
</script>