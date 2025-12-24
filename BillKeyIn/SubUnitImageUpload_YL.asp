<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/BannernoData.asp"-->
	<!--#include File="../Common/css.txt"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:10pt}
A:link {text-decoration : none;color=0044ff;line-height:16px;font-size:10pt}
A:visited {text-decoration : none;color=0044ff;line-height:16px;font-size:10pt}
A:hover {text-decoration : underline;color=ff6600;line-height:16px;font-size:10pt}
td {font-family:新細明體;line-height:16px;font-size:10pt}
input {font-family:新細明體;line-height:16px;font-size:10pt}
select {font-family:新細明體;line-height:16px;font-size:10pt}
-->
</style>
<%
'if trim(request("kinds"))="img_Upload" then
	TodayFolder=year(now)-1911&right("00"&month(now),2)&right("00"&day(now),2)
	'Ftp連結位置
	FtpLocation=""
	strftp="select Value from ApConfigure where ID=37"
	set rsftp=conn.execute(strftp)
	if not rsftp.eof then
		FtpLocation=trim(rsftp("Value"))
		FtpLocationIE8=trim(rsftp("Value"))
	end if
	rsftp.close
	set rsftp=nothing

	'影像存放位置
	FileLocation=""
	strfile="select Value from ApConfigure where ID=110"
	set rsFile=conn.execute(strfile)
	if not rsFile.eof then
		FileLocation=trim(rsFile("Value"))
	end if
	rsFile.close
	set rsFile=nothing
	FtpLocation=FtpLocation & "Type8/Type8/" & trim(Session("Credit_ID")) & "/" & (year(now)-1911)&Right("00"&month(now),2)&Right("00"&day(now),2) & "/"
	FtpLocationIE8=FtpLocationIE8 & "Type8/Type8/"&trim(Session("Credit_ID")) & "/" & (year(now)-1911)&Right("00"&month(now),2)&Right("00"&day(now),2) & "/"
	FileLocation=FileLocation & "Type8\" 

	'response.write FtpLocation
	'response.write "<br>" & FileLocation

	dim fso1 
	set fso1=Server.CreateObject("Scripting.FileSystemObject")
	if (fso1.FolderExists(FileLocation&Session("Credit_ID")))=false then
		fso1.CreateFolder FileLocation&Session("Credit_ID")
	end if
	if (fso1.FolderExists(FileLocation&Session("Credit_ID")&"\"&TodayFolder))=false then
		fso1.CreateFolder FileLocation&Session("Credit_ID")&"\"&TodayFolder
	end if
	set fso1=nothing

	'userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	'If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 
	'if trim(userip) = "10.136.10.173" then
	'	FtpLocationIE8=FtpLocation
	'end if 
'end if
%>
<script language="JavaScript">
	//alert ("請將欲上傳之影像檔拖曳至FTP視窗中!!");
	for(i=0;i<10;i++)	//ie version
	{
	if(navigator.appVersion.match("MSIE "+i)!=null)
	x=i;
	}
	if (x>=7){
		alert ("請點選檢視中的在Windows檔案總管中開啟FTP\n(不需理會Internet Explorer 無法顯示網頁的錯誤訊息)\n，再將欲上傳之影像檔拖曳至FTP視窗中!!\n上傳相片前，請先將相片縮小到適當大小(請勿超過800KB)，檔案過大會造成讀取問題!!");
		window.open("<%=FtpLocationIE8%>","FtpWin135","location=0,width=770,height=455,resizable=yes,scrollbars=yes,menubar=yes");

	}else{
		alert ("請將欲上傳之影像檔拖曳至FTP視窗中!!\n上傳相片前，請先將相片縮小到適當大小(請勿超過800KB)，檔案過大會造成讀取問題!!");
		window.open("<%=FtpLocation%>","FtpWin135","location=0,width=770,height=455,resizable=yes,scrollbars=yes,menubar=yes");

	}
	
	window.close();
</script>   
<title>違規數位影像上傳</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	
</body>
</html>
