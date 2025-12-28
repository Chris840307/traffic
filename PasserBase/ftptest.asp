<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>慢車掃描檔上傳</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.style2 {font-size: 10px; }
-->
</style>
</HEAD>
<%
Server.ScriptTimeout=600000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

'scannerDir=Session("Credit_ID")
'
'fp=server.mappath(".\PBFtp\PasserFtp")&"\"&scannerDir
'
'set fso=Server.CreateObject("Scripting.FileSystemObject")
'
'if fso.FolderExists(fp) = false then
'	fso.CreateFolder(fp) 
'end if
'
'ServerIp=Request.ServerVariables ("LOCAL_ADDR") 
'
'ftpUrl="ftp://passer:pbuser@"&ServerIp&"/PasserFtp/"&scannerDir
'
'set fso=nothing
%>
<BODY>
<form name="myForm" method="post">

</form>
	<input type="button" name="btnAdd" value="新增" onclick="RunExe();">
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	function RunExe(){
		var dt = new Date();
		runServerScript("PasserBaseFtpOpen.asp?nowtime="+dt);
	}
</script>
