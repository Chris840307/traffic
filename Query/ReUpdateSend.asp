<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>重新上傳送達</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Trim(request("kinds"))="Upload" Then
	strUpd="Update dcilog set filename='',seqno='',dcireturnstatusid='' where Sn="&Trim(request("Sn"))
	conn.execute strUpd

	If sys_City="苗栗縣" Then
		strUpd="Update BillBase set CaseInDate2=Sysdate where sn in(select billsn from dcilog where Sn="&Trim(request("Sn"))&")"
		conn.execute strUpd
	end if

	If sys_City="嘉義市" then
		strIns="Insert into ReloadReason values("&Trim(request("Sn"))&",'"&Trim(request("ReUploadReason"))&"',"&Trim(session("User_ID"))&",sysdate)"
		conn.execute strIns
	end if
%>
	<script language="JavaScript">
		alert("重傳完成!");
		window.close();
	</script>	
<%
End If 

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
	<table width='100%' border='1' align="left" cellpadding="1">
		<tr bgcolor="#FFCC33">
			<td colspan="2">重新上傳送達
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<font color="#FF0000"><strong>如遇到監理站未收到送達註記才可使用。</strong></font>
			</td>
		</tr>
		<tr>
			<td>
				重新上傳原因
			</td>
			<td>
				<input type="text" name="ReUploadReason" value="<%=Trim(request("ReUploadReason"))%>" size="80">
			</td>
		</tr>

		<tr>
			<td bgcolor="#EBFBE3" align="center" colspan="2">
				<input type="button" value="重新上傳" name="b1" onclick="funUpload();">
				<input type="hidden" value="" name="kinds" >
			</td>
		</tr>
		<tr>
			<td bgcolor="#EBFBE3" align="center" id="LayerUp">
				
			</td>
		</td>
	</table>
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	function funUpload(){
		if (myForm.ReUploadReason.value==""){
			alert("請輸入重新上傳原因!");
		}else{
			myForm.kinds.value="Upload";
			myForm.submit();
		}
	}
</script>
