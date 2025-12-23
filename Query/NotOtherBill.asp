<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>不須另案舉發</title>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%
Server.ScriptTimeout = 6800
Response.flush
'權限
'AuthorityCheck(234)
RecordDate=split(gInitDT(date),"-")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

'組成查詢SQL字串
if request("DB_Selt")="Update" then
	strUpd="Insert into OtherBill values("&trim(request("ReCoverSn"))&",null,sysdate,"&trim(Session("User_ID"))&",'-1')"
	conn.execute strUpd
%>
	<script language="JavaScript">
		alert("作業完成!!");
		opener.myForm.submit();
		window.close();
	</script>
<%
end if


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
	line-height:28px;
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body bgcolor="#EBFBE3">
<form name=myForm method="post">
	<center>
	<span class="style6">是否確定此筆舉發單(<%=trim(request("ReCoverBillNo"))%>)<font color="red">不須</font>做另案舉發?</span>
	</center>
	<br>	
	<center>
	<input type="button" value="確定" onclick="NotOtherBill()" style="font-size: 12pt; width: 90px; height:28px;">&nbsp; &nbsp; 
	<input type="button" value="取消" onclick="window.close();" style="font-size: 12pt; width: 90px; height:28px;">
	</center>
	<input type="hidden" value="" name="DB_Selt">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function NotOtherBill(){
	myForm.DB_Selt.value="Update";
	myForm.submit();
}

</script>
<%
conn.close
set conn=nothing
%>