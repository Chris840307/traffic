<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">
<script language="JavaScript">
	window.focus();
</script>
<head>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<style type="text/css">
<!--
.style4 {
	font-size: 16px
}
.style5 {
	font-size: 20px;
	color: #FF0000;
}
-->
</style>
<title></title>
<% Server.ScriptTimeout = 3800 %>
<%
'tmpSQL=Session("BillSQLforReport")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="BillCloseUpdate" then
	str1="update billbase set billstatus='3' where billstatus='9' " &_
		" and sn in (select billsn from dcilog where batchnumber='"&Trim(request("BatchNumber"))&"')"
	conn.execute str1
%>
<script language="JavaScript">
	alert("處理完成！");
</script>
<%
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td>撤銷送達回傳後資料處理
				</td>
			</tr>
			<tr>
				<td>
				
					撤銷送達作業批號<input type="text" name="BatchNumber" value="" size="16" >

				</td>
			</tr>


			<tr>
				<td bgcolor="#EBFBE3" align="center">
					<input type="button" value="確 定" name="b1" onclick="BillCloseUpdate();">
					<input type="hidden" value="" name="kinds">
				</td>
			</tr>
			
		</table>

	</form>
<%
conn.close
set conn=nothing
%>
</body>

<script language="JavaScript">
function BillCloseUpdate(){
	if (myForm.BatchNumber.value==""){
		alert("請輸入作業批號!");
	}else{
		myForm.kinds.value="BillCloseUpdate";
		myForm.submit();
	}	
}

</script>
</html>
