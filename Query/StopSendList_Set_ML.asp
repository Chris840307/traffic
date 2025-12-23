<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<title>苗栗監理站攔停移送清冊</title>
<% Server.ScriptTimeout = 800 %>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBatchNumber=trim(request("BatchNumber"))


	
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="myForm" method="post">  
<table width='100%' border='1' align="left" cellpadding="1">
	<tr height="33">
		<td colspan="2" bgcolor="#FFCC66">苗栗監理站攔停移送清冊
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFCC">
			建檔日期
		</td>
		<td>
			<input type="text" name="RecordDate1" value="" maxlength="7" onKeyup="value=value.replace(/[^\d]/g,'')"> ~
			<input type="text" name="RecordDate2" value="" maxlength="7" onKeyup="value=value.replace(/[^\d]/g,'')">
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFCC33" align="center">
			<input type="Button" name="CreateListX" value="產生移送清冊" onclick="CreateList();">
		</td>
	</tr>
</table>		
</form>
<%
conn.close
set conn=nothing
%>
</body>
<script type="text/javascript" src="../js/date.js"></script>

<script language="JavaScript">

function CreateList(){
	var error=0;
		var errorString="";
		if(myForm.RecordDate1.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入建檔日期!!";
		}else if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate2.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入建檔日期!!";
		}else if(myForm.RecordDate2.value!=""){
			if(!dateCheck(myForm.RecordDate2.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}

		if (error>0){
			alert(errorString);
		}else{
			window.open("StopSendList_Excel_ML.asp?RecordDate1="+myForm.RecordDate1.value+"&RecordDate2="+myForm.RecordDate2.value,"inputWin123","width=900,height=500,left=50,top=0,scrollbars=yes,menubar=yes,resizable=yes,status=yes,toolbar=yes");
		}
}
</script>
</html>
