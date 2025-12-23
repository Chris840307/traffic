<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="traffic/Common/css.txt"-->
<title>催繳郵寄未退回清冊</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">
		<br><br>
		<table width='100%' border='0' bgcolor="#CCCCCC" cellpadding="1" cellspacing="2">
			<tr>
				<td colspan="2" height="27" bgcolor="#FFCC33" class="pagetitle">
					<strong>催繳郵寄未退回</strong>
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td nowrap>上傳批號</td>

				<td nowrap>
					<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=request("Sys_BatchNumber")%>" size="8">
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td nowrap>匯入日期</td>

				<td nowrap>
					<input name="Sys_RecordDate1" type="text" class="btn1" value="<%=request("Sys_RecordDate1")%>" size="8" maxlength="7">
					<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_RecordDate1');">
					　～　
					<input name="Sys_RecordDate2" type="text" class="btn1" value="<%=request("Sys_RecordDate2")%>" size="8" maxlength="7">
					<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_RecordDate2');">
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#EBFBE3" align="center">
					<input type="button" value="產生清冊" onclick="funMailListCity();">
					<input type="button" value="匯出檔案" onclick="funMailListCityTxt();">
					<input type="button" value="關  閉" onclick="if(confirm('是否關閉維護系統?')){self.close();}">
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

function funMailListCity(){
	if ((myForm.Sys_RecordDate1.value=='' || myForm.Sys_RecordDate2.value=='')&&(myForm.Sys_BatchNumber.value=='')){

		alert("請輸入匯入日期或批號!");
	}else{


		UrlStr="StopBillReturnList.asp";
		myForm.action=UrlStr;
		myForm.target="BillReturnList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		self.close();
	}
}

function funMailListCityTxt(){
	if ((myForm.Sys_RecordDate1.value=='' || myForm.Sys_RecordDate2.value=='')&&(myForm.Sys_BatchNumber.value=='')){

		alert("請輸入匯入日期或批號!");
	}else{


		UrlStr="StopBillReturnList_txt.asp";
		myForm.action=UrlStr;
		myForm.target="BillReturnList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		self.close();
	}
}
</script>
</html>
