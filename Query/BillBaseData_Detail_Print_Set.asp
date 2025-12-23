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
<title>舉發單詳細列印</title>
<% Server.ScriptTimeout = 800 %>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBatchNumber=trim(request("BatchNumber"))

if trim(request("kinds"))="DB_BillDel" then
		
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>舉發單詳細列印</strong></td>
			</tr>
			<tr>
				<td>
					<input type="checkbox" name="s_Detail" value="1" checked>舉發單資料
					<br>
					<input type="checkbox" name="s_Mail" value="1" >舉發單郵寄歷程記錄
					<br>
					<input type="checkbox" name="s_Image" value="1" checked>違規影像
					<br>
					<input type="checkbox" name="s_Send" value="1" checked>送達証書
					<br>
					<input type="checkbox" name="s_Gov" value="1" checked>公告公文
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 列 印 " onclick="BillPrint();" >
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="BillSnTmp" value="<%=trim(request("BillSnTmp"))%>">
		
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
<script type="text/javascript" src="../js/form.js"></script>
<script language="JavaScript">
function BillPrint(){
	var s_Detail=0;
	var s_Mail=0;
	var s_Image=0;
	var s_Send=0;
	var s_Gov=0;
	if (myForm.s_Detail.checked==true){
		s_Detail=1;
	}
	if (myForm.s_Mail.checked==true){
		s_Mail=1;
	}
	if (myForm.s_Image.checked==true){
		s_Image=1;
	}
	if (myForm.s_Send.checked==true){
		s_Send=1;
	}
	if (myForm.s_Gov.checked==true){
		s_Gov=1;
	}
	window.open("BillBaseData_Detail_Print.asp?s_Detail="+s_Detail+"&s_Mail="+s_Mail+"&s_Image="+s_Image+"&s_Send="+s_Send+"&s_Gov="+s_Gov+"&BillSnTmp="+myForm.BillSnTmp.value,"UploadFile_Print","left=0,top=0,location=0,width=1010,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes")
}



</script>
</html>
