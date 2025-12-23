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
<title>舉發單刪除</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

if trim(request("kinds"))="DB_BillDel" then
	DelMemID=trim(Session("User_ID"))
	theBillSN=trim(request("DBillSN"))

	'抓單號
	theBillNO=""
	theCarNO=""
	strbillno="select BillNo,CarNo from BillBase where SN="&theBillSN
	set rsBillno=conn.execute(strbillno)
	if not rsBillno.eof then
		theBillNO=trim(rsBillno("BillNo"))
		theCarNO=trim(rsBillno("CarNo"))
	end if
	rsBillno.close
	set rsBillno=nothing

		NoteTmp=""
			
		'該筆紀錄的打驗資料表的 BILLSTATUS 更新為 6
		strUpdDelTemp="Update BillBaseTmp set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where BillNo='"&theBillNO&"'"
		conn.execute strUpdDelTemp

		'更新該筆紀錄的 BILLSTATUS 更新為 6
		strUpdDel="Update BillBase set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&theBillSN
		conn.execute strUpdDel

		DeleteReason="無"
		ConnExecute "舉發單刪除 單號:"&theBillNO&" 車號:"&theCarNO&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
%>
<script language="JavaScript">
	opener.myForm.submit();
	alert("刪除完成");
	window.close();
</script>
<%
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>舉發單刪除</strong></td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="15%" align="center" height="30">是否確定要刪除此筆舉發單?</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(237,4)=false then
'							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
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
function BillDel(){
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
}
</script>
</html>
