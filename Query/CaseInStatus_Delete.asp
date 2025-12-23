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
<style type="text/css">
<!--
.style1 {
font-size: 24px; 
line-height:28px;
}
.style2 {
font-size: 20px; 
line-height:26px;
}
-->
</style>
<title>強制刪除</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("kinds"))="DB_Update" then
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

	'寫入LOG
	ConnExecute "強制刪除:"&theBillSN&","&theBillNO,352


	strUpdDel="Update DciLog set DciReturnStatusID='S' " &_
		" where billSN="&theBillSN&" and ExchangeTypeID='E'"
	conn.execute strUpdDel

	strUpdDel="Update BillBase set RecordstateID=-1,BillStatus=6 " &_
		" where SN="&theBillSN
	conn.execute strUpdDel

%>
<script language="JavaScript">
	opener.myForm.submit();
	alert("強制<%if sys_City="高雄市" then%>註銷<%else%>刪除<%end if%>完成!!");
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
				<td colspan="4"><strong>舉發單強制<%if sys_City="高雄市" then%>註銷<%else%>刪除<%end if%></strong></td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="15%" align="center" height="30">
				
				<span class="style2">是否要確定要<strong><%if sys_City="高雄市" then%>註銷<%else%>刪除<%end if%></strong>此舉發單？</span>
				<br>
				<span class="style1"><strong><font color="red">(本功能只有更改縣市端資料庫，不會上傳至監理站，請務必確認監理站無此單號後才執行)</font></strong></span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillUpdate();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(237,4)=false then
'							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("BillSN"))%>">
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
function BillUpdate(){
		myForm.kinds.value="DB_Update";
		myForm.submit();
}
</script>
</html>
