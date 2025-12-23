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
<title>車牌號碼修改</title>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBillSN=trim(request("DBillSN"))
'theDelType=trim(request("DelType"))	'單筆或多筆刪除
if trim(request("kinds"))="CarNoUpdate" then
	strUpd="Update BillBase set CarNo='"&trim(request("CarNo"))&"',BillStatus='0' where SN="&trim(request("DBillSN"))
	conn.execute strUpd
%>
<script language="JavaScript">
	opener.myForm.submit();
	window.close();
</script>
<%
end if
	strCarNo="select CarNo from BillBase where SN="&trim(request("DBillSN"))
	set rsCarNo=conn.execute(strCarNo)
	if not rsCarNo.eof then
%>

<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>車牌號碼修改</strong></td>
			</tr>
			<tr>
				<td width="25%" align="right" bgcolor="#EBFBE3">車牌號碼</td>
				<td width="75%">
				<input type="text" name="CarNo" value="<%=trim(rsCarNo("CarNo"))%>" onblur="getVIPCar()">
				<input type="hidden" name="OldCarNo" value="<%=trim(rsCarNo("CarNo"))%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(234,4)=false then
'							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
			<br>
			<span class="style1">(車牌號碼修改後，請重新進行『上傳監理站-車籍查詢』)</span>
				</td>
			</tr>
		
		</table>		

	</form>
<%
	end if
	rsCarNo.close
	set rsCarNo=nothing
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function BillDel(){
	if (myForm.CarNo.value == myForm.OldCarNo.value){
		window.close();
	}else if (myForm.CarNo.value==""){
		alert("請輸入車牌號碼!");
	}else{
		myForm.kinds.value="CarNoUpdate";
		myForm.submit();
	}
}
function getVIPCar(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			alert("車牌格式錯誤");
			myForm.CarNo.select();
		}
	}
}
myForm.CarNo.select();
</script>
</html>
