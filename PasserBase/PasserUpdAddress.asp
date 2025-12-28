<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人道路障礙整批戶籍地址更新</title>
<style type="text/css">
<!--
.style1 {font-family: "新細明體"; font-size: 12px;}
-->
</style>
</head>
<body>
<%
if trim(request("kinds"))="Update" then
	if trim(request("Upd_CreditID"))<>"" and trim(request("Upd_ChName"))<>"" then
		Pcnt1=0
		strQ="Select Count(*) as cnt from PasserBase " &_
		" where DriverID='"&trim(request("Upd_CreditID"))&"' and Driver='"&trim(request("Upd_ChName"))&"'" &_
		" and RecordStateID=0"
		set rsQ=conn.execute(strQ)
			Pcnt1=Cdbl(rsQ("cnt"))
		rsQ.close
		set rsQ=nothing

		if Pcnt1>0 then
			strUpd1="Update PasserBase set Note=Note || '原地址：' ||DriverAddress " &_
			" where DriverID='"&trim(request("Upd_CreditID"))&"' and Driver='"&trim(request("Upd_ChName"))&"'" &_
			" and RecordStateID=0"
			conn.execute strUpd1

			strUpd2="Update PasserBase set DriverZip='"&trim(request("Upd_DriverZip"))&"',DriverAddress='"&trim(request("Upd_DriverAddress"))&"'" &_
			" where DriverID='"&trim(request("Upd_CreditID"))&"' and Driver='"&trim(request("Upd_ChName"))&"'" &_
			" and RecordStateID=0"
			conn.execute strUpd2
%>
	<script language="JavaScript">
		alert("修改完成！");
		window.close();
	</script>
<%		else
%>
	<script language="JavaScript">
		alert("查無違規人身份證號：<%=trim(request("Upd_CreditID"))%>，姓名：<%=trim(request("Upd_ChName"))%>之案件！");
	</script>
<%
		end if
	end if
end if
%>
<form name=myForm method="post">
<table width="645" border="1">
	<tr bgcolor="#FFCC66">
		<td colspan="4">慢車行人道路障礙整批戶籍地址更新</td>
	</tr>
	<tr>
		<td width="20%" bgcolor="#FFFFCC">違規人身份證號</td>
		<td width="30%"><input type="text" name="Upd_CreditID" value="" onkeyup="this.value=this.value.toUpperCase()" size="18"></td>
		<td width="20%" bgcolor="#FFFFCC">違規人姓名</td>
		<td width="30%"><input type="text" name="Upd_ChName" value="" size="18"></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFCC" >戶籍地址</td>
		<td colspan="3">
			<input type="text" name="Upd_DriverZip" value="" size="10">郵遞區號
			<input type="text" name="Upd_DriverAddress" value="" size="53">
		</td>
	</tr>
	<tr>
		<td colspan="4" >＊此功能會修正行人慢車道路障礙歷史資料庫中，與使用者輸入的身份證號以及姓名相同之舉發單的違規地址。舊的違規地址會改寫入該案件的備註中。</td>
	</tr>
	<tr bgcolor="#FFCC66">
		<td colspan="4" align="center">
			<input type="button" value="確定更新" onclick="funUpdate();">
			<input type="button" value="離開" onclick="window.close();">
			<input type="hidden" value="" name="kinds">
		</td>
	</tr>
</table>
</form>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funUpdate(){
	var error=0;
	var errorString="";
	if (myForm.Upd_CreditID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入身份證號。";
	}
	if (myForm.Upd_ChName.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入姓名。";
	}
	if (myForm.Upd_DriverZip.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入郵遞區號";
	}
	if (myForm.Upd_DriverAddress.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入戶籍地址";
	}
	if (error==0){
		myForm.kinds.value="Update";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
</script>
</html>
<%
conn.close
set conn=nothing
%>