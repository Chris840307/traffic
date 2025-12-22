<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>罕見字資料新增</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<%
If trim(request("DB_Add"))="Add" Then
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	Sys_TD_ADDRESS="":Sys_TD_OwnerName=""

	If not ifnull(Request("Sys_TD_ADDRESS")) Then
		Sys_TD_ADDRESS=replace(trim(Request("Sys_TD_ADDRESS")),"臺","台")

		Sys_TD_ADDRESS=trim(replace(request("Sys_TD_ADDRESS"),"＿","_"))

	End if

	If not ifnull(Request("Sys_TD_OwnerName")) Then
		Sys_TD_OwnerName=trim(replace(request("Sys_TD_OwnerName"),"＿","_"))
	End if
	

	strSQL="Insert Into TDDT_RAREWORD values((select NVL(MAX(TO_Number(TD_SN)),0)+1 from TDDT_RAREWORD),'"&Ucase(trim(replace(request("Sys_TD_CARNO"),"－","-")))&"','"&Sys_TD_OwnerName&"',sysdate,'"&trim(Session("User_ID"))&"','"&trim(Sys_City)&"','0','"&Sys_TD_ADDRESS&"')"

	conn.execute(strSQL)

	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
End if
%>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">罕見字資料新增</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>車號</td>
					<td>
						<input name="Sys_TD_CARNO" class="btn1" type="text" value="" size="13" maxlength="7" onkeyup="funChkCarNo();">
						<br><span id="ChkCarNo"></span>
					</td>
					<td bgcolor="#FFFF99">車主姓名(缺漏字請用『_』替代 )</td>
					<td> 
						<input name="Sys_TD_OwnerName" class="btn1" type="text" value="" size="12" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">車主地址(缺漏字請用『_』替代 )</td>
					<td colspan="3"> 
						<input name="Sys_TD_ADDRESS" class="btn1" type="text" value="" size="50" maxlength="100">
					</td>
				</tr>
		  </table>
		</td>
	</tr>
	<tr bgcolor="#ffffff" align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="save" value=" 儲 存 " onclick="funAdd();">
			<input type="button" name="exit" value=" 離 開 " onclick="funExt();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
<input type="Hidden" name="chk_carno" value="">
</form>
</body>
</html>
<script type="text/javascript" src="/traffic/js/date.js"></script>
<script language="javascript">
function funAdd(){
	var err=0;errmsg='';
	if (myForm.Sys_TD_CARNO.value==''){
		err=1;
		errmsg="車號不可空白!!\n";
	}else if(myForm.Sys_TD_CARNO.value.replace("－","-").search("-") <= 0){
		err=1;
		errmsg=errmsg+"車號格式不正確!!\n";
	}

	if (myForm.Sys_TD_OwnerName.value=='' && myForm.Sys_TD_ADDRESS.value==''){
		err=1;
		errmsg=errmsg+"車主姓名或地址不可空白!!\n";
	}else if(myForm.Sys_TD_OwnerName.value!='' && myForm.Sys_TD_OwnerName.value.replace("＿","_").search("_") <= 0){
		err=1;
		errmsg=errmsg+"車主姓名沒有缺漏字!!\n";

	}else if(myForm.Sys_TD_ADDRESS.value!='' && myForm.Sys_TD_ADDRESS.value.replace("＿","_").search("_") <= 0){
		err=1;
		errmsg=errmsg+"車主地址沒有缺漏字!!\n";
	}

	if (myForm.chk_carno.value=="1"){
		err=1;
		errmsg=errmsg+"該車號已有資料!!\n";
	}

	if (err==0){
		myForm.Sys_TD_OwnerName.value=myForm.Sys_TD_OwnerName.value.replace("＿","_");
		myForm.Sys_TD_CARNO.value=myForm.Sys_TD_CARNO.value.replace("－","-");
		myForm.Sys_TD_CARNO.value=myForm.Sys_TD_CARNO.value.toUpperCase();
		myForm.DB_Add.value="Add";
		myForm.submit();
	}else{
		alert(errmsg);
	}
}
function funChkCarNo() {
	if(myForm.Sys_TD_CARNO.value.replace("－","-").search("-") > 0){
		myForm.Sys_TD_CARNO.value=myForm.Sys_TD_CARNO.value.replace("－","-");
		myForm.Sys_TD_CARNO.value=myForm.Sys_TD_CARNO.value.toUpperCase();
		runServerScript("RareWord_Chk.asp?Sys_TD_CARNO="+myForm.Sys_TD_CARNO.value+"&Sys_TD_OwnerName="+myForm.Sys_TD_OwnerName.value+"&Sys_TD_ADDRESS="+myForm.Sys_TD_ADDRESS.value);
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		opener.myForm.submit();
		self.close();
	}
}
myForm.Sys_TD_CARNO.focus();
</script>
<%conn.close%>