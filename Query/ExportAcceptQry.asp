<!doctype html>
<html lang="en">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>Document</title>
</head>
<body>
	<form name="myForm" method="post">
		<table width="100%" border="0">
			<tr>
				<td bgcolor="#CCCCCC">
					<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
						<tr>
							<td nowrap>
								建檔日期
								<input name="RecordDate1" type="text" value="<%=request("RecordDate1")%>" size="11" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
								~
								<input name="RecordDate2" type="text" value="<%=request("RecordDate2")%>" size="10" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate2');">

								<input type="submit" name="btnSelt" value="攔停案件查詢" onclick="funSelt();">
								<input type="submit" name="btnSelt" value="逕舉案件查詢" onclick="funrunSelt();">
								<input type="submit" name="btnSelt" value="更新檔上傳" onclick="funUpdate();">
								<input type="button" name="cancel" value="清除" onClick="location='ExportAcceptQry.asp'">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(DBKind){
	var error=0;

	if(myForm.RecordDate1.value==""||myForm.RecordDate2.value==""){
		error=1;
		alert("建檔日期輸入不正確!!");
	}

	if(myForm.RecordDate1.value!=""){
		if(!dateCheck(myForm.RecordDate1.value)){
			error=1;
			alert("建檔日期輸入不正確!!");
		}
	}

	if(myForm.RecordDate2.value!=""){
		if(!dateCheck(myForm.RecordDate2.value)){
			error=1;
			alert("建檔日期輸入不正確!!");
		}
	}
	if(error==0){
		UrlStr="ExportAccept.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funrunSelt(DBKind){
	var error=0;

	if(myForm.RecordDate1.value==""||myForm.RecordDate2.value==""){
		error=1;
		alert("建檔日期輸入不正確!!");
	}

	if(myForm.RecordDate1.value!=""){
		if(!dateCheck(myForm.RecordDate1.value)){
			error=1;
			alert("建檔日期輸入不正確!!");
		}
	}

	if(myForm.RecordDate2.value!=""){
		if(!dateCheck(myForm.RecordDate2.value)){
			error=1;
			alert("建檔日期輸入不正確!!");
		}
	}
	if(error==0){
		UrlStr="ExportRunAccept.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funUpdate(DBKind){

	UrlStr="UpdateStyle.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
</script>
