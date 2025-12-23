<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<%
		strSQL="select loginid from memberdata where memberid='"&session("User_ID")&"' and recordstateid=0 and accountstateid=0"
	set rsfound=conn.execute(strSQL)
	If Not rsfound.eof Then 
		loginid=rsfound("loginid")
	Else
		loginid=""
	End If

	FtpLocation="ftp://station:stationdci@10.133.2.178/"
	
%>
<script language=javascript src='../js/form.js'></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>查詢拖吊清冊</title>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>查詢拖吊清冊
		<% '鳳山拖吊隊因監理站未合併，故先不提供該功能，等監理站合併後約101年，再提供
		'If session("Unit_ID") <> "0872" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="拖吊已結上傳說明檔.doc" target="_blank">拖吊已結上傳說明檔</a>
		<%'End if%>
		</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<b>統計期間</b>

						<input name="chkDate" type="radio" value="2" checked>建檔日</td>
						<td><B>建檔人員臂章號碼</B></td><td><B>序號起迄(需要的時候才輸入)<b></td>
						<tr>
						<td>
						<input type='text' size='9' name='startDate_q' value='' maxLength='8'>
						<input name="datestra" type="button" value="..." onclick="OpenWindow('startDate_q');">
						~
						<input type='text' size='9' name='endDate_q' value='' maxLength='8'>
						<input name="datestrb" type="button" value="..." onclick="OpenWindow('endDate_q');">
						</td>
						<td>

						<input type='text' size='50' name='loginid' value='<%=loginid%>' maxLength='50' 
						onkeyup="value=value.toUpperCase()">
						</td>
						<td>
						<input type='text' size='9' name='SNStart' value='' maxLength='8' onkeyup="value=value.replace(/[^\d]/g,'')">
						~
						<input type='text' size='9' name='SNEnd' value='' maxLength='8' onkeyup="value=value.replace(/[^\d]/g,'')">						
					</td>

				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value="產生清冊 A3" onclick="funAdd2();"> 
			<img src="space.gif" width="20" height="5">						
			<input name="btnadd" type="button" value="產生清冊 A4" onclick="funAdd4();"> 
						
			<img src="space.gif" width="20" height="5">
			<input name="btnFile" type="button" value="匯出檔案 " onclick="funExportFile();"> 
			<img src="space.gif" width="20" height="5">
			<%'If session("Unit_ID") <> "0872" then%>
			<input type="button" value="上傳結案檔" name="UpLoadFile" onclick="UploadData();">
			<%'End if%>
		</td>
	</tr>
</table>
	<input type="hidden" name="unitSelectlist">
	<input type="hidden" name="MemSelectlist">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function UploadData()
{
  window.open("<%=FtpLocation%>","FtpWin","location=0,width=770,height=455,resizable=yes,scrollbars=yes,toolbar=yes");

}

function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funAdd2(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	var loginid = myForm.loginid.value;
	var SNStart = myForm.SNStart.value;
	var SNEnd = myForm.SNEnd.value;	

	if (sDate=="" || eDate=="" || loginid==""){
		alert("統計期間請輸入或建檔人員臂章號碼!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{

			UrlStr="StopSendList_TakeCar_Excel_A3.asp?StartDate="+sDate+"&EndDate="+eDate+"&loginid="+loginid+"&SNStart="+SNStart+"&SNEnd="+SNEnd;
			myForm.action=UrlStr;			
			myForm.target="StopSendList_TakeCar_Excel_A3";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	
}
}

function funAdd4(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	var loginid = myForm.loginid.value;
	var SNStart = myForm.SNStart.value;
	var SNEnd = myForm.SNEnd.value;	

	if (sDate=="" || eDate=="" || loginid==""){
		alert("統計期間請輸入或建檔人員臂章號碼!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{

			UrlStr="StopSendList_TakeCar_Excel_A4.asp?StartDate="+sDate+"&EndDate="+eDate+"&loginid="+loginid+"&SNStart="+SNStart+"&SNEnd="+SNEnd;
			myForm.action=UrlStr;			
			myForm.target="StopSendList_TakeCar_Excel_A4";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	
}
}

function funExportFile(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	var loginid = myForm.loginid.value;

	if (sDate=="" || eDate=="" || loginid==""){
		alert("統計期間請輸入或建檔人員臂章號碼!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{

			UrlStr="ExportTakeCar.asp?StartDate="+sDate+"&EndDate="+eDate+"&loginid="+loginid;
			myForm.action=UrlStr;			
			myForm.target="ExportTakeCar";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	
}
}



</script>