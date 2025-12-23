<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<script language=javascript src='../js/form.js'></script>
<script language=javascript>
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing


%>

</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>查詢郵寄未退回清冊</title>
</head>
<body >
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>查詢未退還清冊</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="1" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
									
					<td nowrap valign="top" bgcolor="#FFFFFF">
		
						<b>統計期間</b>
						<input name="chkDate" type="radio" value="2" checked>建檔日<br>
						<input type='text' size='9' name='startDate_q' value='' maxLength='8'>
						<input name="datestra" type="button" value="..." onclick="OpenWindow('startDate_q');">
						~
						<input type='text' size='9' name='endDate_q' value='' maxLength='8'>
						<input name="datestrb" type="button" value="..." onclick="OpenWindow('endDate_q');">
					</td>
					<td nowrap bgcolor="#FFFFFF">

						<b>批號</b>						<input name="batchnumber" type="text" value="" onkeyup="this.value=this.value.toUpperCase()">		
						<br>

					</td>

				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value="產生清冊 " onclick="funAdd();"> 

			<input name="btnexit" type="button" value=" 產生郵局查詢單(A4) " onclick="funMailQry();">

			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
		</td>
	</tr>
</table>
	<input type="hidden" name="unitSelectlist">
	<input type="hidden" name="MemSelectlist">
</form>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funExt() {
	window.close();
}


function funAdd(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;



	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="")){
		alert("入案批號、統計期間請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{
			UrlStr="StopMailNotBakList.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funMailQry(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;



	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="")){
		alert("入案批號、統計期間請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{
			UrlStr="FAXQryMail_HulienR.asp";
			myForm.action=UrlStr;			
			myForm.target="FAXQryMail_HulienR";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

</script>
</body>
</html>
