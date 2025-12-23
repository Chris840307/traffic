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
<form name=myForm method="post">
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
						<input name="chkDate" type="radio" value="0" checked>填單日
						<input name="chkDate" type="radio" value="1">違規日
						<input name="chkDate" type="radio" value="2">建檔日<br>
						<input type='text' size='9' name='startDate_q' value='' maxLength='7'>
						<input name="datestra" type="button" value="..." onclick="OpenWindow('startDate_q');">
						~
						<input type='text' size='9' name='endDate_q' value='' maxLength='7'>
						<input name="datestrb" type="button" value="..." onclick="OpenWindow('endDate_q');">
						
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value="產生第一次未退回統計表 " onclick="funAdd();"> &nbsp; &nbsp; &nbsp;
			<input name="btnadd" type="button" value="產生第二次未退回統計表 " onclick="funAdd2();"> 
		</td>
	</tr>
</table>
	
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funAdd(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;

	if (((sDate=="") || (eDate==""))){
		alert("請輸入統計期間!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{
			UrlStr="MailNotBackReport_1_Excel.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBackReport_1_Excel";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

function funAdd2(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;

	if (((sDate=="") || (eDate==""))){
		alert("請輸入統計期間!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else{
			UrlStr="MailNotBackReport_2_Excel.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBackReport_2_Excel";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}
</script>