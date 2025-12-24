<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Banner.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/date.js"></script>
<%

%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='400' border='1' align="center" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="2"><strong>舉發單綜合查詢</strong></td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="30%" align="right">舉發單號</td>
				<td width="70%">
					<input type="text" name="BillNo" size="12" maxlength="9" value="" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">車牌號碼</td>
				<td>
					<input type="text" name="CarNo" size="12" value="" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規人身份證號</td>
				<td>
					<input type="text" name="illFID" size="12" value="" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規人姓名</td>
				<td>
					<input type="text" name="illName" size="12" value="">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發單類別</td>
				<td>
					<select name="BillType" onchange="ChangeBillType()">
						<option value="">全部</option>
						<option value="A">攔停、逕舉</option>
						<option value="B">行人、攤販</option>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">
					<input type="button" value="查  詢" name="BSel" onclick="SelectBillData()">
					<input type="button" value="清  除" name="B1" onclick="location='QueryBillBaseData.asp'">
					<input type="button" value="離  開" name="B1" onclick="window.close()">
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
function SelectBillData(){
	if (myForm.BillNo.value=="" && myForm.CarNo.value=="" && myForm.illFID.value=="" && myForm.illName.value==""){
		if (myForm.BillType.value=="B"){
			alert("請至少輸入舉發單號、違規人身份證號、違規人姓名任一項！");
		}else{
			alert("請至少輸入舉發單號、車牌號碼、違規人身份證號、違規人姓名任一項！");
		}
	}else{
		myForm.target="_Blank";
		if (myForm.BillType.value=="A"){
			myForm.action="ViewBillBaseData_Car.asp";
		}else if (myForm.BillType.value=="B"){
			myForm.action="ViewBillBaseData_People.asp";
		}else{
			myForm.action="ViewBillBaseData_All.asp";
		}
		myForm.submit();
		myForm.target="";
		myForm.action="";
	}
}
function ChangeBillType(){
	if (myForm.BillType.value=="B"){
		myForm.CarNo.disabled=true;
		myForm.CarNo.value="";
	}else{
		myForm.CarNo.disabled=false;
	}
}
</script>
</html>
