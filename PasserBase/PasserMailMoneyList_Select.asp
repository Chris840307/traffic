<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title><%
if trim(request("MailSendType"))="S" then
	response.write "列印大宗函件"
else
	response.write "列印郵費單"
end if
%></title>
<%
tmpSQL=request("SQLstr")

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td>選擇郵資</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="2">掛號 25 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="2">郵簡 24 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="3" checked>自行輸入&nbsp;
					<input type="text" name="MailMoneyValue" value="<%
						MailMoney=0
						strMailMoney="select Value from ApConfigure where ID=28"
						set rsMailMoney=conn.execute(strMailMoney) 
						if not rsMailMoney.eof then
							MailMoney=cint(rsMailMoney("Value"))
						end if
						rsMailMoney.close
						set rsMailMoney=nothing

						response.write MailMoney
					%>" size="5" onkeyup="value=value.replace(/[^\d]/g,'')"> 元
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="MailMoneyType" value="4">不計算郵資
				</td>
			</tr>
			<tr>
				<td bgcolor="#FAFAF5" align="center">
					<input type="button" value="列  印" onclick="funMailListCity();">
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
function funMailListCity(){
	MailMoneyValueTmp=myForm.MailMoneyValue.value;
	if (myForm.MailMoneyType[0].checked==false && myForm.MailMoneyType[1].checked==false && myForm.MailMoneyType[2].checked==false && myForm.MailMoneyType[3].checked==false){
		alert("請選擇任一郵資!");
	}else if (myForm.MailMoneyType[2].checked==true && myForm.MailMoneyValue.value==""){
		alert("請輸入郵資!");
	}else{
		if (myForm.MailMoneyType[0].checked==true){
			window.opener.myForm.MailMoneyValue.value=25;
		}else if (myForm.MailMoneyType[1].checked==true){
			window.opener.myForm.MailMoneyValue.value=24;
		}else if (myForm.MailMoneyType[2].checked==true){
			window.opener.myForm.MailMoneyValue.value=MailMoneyValueTmp;
		}else if (myForm.MailMoneyType[3].checked==true){
			window.opener.myForm.MailMoneyValue.value=0;
		}
		window.opener.funPasserReportList();
		window.close();
	}
}
</script>
</html>
