<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>直接人員及共同人員年度總額清冊</title>
<%
 	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing


if trim(request("kinds"))="save" then
	UpdFDateTmp=trim(request("FYear"))&right("00"&trim(request("FMonth")),2)
	strUpdF="Update Apconfigure set value='"&UpdFDateTmp&"' where id=200"
	conn.execute strUpdF

	UpdSDateTmp=trim(request("SYear"))&right("00"&trim(request("SMonth")),2)
	strUpdS="Update Apconfigure set value='"&UpdSDateTmp&"' where id=201"
	conn.execute strUpdS
%>
<script language="JavaScript">
	alert("儲存完成!!");
</script>
<%
end if
%>
<style type="text/css">
<!--
.Text1{
font-weight:bold;
line-height:26px;
font-size:16pt;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" >
<form name=myForm method="post">
  <table border="1" width="450" align="center">
	<tr>
		<td align="center" bgcolor="#CCFFCC" height="35">
			<strong>直接人員及共同人員年度總額清冊</strong>

		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#CCFFFF" >
			統計期間
		</td>
	</tr>
	<tr>
		<td height="70">
			民國 <input type="text" name="sYear" size="6" maxlength="3"> 年
			<input type="text" name="sMonth" size="6" maxlength="2"> 月&nbsp; 至&nbsp;
			民國 <input type="text" name="eYear" size="6" maxlength="3"> 年
			<input type="text" name="eMonth" size="6" maxlength="2"> 月
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#CCFFFF">
			統計單位

		</td>
	</tr>

	<tr>
		<td align="center" bgcolor="#FFCC99">
			<input type="button" value="列印" onclick="OpenRewardList();">
			<input type="button" value="離開" onclick="window.close();">
			<input type="hidden" name="kinds" >
		</td>
	</tr>
  <%

conn.close
set conn=nothing
%>
</table>
</form>
</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">
function OpenRewardList(){
	var error=0;
	var errorString="";
	var sUnitID="";
	var UnitName="";

	if (myForm.sYear.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入起始日期年份。";
	}
	if (myForm.sMonth.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入起始日期月份。";
	}else if (myForm.sMonth.value<1 || myForm.sMonth.value>12){
		error=error+1;
		errorString=errorString+"\n"+error+"：起始日期月份輸入錯誤。";
	}
	if (myForm.eMonth.value != ""){
		if (myForm.eMonth.value<1 || myForm.eMonth.value>12){
			error=error+1;
			errorString=errorString+"\n"+error+"：結束日期月份輸入錯誤。";
		}
	}
	
	if (errorString!=""){
		alert(errorString);
	}else{

		window.open("RewardList_Person_Total.asp?sYear="+myForm.sYear.value+"&sMonth="+myForm.sMonth.value+"&eYear="+myForm.eYear.value+"&eMonth="+myForm.eMonth.value+"&sUnitID="+sUnitID,"RewardList_Person_Total","width=800,height=700,left=10,top=10,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");	    
	}
				    
}


</script>

</html>
