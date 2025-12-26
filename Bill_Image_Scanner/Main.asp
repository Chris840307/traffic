<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">
	window.focus();
</script>
<style type="text/css">
<!--

-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>獎勵金系統</title>
<%
 	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing


%>
</head>
<body leftmargin="0" topmargin="25" marginwidth="0" marginheight="0" >
<form name=myForm method="post">
  <table border="0" width="350" align="center">
  <%if Trim(Session("Unit_ID"))="0807" then%>
	<tr>
		<td >
		<%
		RewardTotal=0
		strR1="select * from Apconfigure where ID=46"
		set rsR1=conn.execute(strR1)
		if not rsR1.eof then
			RewardTotal=trim(rsR1("value"))
		end if
		rsR1.close
		set rsR1=nothing
		%>
		發放月份
		<input type="text" Name="RewardMonth1" value="<%'RewardTotal%>" size="5" readonly>年
		<input type="text" Name="RewardMonth2" value="<%'RewardTotal%>" size="5" readonly>月
		<input type="button" value="金額設定" onclick="MonthRewardSet();" style=" width: 80px;">
		<br>
		獎勵金總額
		<input type="text" Name="RewardAll" value="<%'RewardTotal%>" size="12" readonly> 元
			<br>
		共同人員 28 % <input type="text" Name="Reward28" value="<%
		'response.write Round(RewardTotal*0.28)
		%>" size="15" readonly> 元  <br>
		直接人員 72 % <input type="text" Name="Reward72" value="<%
		'response.write Round(RewardTotal*0.72)
		%>" size="15" readonly> 元  

		</td>
	</tr>
	<%end if%>
	<tr>
		<td >
		
			<input type="button" value="共同人員比率設定" onclick="setCommonRewardRercent();" style=" width: 280px;">
			<br>
			<input type="button" value="共同人員資料設定" onclick="setCommonRewardMem();" style=" width: 280px;">
			<br>
			<!-- <input type="button" value="減發獎勵金人員設定" onclick="getReward_SpecMem_Set();" style=" width: 280px;">
			<br> -->
			<input type="button" value="減發獎勵金人員設定" onclick="getReward_SpecMem_Set2();" style=" width: 280px;">
			<br>
		<%if Trim(Session("Unit_ID"))="0807" then%>
			<input type="button" value="獎勵金計算作業" onclick="getRewardList_Person_SubUnit();" style=" width: 280px;">
			<br>
		<%end if%>
			<input type="button" value="扣款資料表" onclick="getReward_SpecMem();" style=" width: 280px;">
			<br>
			<input type="button" value="實領獎金發放金額表" onclick="getReward_Unit();" style=" width: 280px;">
			<br>
		<%if Trim(Session("Unit_ID"))="0807" then%>
			<input type="button" value="作業項目結餘" onclick="getRewardList_Balance();" style=" width: 280px;">
			<br>
		<%end if%>
			<input type="button" value="獎勵金轉撥款帳清單" onclick="getReward_Bank();" style=" width: 280px;">
			<br>
		
			<input type="button" value="列印直接人員及共同人員請領清冊" onclick="Print_RewardList_Person();" style=" width: 280px;">
		<%if Trim(Session("Unit_ID"))="0807" then%>
			<input type="button" value="列印分隊、拖吊場請領清冊" onclick="Print_RewardList_Person_Unit();" style=" width: 280px;">
			
		<%end if%>
		<%if Trim(Session("Unit_ID"))="0807" then%>
			<input type="button" value="匯出電子檔" onclick="Out_RewardList_Txt();" style=" width: 280px;">
			<br>
			<input type="button" value="列印直接人員及共同人員年度總額清冊" onclick="Print_RewardList_Person_Total();" style=" width: 280px;">
			<br>
			<input type="button" value="系統設定" onclick="System_Set();" style=" width: 280px;">
		<%end if%>

		</td>
	</tr>
  <%

conn.close
set conn=nothing
%>
</table>
</form>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<script language="JavaScript">

function MonthRewardSet(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;
	AnalyzeMoney=myForm.Reward72.value;
	AnalyzeMoney28=myForm.Reward28.value;
	window.open("MonthReward_Set.asp?AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney+"&AnalyzeMoney28="+AnalyzeMoney28,"MonthReward_Set","width=420,height=520,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function getRewardList_Person_SubUnit(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;
	AnalyzeMoney=myForm.Reward72.value;
	AnalyzeMoney28=myForm.Reward28.value;
	var error=0;
	var errorString="";
	if(myForm.RewardMonth1.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請先設定發放月份!!";
	}

	if (error>0){
		alert(errorString);
	}else{
		window.open("getRewardList_Person_SubUnit_Set.asp?AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney+"&AnalyzeMoney28="+AnalyzeMoney28+"&tbDate1="+myForm.RewardMonth1.value+"&tbDate2="+myForm.RewardMonth2.value,"getRewardList_Person1","width=420,height=520,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
	}	
}

function getRewardList_Balance(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("getReward_Balance_Set.asp","getRewardList_Balance_Set","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function getReward_Unit(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;
	window.open("getReward_Unit_Set.asp","getReward_Unit_Set1","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function getReward_Bank(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("getReward_Bank_Set.asp","getReward_Bank_Set1","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function getReward_SpecMem(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("getReward_SpecMem_Set.asp","getReward_SpecMem","width=420,height=480,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no");
}

function getReward_SpecMem_Set(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("/traffic/Reward/PunisheMgr_Ksc.aspx?LoginID=<%=trim(Session("User_ID"))%>&sUnitID=<%=trim(Session("Unit_ID"))%>&UnitLevel=<%=trim(Session("UnitLevelID"))%>","getReward_SpecMem","width=720,height=520,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=yes");
}

function getReward_SpecMem_Set2(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("PunisheMgr_Ksc.asp","getReward_SpecMem2","width=920,height=420,left=50,top=50,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=yes");
}


function setCommonRewardRercent(){
	window.open("setCommonRewardRercent.asp","getRewardList_Person2","width=720,height=680,left=150,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}
function setCommonRewardMem(){
	window.open("setCommonRewardMem.asp","getRewardList_Person3","width=920,height=680,left=50,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}
function OpenRewardSet(){
	var RewardAll=myForm.RewardAll.value;

	runServerScript("setMonthReward.asp?RewardAll="+RewardAll);
}
function Print_RewardList_Person(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;
	window.open("Print_RewardList_Person_Set.asp","getRewardList_Person3","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function Print_RewardList_Person_Unit(){
	var AnalyzeType=0;
	var AnalyzeMoney=0;

	window.open("Print_RewardList_Person_Unit_Set.asp","getRewardList_Person3","width=420,height=480,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function Out_RewardList_Txt(){

	window.open("RewardList_Person_TXT_Set.asp","RewardList_Person_TXT","width=420,height=580,left=150,top=0,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}

function Print_RewardList_Person_Total(){
	window.open("RewardList_Person_Total_Set.asp","RewardList_Person_Total_Set","width=520,height=580,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no");
}
function System_Set(){
	window.open("System_Set.asp","System_Set","width=520,height=500,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
}
</script>

</html>
