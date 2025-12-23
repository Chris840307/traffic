<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<style type="text/css">
<!--
.style4 {
	font-size: 12px
}
-->
</style>
<title>業管車輛查驗</title>
<% Server.ScriptTimeout = 800 %>
<%
tmpSQL=Session("BillSQLforReport")

if trim(request("kinds"))="Del_NoDci" then
	strToDel="select a.Sn,a.BillNo,a.CarNo from BillBase a,MemberData b,SpecCar c where a.RecordStateID<>-1 and a.RecordMemberID=b.MemberID(+) and a.CarNo=c.CarNO and c.RecordStateID=0"&tmpSQL&" order by a.RecordDate"
	set rsToDel=conn.execute(strToDel)
	if not rsToDel.Bof then
		rsToDel.MoveFirst
	end if
	While Not rsToDel.Eof

		strDel="Update BillBase set BillStatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&trim(rsToDel("SN"))
		conn.execute strDel

		DeleteReason="無"
		NoteTmp=""
		CaseInStatus=""
		ConnExecute "舉發單刪除 單號:"&trim(rsToDel("BillNo"))&" 車號:"&trim(rsToDel("CarNo"))&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
	rsToDel.MoveNext
	Wend
	rsToDel.close
	set rsToDel=nothing
%>
<script language="JavaScript">
	alert("本批業管車輛刪除完成！");
	opener.funSelt();
	window.close();
</script>
<%
end if
	strToDCI="select c.CarNo,c.Note,a.IllegalDate,a.Rule1,a.Rule2 from BillBase a,MemberData b,SpecCar c where a.RecordStateID<>-1 and a.RecordMemberID=b.MemberID(+) and a.CarNo=c.CarNO and c.RecordStateID=0"&tmpSQL&" order by a.RecordDate"
	set rsToDCI=conn.execute(strToDCI)
	If Not rsToDCI.Bof Then
		rsToDCI.MoveFirst
	else
%>
<script language="JavaScript">
	alert("本批舉發單，查無業管車輛！");
	//window.close();
</script>
<%
	end if
%>


</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td width="15%">車號</td>
				<td width="20%">違規時間</td>
				<td width="30%">違規法條</td>
				<td width="35%">備註</td>
			</tr>
<%if trim(Session("SpecUser"))="1" then
	While Not rsToDCI.Eof
%>
			<tr>
				<td><%
				response.write trim(rsToDCI("CarNo"))
				%></td>
				<td><%
				response.write year(rsToDCI("IllegalDate"))&"/"&month(rsToDCI("IllegalDate"))&"/"&day(rsToDCI("IllegalDate"))&" "&hour(rsToDCI("IllegalDate"))&":"&minute(rsToDCI("IllegalDate"))
				%></td>
				<td><%
				if trim(rsToDCI("Rule1"))<>"" and not isNull(rsToDCI("Rule1")) then
					response.write trim(rsToDCI("Rule1"))
				end if
				if trim(rsToDCI("Rule2"))<>"" and not isNull(rsToDCI("Rule2")) then
					response.write ","&trim(rsToDCI("Rule2"))
				end if
				%></td>
				<td><%
				response.write trim(rsToDCI("Note"))
				%></td>
			</tr>
<%
	rsToDCI.MoveNext
	Wend
end if
	rsToDCI.close
	set rsToDCI=nothing
%>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="4">
				<%if trim(Session("SpecUser"))="1" then%>
					<input type="button" value="全部刪除" onclick="funReport_Del();">
				<%end if%>
					<input type="button" value="關閉" onclick="window.close();">
					<input type="hidden" value="" name="kinds">
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
function funReport_Del(){
	myForm.kinds.value="Del_NoDci";
	myForm.submit();
}
</script>
</html>
