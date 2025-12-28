<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<%
	Server.ScriptTimeout=60000

	If Session("UnitLevelID") > 1 Then

		strUit=" and MEMBERSTATION =(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"
	else

		strUit=" and MEMBERSTATION in(select UnitId from Unitinfo where UnitName like '%分局')"
	End if 
	
	strSQL="select sum((case when (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)>0 then 1 else 0 end)) CreditCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) > 0" & _
			" and (select count(1) from PASSERSEND where billsn=passerBase.sn)=0 then 1 else 0 end)) SendCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) = 0 then 1 else 0 end)) JudeCnt," & _
			"sum((case when (select count(1) from PASSERSEND where billsn=passerBase.sn) > 0" & _
			" and (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)=0 then 1 else 0 end)) OtherCnt," & _
			"sum((select count(1) from passerbase pr where TRUNC(SYSDATE-DEALLINEDATE) > 184" & _
			" and Not Exists(select 'N' from PasserJude where billsn=pr.SN) and SN=PasserBase.SN)) NotJudeCnt," & _
			"sum((select count(1) from passerbase pr where Exists(select 'Y' from PasserJude where TRUNC(SYSDATE-JudeDate) > 184" & _
			" and billsn=pr.SN) and Not Exists(select 'N' from PasserSend where billsn=pr.SN) and SN=PasserBase.SN)) NotPasserSend," & _
			"sum((select count(1) from passerbase pr where TRUNC(SYSDATE-DEALLINEDATE) > 184" & _
			" and Not Exists(select 'N' from PasserJude where billsn=pr.SN) and SN=PasserBase.SN)) NotJudeCnt," & _
			"sum((case when (select count(1) cnt from PasserSendDetail where billsn=passerbase.sn)=1 " & _
			" and exists(select 'Y' from PasserSendDetail where TRUNC(sysdate-SENDDATE)>1097 and billsn=passerbase.sn) " & _
			" and not Exists(select 'N' from PasserCreditor where PETITIONDATE is not null and BillSn=PasserBase.SN) " & _
			" then 1 else 0 end)) Send3Year " & _
			" from passerbase where recordstateid=0 and billstatus <> 9"&strUit

	set rs=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人裁罰管制表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post">
<table height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33">慢車行人裁罰管制表</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="1" bgcolor="#FFFFFF" width="100%">
				<tr  height="30" align="right">
					<td colspan="1"><b>舉發未完成收繳件數</b> </td>
					<td colspan="11" align="left"> <font size="3">
						<%=cdbl(rs("CreditCnt"))+cdbl(rs("SendCnt"))+cdbl(rs("JudeCnt"))+cdbl(rs("OtherCnt"))%> 
					</font>
					</td>
				</tr>
				<tr>
					<td colspan="6"></td>
				</tr>
				<tr  height="30"  align="right" >
					<td><b>未裁罰件數</b></td>
					<td >
						<%=rs("JudeCnt")%> 
					</td>
					<td><b>未移送件數</b></td>
					<td>
						<%=rs("SendCnt")%> 
					</td>
					<td><b>取得債權憑證</b></td>
					<td>
						<%=rs("CreditCnt")%> 
					</td>
					<td><b>移送執行中</b></td>
					<td>
						<%=rs("OtherCnt")%> 
					</td>
					<td><b>逾期未裁決</b></td>
					<td>
						<%=rs("NotJudeCnt")%> 
					</td>
					<td><b>逾期未移送</b></td>
					<td>
						<%=rs("NotPasserSend")%> 
					</td>
					<td><b>移送逾三年無債權</b></td>
					<td>
						<%=rs("Send3Year")%> 
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">
		各單位裁罰狀況
		</td>
	</tr>
	<%
	rs.close

	strSQL="select (select UnitName from Unitinfo where UnitID=PasserBase.MEMBERSTATION) UnitName," & _
	        "(select UnitOrder from Unitinfo where UnitID=PasserBase.MEMBERSTATION) UnitOrder," & _
			"sum((case when (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)>0 then 1 else 0 end)) CreditCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) > 0 " & _
			"and (select count(1) from PASSERSEND where billsn=passerBase.sn)=0 then 1 else 0 end)) SendCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) = 0 then 1 else 0 end)) JudeCnt," & _
			"sum((case when (select count(1) from PASSERSEND where billsn=passerBase.sn) > 0 " & _
			"and (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)=0 then 1 else 0 end)) OtherCnt," & _
			"sum((select count(1) from passerbase pr where TRUNC(SYSDATE-DEALLINEDATE) > 184" & _
			" and Not Exists(select 'N' from PasserJude where billsn=pr.SN) and SN=PasserBase.SN)) NotJudeCnt," & _
			"sum((select count(1) from passerbase pr where Exists(select 'Y' from PasserJude where TRUNC(SYSDATE-JudeDate) > 184" & _
			" and billsn=pr.SN) and Not Exists(select 'N' from PasserSend where billsn=pr.SN) and SN=PasserBase.SN)) NotPasserSend," & _
			"sum((case when (select count(1) cnt from PasserSendDetail where billsn=passerbase.sn)=1 " & _
			" and exists(select 'Y' from PasserSendDetail where TRUNC(sysdate-SENDDATE)>1097 and billsn=passerbase.sn) " & _
			" and not Exists(select 'N' from PasserCreditor where PETITIONDATE is not null and BillSn=PasserBase.SN) " & _
			" then 1 else 0 end)) Send3Year " & _
			"from passerbase where recordstateid=0 and billstatus <> 9"&strUit&" group by MEMBERSTATION order by UnitOrder"
	
	set rs=conn.execute(strSQL)
	%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="1" cellpadding="1" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="30">分局</th>
					<th height="34">未裁罰件數</th>
					<th height="34">未移送件數</th>
					<th height="34">取得憑證件數</th>
					<th height="34">移送執行中</th>
					<th height="34">逾期未裁決</th>
					<th height="34">逾期未移送</th>
					<th height="34">移送逾三年無債權</th>
					<th height="34">操作</th>
				</tr><%
					
					while Not rs.eof 
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"
						
						response.write "<td align='left'>"&rs("UnitName")&"</td>"
						response.write "<td align='right'>"&rs("JudeCnt")&"</td>"
						response.write "<td align='right'>"&rs("SendCnt")&"</td>"
						response.write "<td align='right'>"&rs("CreditCnt")&"</td>"
						response.write "<td align='right'>"&rs("OtherCnt")&"</td>"
						response.write "<td align='right'>"&rs("NotJudeCnt")&"</td>"
						response.write "<td align='right'>"&rs("NotPasserSend")&"</td>"
						response.write "<td align='right'>"&rs("Send3Year")&"</td>"

						response.write "<td>"

						response.write "<input type=""button"" name=""Detail"" value=""詳 細"" onclick=""funchgExecel3('"&rs("UnitName")&"');"""
						response.write ">"

						response.write "</td>"
						response.write "</tr>"
						
						rs.movenext
					wend
					rs.close%>
			</table>
		</td>
	</tr>
</table>

<input type="Hidden" name="Sys_UnitName" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funchgExecel3(Sys_UnitName){

	myForm.Sys_UnitName.value=Sys_UnitName;

	myForm.action="PasserReportUnitDetail.asp";
	myForm.target="Execel3";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
</script>
<%conn.close%>