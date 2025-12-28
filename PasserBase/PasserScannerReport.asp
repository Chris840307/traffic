<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人裁罰管制表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post">
<table width="900" height="100%" border="0">	
	<tr>
		<td bgcolor="#FFCC33" height="33">
		各單位裁罰狀況
		</td>
	</tr>
	<%

	If Session("UnitLevelID") > 1 Then

		strUit=" and MEMBERSTATION =(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"
	else

		strUit=" and MEMBERSTATION in(select UnitId from Unitinfo where UnitName like '%分局')"
	End if 

	strSQL="select (select UnitName from Unitinfo where UnitID=PasserBase.MEMBERSTATION) UnitName," & _
	        "(select UnitOrder from Unitinfo where UnitID=PasserBase.MEMBERSTATION) UnitOrder," & _
			"sum((case when (select count(1) from PasserCreditor where PetitionDate is not null and TRUNC(sysdate-PetitionDate)>90 and imagefilename is null and imagefilename2 is null and imagefilename3 is null and imagefilename4 is null and billsn=passerBase.sn)>0 then 1 else 0 end)) CreditCnt," & _
			"sum((case when (select count(1) from PasserSendDetail where TRUNC(sysdate-SendDate)>7 and not exists(select 'N' from PasserImage where BillSN=passerBase.sn and PkeySN=PasserSendDetail.sn and ImgKindID='3' and ImgTypeID='1') and billsn=passerBase.sn)>0 then 1 else 0 end)) SendCnt," & _
			"sum((case when (select count(1) from PasserJude where TRUNC(sysdate-JudeDate)>30 and billsn=passerBase.sn) > 0 and not exists(select 'N' from PasserImage where BillSN=passerBase.sn and PkeySN=passerBase.sn and ImgKindID='1' and ImgTypeID='1') then 1 else 0 end)) JudeCnt," & _
			"sum((case when (select count(1) from PasserSendArrived where TRUNC(sysdate-ArrivedDate)>30 and ImageFileName is null and PASSERSN=passerBase.sn) > 0 then 1 else 0 end)) ArrivedCnt " & _
			"from passerbase where recordstateid=0 and to_number(to_char(illegaldate,'YYYY')) between to_number(to_char(sysdate,'YYYY'))-10 and to_number(to_char(sysdate,'YYYY')) and billstatus <> 9 "&strUit&" group by MEMBERSTATION order by UnitOrder"
	
	set rs=conn.execute(strSQL)
	%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="1" cellpadding="1" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="30">分局</th>
					<th height="34">送達未上傳</th>
					<th height="34">裁決未上傳</th>
					<th height="34">移送未上傳</th>
					<th height="34">債權未上傳</th>
					<th height="34">操作</th>
				</tr><%
					
					while Not rs.eof 
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"
						
						response.write "<td align='left'>"&rs("UnitName")&"</td>"
						response.write "<td align='right'>"&rs("ArrivedCnt")&"</td>"
						response.write "<td align='right'>"&rs("JudeCnt")&"</td>"
						response.write "<td align='right'>"&rs("SendCnt")&"</td>"
						response.write "<td align='right'>"&rs("CreditCnt")&"</td>"

						response.write "<td>"

						response.write "<input type=""button"" name=""Detail"" value=""詳 細"" onclick=""funSetQry('"&rs("UnitName")&"');"""
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

function funSetQry(uitid){

	myForm.Sys_UnitName.value=uitid;

	myForm.action="PasserScannerReportUnit.asp";
	myForm.target="ScannerUnit";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
</script>
<%conn.close%>