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
<title>催繳地址調整</title>
<% Server.ScriptTimeout = 800 %>
<%
'tmpSQL=Session("BillSQLforReport")
tmpSQL=replace(trim(request("DciLogSQLforReport")),"@!@","%")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="Update" then
	strUpd="Update BillbaseDcireturn set DriverHomeZip='"&trim(request("sys_DriverHomeADDRESSZip"))&"'" &_
		",DriverHomeAddress='"&trim(request("sys_DriverHomeADDRESS"))&"'" &_
		",Owner='"&Trim(request("sys_OWNER"))&"'" &_
		",OwnerZip='"&trim(request("sys_OWNERADDRESSZip"))&"'" &_
		",OWNERADDRESS='"&trim(request("sys_OWNERADDRESS"))&"'" &_
		",OWNERNOTIFYADDRESS='"&trim(request("sys_OWNERNOTIFYADDRESS"))&"'" &_
		" where CarNo='"&trim(request("sys_CarNo"))&"' and exchangeTypeID='A'"
	conn.execute strUpd
	ConnExecute "地址調整"&trim(request("BillSN"))&":"&strUpd,353

	if trim(request("FileName"))<>"" then
		strUpd2="Update Billbase set DriverZip='"&trim(request("sys_DriverHomeADDRESSZip"))&"'" &_
		",DriverAddress='"&trim(request("sys_DriverHomeADDRESS"))&"'" &_
		",OwnerZip='"&trim(request("sys_OWNERADDRESSZip"))&"'" &_
		",OWNERADDRESS='"&trim(request("sys_OWNERADDRESS"))&"'" &_
		",Owner='"&Trim(request("sys_OWNER"))&"'" &_
		" where ImageFileNameB='"&trim(request("FileName"))&"'"
		conn.execute strUpd2
		ConnExecute "地址調整"&trim(request("BillSN"))&":"&strUpd2,353
	else
'		strUpd2="Update Billbase set DriverZip='"&trim(request("sys_DriverHomeADDRESSZip"))&"'" &_
'		",DriverAddress='"&trim(request("sys_DriverHomeADDRESS"))&"'" &_
'		",OwnerZip='"&trim(request("sys_OWNERADDRESSZip"))&"'" &_
'		",OWNERADDRESS='"&trim(request("sys_OWNERADDRESS"))&"'" &_
'		",Owner='"&Trim(request("sys_OWNER"))&"'" &_
'		" where BillNo is null and CarNo='"&trim(request("sys_CarNo"))&"' " &_
'		" and RecordMemberID in (3480,3503) and ImageFileNameB is null"
	end if
	
	
	
%>
<script language="JavaScript">
	alert("修改完成!")
</script>
<%
	
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="2">催繳地址調整</td>
			</tr>
			<tr>
				<td>通訊地址</td>
				<td><input name="sys_OWNERNOTIFYADDRESS" size="58" value="<%
		sys_OWNERADDRESS=""
		sys_OWNERADDRESSZip=""
		sys_DriverHomeADDRESSZip=""
		sys_DriverHomeADDRESS=""
		sys_OWNER=""
		strSql="select * from billbasedcireturn where CarNo='"&trim(request("sys_CarNo"))&"' and exchangetypeid='A'"
		set rs1=conn.execute(strSql)
		if not rs1.eof then
			response.write trim(rs1("OWNERNOTIFYADDRESS"))
			sys_OWNERADDRESSZip=trim(rs1("OWNERZip"))
			sys_OWNERADDRESS=trim(rs1("OWNERADDRESS"))
			sys_DriverHomeADDRESSZip=trim(rs1("DriverHomeZip"))
			sys_DriverHomeADDRESS=trim(rs1("DriverHomeADDRESS"))
			sys_OWNER=trim(rs1("OWNER"))
		end if
		rs1.close
		set rs1=Nothing
		
		if trim(request("FileName"))<>"" Then
			strAddr2="select * from billbase where ImageFileNameB='"&trim(request("FileName"))&"'"
			Set rsAddr2=conn.execute(strAddr2)
			If Not rsAddr2.eof Then
				If Not IsNull(rsAddr2("OWNERADDRESS")) Then
					sys_OWNERADDRESSZip=trim(rsAddr2("OWNERZip"))
					sys_OWNERADDRESS=trim(rsAddr2("OWNERADDRESS"))
				End If
				If Not IsNull(rsAddr2("DriverADDRESS")) Then
					sys_DriverHomeADDRESSZip=trim(rsAddr2("DriverZip"))
					sys_DriverHomeADDRESS=trim(rsAddr2("DriverADDRESS"))
				End If
			End If
			rsAddr2.close
			Set rsAddr2=Nothing 
		End If 
				%>" >請勿修改通訊地址</td>
			</tr>
			<tr>
				<td>收件人</td>
				<td><input name="sys_OWNER" size="26" value="<%=sys_OWNER%>"></td>
			</tr>
			<tr>
				<td>第一次郵寄地址</td>
				<td><input name="sys_OWNERADDRESSZip" size="6" value="<%=sys_OWNERADDRESSZip%>"> <input name="sys_OWNERADDRESS" size="50" value="<%=sys_OWNERADDRESS%>" ></td>
			</tr>
			<tr>
				<td>第二次郵寄地址</td>
				<td><input name="sys_DriverHomeADDRESSZip" size="6" value="<%=sys_DriverHomeADDRESSZip%>"> <input name="sys_DriverHomeADDRESS" size="50" value="<%=sys_DriverHomeADDRESS%>" ></td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="2">
					<input type="button" value="確 定" name="b1" onclick="funReport_CaseIn();" <%
					
					%>>
					<input type="button" value="離 開" name="b23" onclick="window.close();">
					<input type="hidden" value="" name="kinds">
					<input type="hidden" value="<%=trim(request("sys_CarNo"))%>" name="sys_CarNo">
			<%
					
					%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" id="LayerUp">
					
				</td>
			</td>
			<tr><td colspan="2">
			<hr>
			</td></tr>
			<%
			strSQL="select * from billbasedcireturn where carno='"&trim(request("sys_CarNo"))&"' and exchangetypeid='A'"
			set rscar=conn.execute(strSQL)
			If not rscar.eof Then%>
				<tr><td colspan="2">
				<table width='100%' border='1' align="left" cellpadding="1">
					<tr bgcolor="#FFCC33">
						<td colspan="2">原違規人地址資料</td>
					</tr>
					<tr>
						<td>通訊地址</td>
						<td><%response.write trim(rscar("OWNERNOTIFYADDRESS"))%></td>
					</tr>
					<tr>
						<td>車籍地址</td>
						<td><%=trim(rscar("OWNERADDRESS"))%></td>
					</tr>
					<tr>
						<td>戶籍地址</td>
						<td><%=trim(rscar("DriverHomeADDRESS"))%></td>
					</tr>
				</table>
				</td></tr>
			<%End if
			rscar.close%>

		</table>

	</form>
<%
conn.close
set conn=nothing
%>
</body>

<script language="JavaScript">
function funReport_CaseIn(){
	myForm.kinds.value="Update";
	myForm.submit();
}
</script>
</html>
