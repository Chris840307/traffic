<%
dim Conn
Set Conn = Server.CreateObject("ADODB.Connection")
Provider="Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=D:\Inetpub\wwwroot\traffic\Common\;Extensions=asc,csv,tab,txt;Persist Security Info=False" 

Conn.Open Provider

Const PageSize = 10

If request("DB_Selt")="Selt" Then 



	strSQL="Select * From BillData.csv where billno='"&request("sys_BillNo")&"'"
	set rs=conn.execute(strSQL)
End if
%>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
						<img src="space.gif" width="20" height="2">
						<b>舉發單號</b>
						<input name="sys_BillNo" type="text" value="<%=request("sys_BillNo")%>" size="20" maxlength="20" class="btn1" onkeyup="value=value.toUpperCase()">		  
   						<img src="space.gif" width="15" height="1">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" >
						<input type="button" name="cancel" value="清除" onClick="location='BillNoArrive.asp'"> 
		</td>
		</table>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th >舉發單號</th>
					<th >輸入序號</th>
					<th >送達日期</th>
					<th >建檔日期</th>
				</tr>
				<%
If request("DB_Selt")="Selt" Then 
  while Not rs.eof
	response.write "<tr bgcolor=""#FFFFFF"" align=""center"">"
	response.write "<td>"&rs("billno")&"</td>"
	response.write "<td>"&rs("SN")&"</td>"
	response.write "<td>"&rs("ArriveDate")&"</td>"
	response.write "<td>"&rs("RecordDate")&"</td>"
	response.write "<tr>"
	rs.movenext
Wend
End if%>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
</form>
</body>
</html>
<script>
  function funSelt()
  {
			myForm.DB_Selt.value="Selt";
			myForm.submit();
  }
</script>
<%
conn.close
%>
