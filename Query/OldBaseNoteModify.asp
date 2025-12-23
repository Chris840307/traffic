<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->

<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
sys_City="南科"
If sys_City="彰化縣" Then 
	set conn4=Server.CreateObject("ADODB.connection")
	conn4.Provider="Microsoft.Jet.OLEDB.4.0"
	conn4.Open "D:\Inetpub\wwwroot\Traffic\olddb4.mdb"
End If

If sys_City="南科" Then 
%>
<!--#include virtual="Traffic/Common/Oldsp.ini"-->
<%
Else
%>
<!--#include virtual="Traffic/Common/OldbAccessHualien.ini"-->

<%
End if

if request("DB_Selt")="Selt" Then

			strSQL="Update FMaster set [note]='"&Trim(request("note"))&"' where FSEQ='"&request("Billno")&"'" 
'			response.write strsql
			conn1.execute(strSQL)
			If sys_City<>"南科" then
				conn2.execute(strSQL)
				conn3.execute(strSQL)
				If sys_City="彰化縣" Then conn4.execute(strSQL)

				If sys_City<>"彰化縣" Then
					strSQL="Update FMaster_S set [note]='"&Trim(request("note"))&"' where FSEQ='"&Request("Billno")&"'" 
					conn1.execute(strSQL)
				End if
			End if
			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			Response.write "window.close();"   
			Response.write "</script>"
Else
			strSQL="select [note] from FMaster where FSEQ='"&request("Billno")&"'" 
'			response.write strsql
			note=""
			Set rs=conn1.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End if
			If sys_City<>"南科" then
			Set rs=conn2.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End if

			Set rs=conn3.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If
			
		If sys_City="彰化縣" Then
			Set rs=conn4.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If
		End if

		If sys_City<>"彰化縣" Then
			strSQL="select [note] from FMaster_s where FSEQ='"&request("Billno")&"'" 
			Set rs=conn1.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If
		End if
			End if

End if		

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>備註</title>
</head>

<body>
<form name=myForm method="post">

<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr height="25">
					<td bgcolor="#FFCC33" colspan="5"><b>備註
						<%If sys_City="宜蘭縣" Then response.write "-實收金額可填寫在備註欄位"%>
					</b></td>
				<tr>
				<td>
						<input name="Note" type="text" value="<%
						if request("DB_Selt")="Selt" Then
							response.write request("Note")							
						Else
							response.write note
						End If 
						%>" size="80%" maxlength="100" class="btn1">
						</td>
				<tr height="25" align="center">
					<td bgcolor="#FFCC33" colspan="5"><input type=button name=btnModify value="儲存"  onclick="funSelt();" ></td>
			</table>
					<input type="Hidden" name="Billno" value="<%=request("Billno")%>">
					<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
</form>
</body>
<script language="javascript">
	function funSelt()
	{
			myForm.DB_Selt.value="Selt";
			myForm.submit();
	}
</script>

</html>
