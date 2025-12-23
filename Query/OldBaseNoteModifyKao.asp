<!--#include virtual="Traffic/Common/OlddbAccessKao.ini"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%




if request("DB_Selt")="Selt" Then

			strSQL="Update FMaster set [note]='"&Trim(request("note"))&"' where FSEQ='"&request("Billno")&"'" 
'			response.write strsql
			conn1.execute(strSQL)
			conn2.execute(strSQL)
			conn3.execute(strSQL)



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
			End If
			
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
			
			Set rs=conn4.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If

			Set rs=conn5.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If

			Set rs=conn6.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If			

			Set rs=conn7.execute(strSQL)
			If Not rs.eof Then 
				If Trim(rs("Note"))<>"" Then 
					note=rs("Note")&note
				End If
				Set rs=Nothing
			End If

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
