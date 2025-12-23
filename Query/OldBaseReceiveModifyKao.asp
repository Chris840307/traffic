<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OldBillDataDB.ini"-->
<%

'  CREATE TABLE traffic1.OLDPASSERBASE
'   (    BILLNO VARCHAR2(10),
'        RECENO VARCHAR2(30),
'        RECEMONEY NUMBER(38,0),
'        RECORDDATE DATE,
'        ERRREASON VARCHAR2(200)
'   )
if request("DB_Selt") = "Selt" Then

			strSQL = "delete traffic1.oldPasserBase where billno = '" & request("Billno") & "'" 
			conn.execute(strSQL)

			strSQL = "insert into traffic1.oldPasserBase(billno,receNo,receMoney,RecordDate,errReason) "
			strSQL = strSQL&" values('" & Trim(request("Billno")) & "','" & Trim(request("receNo")) & "',0" & Trim(request("receMoney")) & ",sysdate,'" & Trim(request("errReason")) & "')"
'			response.write strsql
			conn.execute(strSQL)

			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			Response.write "window.close();"   
			Response.write "</script>"
Else
			strSQL = "select receNo,receMoney,RecordDate,errReason from traffic1.oldPasserBase where BillNo='" & request("Billno") & "'" 
'			response.write strsql
			receNo    = ""
			receMoney = ""
			errReason = ""
			Set rs = conn.execute(strSQL)
			If Not rs.eof Then 
			  receNo    = rs("receNo")
			  receMoney = rs("receMoney")
			  errReason = rs("errReason")
			End If
	    	Set rs=Nothing			

			

End if		

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人繳費</title>
</head>

<body>
<form name=myForm method="post">

<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr height="25">
					<td bgcolor="#FFCC33" colspan="5"><b>收據單號</td>
				<tr>
				<td>
						<input name="receNo" type="text" value="<%
						if request("DB_Selt") = "Selt" Then
							response.write request("receNo")							
						Else
							response.write receNo
						End If 
						%>" size="40%" maxlength="30" class="btn1">
						</td>
				<tr height="25">
					<td bgcolor="#FFCC33" colspan="5"><b>金額</td>
				<tr>
				<td>
						<input name="receMoney" type="text" value="<%
						if request("DB_Selt") = "Selt" Then
							response.write request("receMoney")							
						Else
							response.write receMoney
						End If 
						%>" size="40%" maxlength="10" class="btn1" onkeypress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==46" style="ime-mode:Disabled">
						</td>
				<tr height="25">
					<td bgcolor="#FFCC33" colspan="5"><b>無法處理原因-<b>無法處理時才填</b></td>
				<tr>
				<td>
						<input name="errReason" type="text" value="<%
						if request("DB_Selt") = "Selt" Then
							response.write request("errReason")							
						Else
							response.write errReason
						End If 
						%>" size="40%" maxlength="50" class="btn1">
						</td>
						<tr>
					<td bgcolor="#FFCC33" colspan="5" align="center"><input type="button" name="btnModify" value="儲存"  onclick="funSelt();" ></td>
			</table>
					<input type="Hidden" name="Billno"  value="<%=request("Billno")%>">
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
<%conn.close%>