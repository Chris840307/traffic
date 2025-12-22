<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>地址查詢</title>

<!--#include virtual="Traffic/Common/AddressAccess.ini"-->

<%

'組成查詢SQL字串
		strwhere=" where 1=1 "
if request("DB_Selt")="Selt" then


		
		if request("Road")<>"" Then strwhere=strwhere&" and Road like '%"&request("Road")&"%'"

		if request("City")<>"" Then strwhere=strwhere&" and City = '"&request("City")&"'"
		if request("area")<>"" Then strwhere=strwhere&" and area = '"&request("area")&"'"


					strSQL="select distinct Road from address " & strwhere &" order by 1"

					set rs=conn.execute(strSQL)
End if

%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr height="25">
					<td bgcolor="#FFCC33" colspan="5"><b>地址查詢</b></font></td>
				</tr>			
				<tr>
					<td>
							縣市/鄉鎮【市】區 
								 <Select Name="City" onchange="myForm.submit();">
								 <% strSQL="Select distinct City from Address order by 1"
								 			set rsCity=conn.execute(strSQL)
								 While Not rsCity.eof %>
								   <option value="<%=rsCity("City")%>" <%if trim(request("City"))=rsCity("City") then response.write " Selected"%>><%=rsCity("City")%></option>
								   
								<%

								rsCity.movenext
									Wend
									Set rsCity=Nothing
								%>

									</select>
								 <Select Name="Area">
								 <% strSQL="Select distinct Area from Address where City='"&trim(request("City"))&"' order by 1"
								 			set rsCity=conn.execute(strSQL)
								 While Not rsCity.eof %>
								   <option value="<%=rsCity("Area")%>" <%if trim(request("Area"))=rsCity("Area") then response.write " Selected"%>><%=rsCity("Area")%></option>
								<%
									rsCity.movenext
									Wend
									Set rsCity=Nothing
									%>
								 路(街)名或鄰里名稱
								<input name="Road" type="text" value="<%=request("Road")%>" size="28" maxlength="20" class="btn1">
						<br>
   						<img src="space.gif" width="100" height="1">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" >
						<input type="button" name="cancel" value="清除" onClick="location='AddressQry.asp'"> 
						
					  </td>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	
	<tr height="30">
		<td bgcolor="#FFCC33" class="style3">
		</td>
	</tr>
	
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th nowrap>路名</th>
				</tr>
				<%
	if request("DB_Selt")="Selt"  then
				While Not rs.eof 
				response.write "<tr bgcolor=""#FFFFFF"" align=""center"">"
				response.write "<td>"
				response.write rs("road")&"&nbsp;"
				response.write "</td>"
						rs.movenext
				wend
	End if
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
				
				%>
			</table>
		</td>
	</tr>
</table>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<script language="javascript">
function funSelt()
{
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}
</script>
</form>
</body>
</html>

<%
		conn.close
		set conn=nothing	
%>