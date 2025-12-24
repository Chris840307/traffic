<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>檢查設定</title>
<%

if trim(request("kinds"))="Add" then
	strUpd="Update apconfigure set value='"&Trim(request("Chk"))&"' where id=777"
	conn.execute strUpd
%>
<script language="JavaScript">
	alert("儲存成功!");
</script>
<%
end If

CheckFlag=0
str1="select * from apconfigure where id=777"
Set rs1=conn.execute(str1)
If Not rs1.eof Then
	CheckFlag=Trim(rs1("value"))
End If
rs1.close
Set rs1=Nothing 

strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td >
				<%If sys_City="南投縣" then%>
				建檔是否檢查6分鐘內有無同車號、同法條案件
				<%else%>
				建檔是否檢查一日內有無同車號、同法條案件
				<%End If %>
				</td>
			</tr>
			<tr>
				<td>
					是<input type="radio" name="Chk" value="1" <%
					If CheckFlag="1" Then
						response.write "checked"
					End If 
					%>>
					<br>
					否<input type="radio" name="Chk" value="0" <%
					If CheckFlag="0" Then
						response.write "checked"
					End If 
					%>>
				</td>
			</tr>
			<tr>
				<td colspan="2">
				<input type="button" value="儲存" onclick="Add_LawPlus();">
				<input type="hidden" value="" name="kinds">
				<input type="hidden" value="" name="LawPlusID">
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

function Add_LawPlus(){

		myForm.kinds.value="Add";
		myForm.submit();

}
</script>
</html>
