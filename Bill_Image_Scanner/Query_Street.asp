<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>違規地點代碼查詢</title>
<%
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">違規地點代碼查詢</td>
			</tr>
			<tr>
				<td colspan="4">路段名稱：<input type="text" name="StreetName" value="<%=trim(request("StreetName"))%>" size="20">
				<input type="button" value="查詢" onclick="select_street();">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" value="" name="kinds">
				</td>
			</tr>
			<tr bgcolor="#FFCC33">
				<td colspan="4">違規地點代碼列表</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="25%" align="center">代碼</td>
				<td width="75%" align="center">路段</td>
			</tr>
<%
if trim(request("kinds"))="DB_select" then
	strProject="select StreetID,Address from Street where Address Like '%"&trim(request("StreetName"))&"%' order by StreetID"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("StreetID"))%>','<%=trim(rsProject("Address"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#FFFFCC" align="center"><%=trim(rsProject("StreetID"))%>　</td>
				<td><%=trim(rsProject("Address"))%>　</td>
			</tr>
<%	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
elseif trim(request("kinds"))="" and trim(request("OStreet"))<>"" then
	strProject="select StreetID,Address from Street where Address Like '%"&trim(request("OStreet"))&"%' order by StreetID"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("StreetID"))%>','<%=trim(rsProject("Address"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#FFFFCC" align="center"><%=trim(rsProject("StreetID"))%>　</td>
				<td><%=trim(rsProject("Address"))%>　</td>
			</tr>
<%	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing

end if
%>

		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function select_street(){
	myForm.kinds.value="DB_select";
	myForm.submit();
}
function Inert_Data(SCode,SStreet){
	opener.myForm.IllegalAddressIDQry.value=SCode;
	opener.myForm.IllegalAddressQry.value=SStreet;
	window.close();
}
    </script>
</html>