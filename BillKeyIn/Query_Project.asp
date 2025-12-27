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
<title>專案列表</title>
<%
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4">專案列表</td>
			</tr>
			<tr bgcolor="#FAFAF5">
				<td width="15%" align="center">代碼</td>
				<td width="35%" align="center">專案名稱</td>
				<td width="25%" align="center">專案施行起始日期</td>
				<td width="25%" align="center">專案施行結束日期</td>
			</tr>
<%
	strProject="select ProjectID,Name,StartDate,EndDate from Project where RecordStateID=0"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("ProjectID"))%>','<%=trim(rsProject("Name"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#EBE5FF" align="center"><%=trim(rsProject("ProjectID"))%></td>
				<td><%=trim(rsProject("Name"))%></td>
				<td><%=gInitDT(trim(rsProject("StartDate")))%></td>
				<td><%=gInitDT(trim(rsProject("EndDate")))%></td>
			</tr>
<%	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
%>
			<tr>
				<td bgcolor="#1BF5FF" colspan="4" align="center">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
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
function Inert_Data(SCode,SStreet){
	opener.myForm.ProjectID.value=SCode;
	opener.Layer001.innerHTML=SStreet;
	opener.TDProjectIDErrorLog=0;
	window.close();
}
</script>
</html>
