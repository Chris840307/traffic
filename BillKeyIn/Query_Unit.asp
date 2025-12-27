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
<title>單位代碼查詢</title>
<%
'行人攤販或車輛
Stype=trim(request("SType"))
'response.write Stype
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4">單位列表</td>
			</tr>
			<tr bgcolor="#FAFAF5">
				<td width="15%" align="center">單位代碼</td>
				<td width="25%" align="center">單位名稱</td>
				<td width="15%" align="center">電話</td>
				<td width="45%" align="center">地址</td>
			</tr>
<%
if Stype="U" then
	strProject="select * from UnitInfo order by UnitID"
elseif Stype="S" then
	strProject="select * from UnitInfo where UnitLevelID in (1,2) order by UnitID"
end if
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("UnitID"))%>','<%=trim(rsProject("UnitName"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#EBE5FF" align="center"><%=trim(rsProject("UnitID"))%>&nbsp;</td>
				<td><%=trim(rsProject("UnitName"))%>&nbsp;</td>
				<td><%=trim(rsProject("TEL"))%>&nbsp;</td>
				<td><%=trim(rsProject("Address"))%>&nbsp;</td>
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
	<%if Stype="U" then%>
		opener.myForm.BillUnitID.value=SCode;
		opener.Layer6.innerHTML=SStreet;
		opener.TDUnitErrorLog=0;
		window.close();
	<%elseif Stype="S" then%>
		opener.myForm.MemberStation.value=SCode;
		opener.Layer5.innerHTML=SStreet;
		opener.TDStationErrorLog=0;
		window.close();
	<%else%>
		opener.myForm.BillUnitID.value=SCode;
		opener.Layer6.innerHTML=SStreet;
		opener.TDUnitErrorLog=0;
		window.close();
	<%end if%>
}
</script>
</html>
