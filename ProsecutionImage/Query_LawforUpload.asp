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
<!--#include virtual="/traffic/Common/css.txt"-->
<title>法條查詢</title>
<%
LawOrder=trim(request("LawOrder"))
theRuleVer=trim(request("RuleVer"))
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="7">法條查詢</td>
			</tr>
			<tr>
				<td colspan="7">
				代碼：<input type="text" name="LawID" value="">
				<input type="button" name="BB1" value="查詢" onclick="DB_Select();">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				</td>
			<tr>
			<tr bgcolor="#FFCC33">
				<td colspan="7">法條列表</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="10%" align="center">法條代碼</td>
				<td width="9%" align="center">簡式車種</td>
				<td width="40%" align="center">法條內容</td>
				<td width="10%" align="center">罰金<br>Level1</td>
				<td width="10%" align="center">罰金<br>Level2</td>
				<td width="10%" align="center">罰金<br>Level3</td>
				<td width="10%" align="center">罰金<br>Level4</td>
			</tr>
<%
if trim(request("kinds"))="DB_Select" then
	strProject="select ItemID,CarSimpleID,IllegalRule,Level1,Level2,Level3,Level4 from Law where ItemID Like '"&trim(request("LawID"))&"%' and Version='"&theRuleVer&"' order by ItemID"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("ItemID"))%>','<%=trim(rsProject("IllegalRule"))%>','<%=trim(rsProject("Level1"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#FFFFCC" align="center"><%=trim(rsProject("ItemID"))%></td>
				<td><%
				if trim(rsProject("CarSimpleID"))="1" then
					response.write "汽車"
				elseif trim(rsProject("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rsProject("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rsProject("CarSimpleID"))="4" then
					response.write "輕機"
				end if
				%>&nbsp;</td>
				<td><%=trim(rsProject("IllegalRule"))%></td>
				<td><%=trim(rsProject("Level1"))%></td>
				<td><%=trim(rsProject("Level2"))%>&nbsp;</td>
				<td><%=trim(rsProject("Level3"))%></td>
				<td><%=trim(rsProject("Level4"))%></td>
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
function DB_Select(){
	myForm.kinds.value="DB_Select";
	myForm.submit();
}
function Inert_Data(LCode,LValue,LMoney){
	opener.myForm.Rule1.value=LCode;
	opener.myForm.ForFeit1.value=LMoney;
	opener.Layer1.innerHTML=LValue;
	opener.TDLawErrorLog1=0;
	window.close();

}
</script>
</html>
