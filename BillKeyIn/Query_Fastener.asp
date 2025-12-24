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
<title>代保管物代碼列表</title>
<%
FaOrder=trim(request("FaOrder"))
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">代保管物代碼列表</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="40%" align="center">代碼</td>
				<td width="60%" align="center">代保管物名稱</td>
			</tr>
<%
	strStation="select ID,Content from DCIcode where TypeID=6"
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then rsStation.MoveFirst 
	While Not rsStation.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsStation("ID"))%>','<%=trim(rsStation("Content"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#FFFFCC" align="center"><%=trim(rsStation("ID"))%></td>
				<td><%=trim(rsStation("Content"))%></td>
			</tr>
<%	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing
%>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
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
function Inert_Data(FCode,FValue){
	<%if FaOrder="1" then%>
		opener.myForm.Fastener1.value=FCode;
		opener.myForm.Fastener1Val.value=FValue;
		opener.Layer8.innerHTML=FValue;
		opener.TDFastenerErrorLog1=0;
		window.close();
	<%elseif FaOrder="2" then%>
		opener.myForm.Fastener2.value=FCode;
		opener.myForm.Fastener2Val.value=FValue;
		opener.Layer9.innerHTML=FValue;
		opener.TDFastenerErrorLog2=0;
		window.close();	
	<%elseif FaOrder="3" then%>
		opener.myForm.Fastener3.value=FCode;
		opener.myForm.Fastener3Val.value=FValue;
		opener.Layer10.innerHTML=FValue;
		opener.TDFastenerErrorLog3=0;
		window.close();
	<%end if%>
}
</script>
</html>
