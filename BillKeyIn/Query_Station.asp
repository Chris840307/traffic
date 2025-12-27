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
<title>到案監理所列表</title>
<%
	'抓縣市
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
			<tr bgcolor="#1BF5FF">
				<td colspan="4">到案監理所列表</td>
			</tr>
			<tr bgcolor="#FAFAF5">
				<td width="8%" align="center">代碼</td>
				<td width="25%" align="center">監理所名稱</td>
				<td width="42%" align="center">監理所住址</td>
				<td width="25%" align="center">監理所電話</td>
			</tr>
<%
	'106/5/1基隆42改25 旗山85改33
	NoUseStation=""
	If Now > "2018/1/15 0:0:0" Then
		NoUseStation=",'42','85'"
	End if
	'58 105/4/1才使用金門36改成26
	strStation="select a.DciStationID,a.DciStationName,a.StationAddress,a.StationTel from Station a," &_
			"(select distinct(StationID) from Station) b where a.DciStationID=b.StationID and DciStationID not in ('36'"&NoUseStation&") order by a.StationID"
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then rsStation.MoveFirst 
	While Not rsStation.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsStation("DciStationID"))%>','<%=trim(rsStation("DciStationName"))%><%
				If trim(rsStation("DciStationID"))="41" And sys_City<>"高雄市" Then
					response.write "(中和辦公室)"
				ElseIf trim(rsStation("DciStationID"))="46" Then
					response.write "(蘆洲辦公室)"
				ElseIf trim(rsStation("DciStationID"))="60" Then
					response.write "(大肚辦公室)"
				ElseIf trim(rsStation("DciStationID"))="61" Then
					response.write "(北屯辦公室)"
				ElseIf trim(rsStation("DciStationID"))="63" Then
					response.write "(豐原辦公室)"
				End if
			%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#EBE5FF" align="center"><%=trim(rsStation("DciStationID"))%></td>
				<td><%=trim(rsStation("DciStationName"))%><%
				If trim(rsStation("DciStationID"))="41" And sys_City<>"高雄市" Then
					response.write "(中和辦公室)"
				ElseIf trim(rsStation("DciStationID"))="46" Then
					response.write "(蘆洲辦公室)"
				ElseIf trim(rsStation("DciStationID"))="60" Then
					response.write "(大肚辦公室)"
				ElseIf trim(rsStation("DciStationID"))="61" Then
					response.write "(北屯辦公室)"
				ElseIf trim(rsStation("DciStationID"))="63" Then
					response.write "(豐原辦公室)"
				End if
				%></td>
				<td><%=trim(rsStation("StationAddress"))%></td>
				<td><%=trim(rsStation("StationTel"))%></td>
			</tr>
<%	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing
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
	opener.myForm.MemberStation.value=SCode;
	opener.Layer5.innerHTML=SStreet;
	opener.TDStationErrorLog=0;
	window.close();
}
</script>
</html>
