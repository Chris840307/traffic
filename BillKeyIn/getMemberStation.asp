<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getMemberStation.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	
	'到案處所
	'106/5/1基隆42改25 旗山85改33
	NoUseStation=""
	If Now > "2018/1/15 0:0:0" Then
		NoUseStation=",'42','85'"
	End If
	'58 105/4/1才使用金門36改成26
	strStation="select DciStationName from Station where StationID='"&trim(request("StationID"))&"' and DciStationID not in ('36'"&NoUseStation&") "
	set rsStation=conn.execute(strStation)
	if not rsStation.eof then
		StationName=trim(rsStation("DciStationName"))
		If trim(request("StationID"))="41" And sys_City<>"高雄市" Then
			StationName=StationName&"(中和辦公室)"
		ElseIf trim(request("StationID"))="46" Then
			StationName=StationName&"(蘆洲辦公室)"
		ElseIf trim(request("StationID"))="60" Then
			StationName=StationName&"(大肚辦公室)"
		ElseIf trim(request("StationID"))="61" Then
			StationName=StationName&"(北屯辦公室)"
		ElseIf trim(request("StationID"))="63" Then
			StationName=StationName&"(豐原辦公室)"
		End if
	end if
	rsStation.close
	set rsStation=nothing
%>setStationName("<%=StationName%>");

<%
conn.close
set conn=nothing
%>
