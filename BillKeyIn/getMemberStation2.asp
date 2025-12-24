<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getMemberStation2.asp
	'到案處所(行人到案處所為分局)
	strStation="select UnitName from UnitInfo where UnitID='"&trim(request("StationID"))&"' and UnitLevelID='2'"
	set rsStation=conn.execute(strStation)
	if not rsStation.eof then
		StationName=trim(rsStation("UnitName"))
	end if
	rsStation.close
	set rsStation=nothing
%>setStationName("<%=StationName%>");

<%
conn.close
set conn=nothing
%>
