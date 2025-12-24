<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getBillUnitID.asp
	'舉發單位
	strUnit="select UnitName from UnitInfo where UnitID='"&trim(request("BillUnitID"))&"'"
	set rsUnit=conn.execute(strUnit)
	if not rsUnit.eof then
		UnitName=trim(rsUnit("UnitName"))
	end if
	rsUnit.close
	set rsUnit=nothing
%>setUnitName("<%=UnitName%>");

<%
conn.close
set conn=nothing
%>
