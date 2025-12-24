<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getIllStreet.asp
	'違規地點
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	If trim(sys_City)="台東縣" Then
		if Session("UnitLevelID")<>"1" then
			strPlus=" and UnitID='"&session("Unit_ID")&"'"
		end if
	elseIf trim(sys_City)="高雄縣" Then
		strPlus=" and UnitID='"&session("Unit_ID")&"'"
	End if
	strAddress="select Address from Street where StreetID='"&trim(request("illAddrID"))&"'"&strPlus
	set rsAddress=conn.execute(strAddress)
	if not rsAddress.eof then
		AddressName=trim(rsAddress("Address"))
	end if
	rsAddress.close
	set rsAddress=nothing
%>setIllStreetName("<%=AddressName%>");
<%
conn.close
set conn=nothing
%>
