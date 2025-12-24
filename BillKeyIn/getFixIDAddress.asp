<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getFixIDAddress.asp
	'用固定桿編號抓出違規地點
	strFix="select EquipMentID,TypeID,Address,StreetID from FixEquip where EquipMentID='"&trim(request("FixNum"))&"'"
	set rsFix=conn.execute(strFix)
	if not rsFix.eof then
		FixAddr=trim(rsFix("Address"))
		FixStreetID=trim(rsFix("StreetID"))
	end if
	rsFix.close
	set rsFix=nothing
%>setFixIDAddress("<%=trim(request("FixNum"))%>","<%=FixAddr%>","<%=FixStreetID%>");
<%
conn.close
set conn=nothing
%>
