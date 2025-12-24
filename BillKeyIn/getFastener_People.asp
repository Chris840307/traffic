<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getFastener.asp
	'保管物品
	strFastener="select Content from Code where ID='"&trim(request("FastenerID"))&"' and TypeID=2"
	set rsFastener=conn.execute(strFastener)
	if not rsFastener.eof then
		FastenerName=trim(rsFastener("Content"))
	end if
	rsFastener.close
	set rsFastener=nothing
%>setFastenerName("<%=trim(request("FastenerOrder"))%>","<%=FastenerName%>");
<%
conn.close
set conn=nothing
%>
