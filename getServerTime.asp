<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getServerTime.asp
	'抓系統時間
	strTime="目前時間  "&hour(now)&"："&minute(now)
%>

setServerTime="<%=strTime%>";
LayerTime.innerHTML=setServerTime;
<%
conn.close
set conn=nothing
%>
