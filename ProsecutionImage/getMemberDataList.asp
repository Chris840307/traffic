<!--#include virtual="/traffic/Common/db.ini"-->
<%
' ÀÉ®×¦WºÙ¡G getMemberList.asp
' 
	if trim(request("UnitID"))<>"" then
		strMem="select MemberID,chName from MemberData where UnitID='"&trim(request("UnitID"))&"' order by MemberID"
		set rsMem=conn.execute(strMem)
		If Not rsMem.Bof Then rsMem.MoveFirst
		While Not rsMem.Eof
%>setMemberDataList('<%=trim(rsMem("MemberID"))%>','<%=trim(rsMem("chName"))%>');
<%
		rsMem.MoveNext
		Wend
		rsMem.Close
		set rsMem=nothing
	end if



conn.close
set conn=nothing
%>
