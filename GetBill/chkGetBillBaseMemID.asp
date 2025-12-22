<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
startHead = trim(Request("startHead"))
startTail = trim(Request("startTail"))
endHead = trim(Request("endHead"))
endTail = trim(Request("endTail"))
Sys_MemberID = trim(Request("MemberID"))

strTableA="(select count(1) cmt from GetBillBase where counterfoireturn=0 and GetBillMemberID in(select memberid from memberdata where chname in (select chname from memberdata where memberid="&Sys_MemberID&")))"

strSQL="select * from "&strTableA&" where cmt>1"

set rsbill=conn.execute(strSQL)
If rsbill.eof Then
	response.Write "addGetBillBase.Chk_GetBillNo.value=true;"& vbcrlf
else
	response.Write "addGetBillBase.Chk_GetBillNo.value='';"
End if
rsbill.close
%>
