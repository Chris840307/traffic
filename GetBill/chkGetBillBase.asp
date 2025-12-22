<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
startHead = trim(Request("startHead"))
startTail = trim(Request("startTail"))
endHead = trim(Request("endHead"))
endTail = trim(Request("endTail"))
Sys_UnitID = trim(Request("UnitID"))

strTableA="(select * from GetBillBase where BILLIN=1 and SUBSTR(BillStartNumber,1,"&len(startHead)&")='"&startHead&"' and SUBSTR(BillStartNumber,"&len(startHead)+1&")<='"&startTail&"' and SUBSTR(BillEndNumber,1,"&len(endHead)&")='"&endHead&"' and SUBSTR(BillEndNumber,"&len(endHead)+1&")>='"&endTail&"')"

strSQL="select * from "&strTableA

set rsbill=conn.execute(strSQL)
If rsbill.eof Then
	response.Write "addGetBillBase.Chk_GetBillNo.value='';"& vbcrlf
else
	response.Write "addGetBillBase.Chk_GetBillNo.value=true;"
End if
rsbill.close
%>
