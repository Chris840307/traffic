<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


strSQL="select count(1) cnt from getbillbase where GetBillMemberID="&Request("GetBillMemberID_q")&" and CounterfoiReturn=0"

set rsbill=conn.execute(strSQL)

If cdbl(rsbill("cnt")) > 0 Then
	Response.Write "if(confirm(""該人員尚有紅單未繳回，是否繼續領單?"")){"
	Response.Write "UrlStr=""GetBill_mdy.asp"";"
	Response.Write "addGetBillBase.action=UrlStr;"
	Response.Write "addGetBillBase.target=""AddGetBill"";"
	Response.Write "addGetBillBase.submit();"
	Response.Write "addGetBillBase.action="";"
	Response.Write "addGetBillBase.target="";"
	Response.Write "}"
else
	Response.Write "UrlStr=""GetBill_mdy.asp"";"
	Response.Write "addGetBillBase.action=UrlStr;"
	Response.Write "addGetBillBase.target=""AddGetBill"";"
	Response.Write "addGetBillBase.submit();"
	Response.Write "addGetBillBase.action="";"
	Response.Write "addGetBillBase.target="";"
End if
rsbill.close
%>