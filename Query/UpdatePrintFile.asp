<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Sys_BatchNumber=trim(Request("BatchNumber"))
If Not ifnull(Request("ReBillno")) Then
	Sys_ReBillNo=split(trim(Request("ReBillno")),"~")

	strSQL="Update BillPrintJob set PrintMemberID="&session("User_ID")&",PrintDateTime=sysdate,PrintStatus=1 where BatchNumber='"&Sys_BatchNumber&"' and BillNo='"&trim(Request("ReBillno"))&"'"

	conn.execute(strSQL)
	For j = 0 to Ubound(Sys_ReBillNo)
		Response.Write "window.opener.myForm.Sys_BillNo"&(j+1)&".value='"&Sys_ReBillNo(j)&"';"& vbcrlf
	Next
else
	strSQL="Update BillPrintJob set PrintMemberID="&session("User_ID")&",PrintDateTime=sysdate,PrintStatus=1 where BatchNumber='"&Sys_BatchNumber&"' and BillNo is null"

	conn.execute(strSQL)

	Response.Write "window.opener.myForm.Sys_BillNo1.value='';"& vbcrlf
	Response.Write "window.opener.myForm.Sys_BillNo2.value='';"& vbcrlf
End if

%>
window.opener.myForm.Sys_BatchNumber.value='<%=trim(request("batchnumber"))%>';
window.opener.funSelt('BatchSelt');
window.close();
