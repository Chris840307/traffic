<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
BillStartNumber = trim(Request("sn"))

strSQL="select distinct a.billendnumber,a.getbillmemberid,b.loginID from getbillbase a,Memberdata b where a.RecordstateID=0 and a.BillIn=0 and a.BillStartNumber='"&BillStartNumber&"' and a.getbillmemberid=b.MemberID"

set rsbill=conn.execute(strSQL)
If Not rsbill.eof Then%>
	myForm.BillEndNumber.value='<%=trim(rsbill("billendnumber"))%>';
	myForm.chekChMemID.value='<%=trim(rsbill("loginID"))%>';
	myForm.chekChMemID.focus();
<%
End if
rsbill.close
%>