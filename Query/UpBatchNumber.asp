<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Sys_BatchNumber=trim(Request("BatchNumber"))
If not ifnull(Request("ReBillno")) Then
	Sys_ReBillNo=split(trim(Request("ReBillno")),",")
	For i = 0 to Ubound(Sys_ReBillNo)
		strSQL="delete BillPrintJob where BatchNumber='"&trim(request("batchnumber"))&"' and BillNo='"&trim(Sys_ReBillNo(i))&"'"
		conn.execute(strSQL)
		billcmt=1
		tmp_ReBillNo=""

		If instr(trim(Sys_ReBillNo(i)),"~")>0 then
			tmp_ReBillNo=split(trim(Sys_ReBillNo(i)),"~")
			billcmt=Abs(cdbl(right(trim(tmp_ReBillNo(0)),4))-cdbl(right(trim(tmp_ReBillNo(1)),4)))+1
		end if
		
		strSQL="insert into BillPrintJob (BatchNumber,BillNo,PrintCnt,ActDate,RequestUnitID,RequestMemberID,PrintStatus) values('"&trim(request("batchnumber"))&"','"&trim(Sys_ReBillNo(i))&"','"&billcmt&"',sysdate,'"&session("Unit_ID")&"',"&session("User_ID")&",0)"

		conn.execute(strSQL)
	Next
else
	strSQL="delete BillPrintJob where BatchNumber='"&trim(request("batchnumber"))&"' and BillNo is null"
	conn.execute(strSQL)

	strSQL="insert into BillPrintJob (BatchNumber,PrintCnt,ActDate,RequestUnitID,RequestMemberID,PrintStatus) values('"&trim(request("batchnumber"))&"','"&trim(request("PrintCnt"))&"',sysdate,'"&session("Unit_ID")&"',"&session("User_ID")&",0)"

	conn.execute(strSQL)
End if
%>
document.all.Sys_ReBillNo.value='';
alert("上傳完成!!");
