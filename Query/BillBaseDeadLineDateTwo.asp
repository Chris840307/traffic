<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=60000

sys_maiNumberSN=""

If sys_City="台東縣" Then
	sys_maiNumberSN="MailNumber_Stop_Sn"

elseif sys_City="花蓮縣" then
	sys_maiNumberSN="stopcarbillprintno"
end if

PBillNo=split(trim(Request("PBillNo")),",")
PCarNo=split(trim(Request("PCarNo")),",")
Sys_DeallineDate=request("Sys_DeallineDate")
For i = 0 to Ubound(PBillNo)
	StoreAndSendMailNumber=""
	strSQL="Update BillBase set DeallineDate="&funGetDate(gOutDT(Sys_DeallineDate),0)&" where ImageFileNameB='"&trim(PBillNo(i))&"'"

	conn.execute(strSQL)
	
	strSQL="select BillSN from StopBillMailHistory where BillNo='"&trim(PBillNo(i))&"' and StoreAndSendMailNumber is null"
	set rsck=conn.execute(strSQL)
	If Not rsck.eof Then
		strSQL="select LPAD("& sys_maiNumberSN &".NextVal,6,'0') cmt from Dual"
		set rsnum=conn.execute(strSQL)
		StoreAndSendMailNumber=trim(rsnum("cmt"))
		rsnum.close

		strSQL="Update StopBillMailHistory set StoreAndSendMailNumber='"&StoreAndSendMailNumber&"' where BillNo='"&trim(PBillNo(i))&"' and StoreAndSendMailNumber is null"
		conn.execute(strSQL)
	End if
	rsck.close

	strSQL="Update StopBillMailHistory set MailDate="&funGetDate(date,0)&" where BillNo='"&trim(PBillNo(i))&"'"
	conn.execute(strSQL)
next
%>
<script language="JavaScript">
	alert ("儲存完成!!");
	self.close();
</script>