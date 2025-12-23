<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
if UCase(request("Sys_BatchNumber"))<>"" then
	tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
	next
	strwhere=" and BatchNumber in('"&Sys_BatchNumber&"')"
end if

if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
elseif trim(request("Sys_BillNo1"))<>"" then
	strwhere=strwhere&" and BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo1"))&"'"
elseif trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and BillNo between '"&trim(request("Sys_BillNo2"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
end if
strSQL="select distinct BillSN from DCILog where DciReturnStatusID<>'n'"&strwhere

set rs=conn.execute(strSQL)

Sys_SendOpenDate=request("StoreAndSendMailDate")
while Not rs.eof
	strSQL="Update BillMailHistory set StoreANDSendSendDate="&funGetDate(gOutDT(Sys_SendOpenDate),0)&" where BillSn="&trim(rs("BillSN"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close
%>