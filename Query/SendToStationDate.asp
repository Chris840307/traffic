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
	strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
end if

if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
elseif trim(request("Sys_BillNo1"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo1"))&"'"
elseif trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo2"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
end if
strSQL="select distinct a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&strwhere

set rs=conn.execute(strSQL)

Sys_SendOpenDate=request("SendOpenDate")
while Not rs.eof
	strSQL="Update BillMailHistory set SendOpenGovDocToStationDate="&right("00"&Sys_SendOpenDate,7)&" where BillSn="&trim(rs("BillSN"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close
%>