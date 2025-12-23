<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
if UCase(request("Sys_BatchNumber"))<>"" then
	tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(tmp_BatchNumber(i))
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(tmp_BatchNumber(i))
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
	next
	strwhere=" and b.BatchNumber in('"&Sys_BatchNumber&"')"
end if

if trim(request("Sys_ImageFileNameB1"))<>"" and trim(request("Sys_ImageFileNameB2"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB1")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB2")))&"'"
elseif trim(request("Sys_ImageFileNameB1"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB1")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB1")))&"'"
elseif trim(request("Sys_ImageFileNameB2"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(request("Sys_ImageFileNameB2")))&"' and '"&trim(UCase(request("Sys_ImageFileNameB2")))&"'"
end if
strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and BillStatus=1 and RecordStateId <> -1) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"

set rs=conn.execute(strSQL)

Sys_SendOpenDate=request("SendOpenDate")
while Not rs.eof
	strSQL="Update BillMailHistory set SendOpenGovDocToStationDate="&right("00"&Sys_SendOpenDate,7)&" where BillSn="&trim(rs("SN"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close
%>