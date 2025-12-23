<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

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
	strwhere=" sn in(select BillSN from DCILog where ExchangeTypeID='A' and BatchNumber in('"&Sys_BatchNumber&"'))"
end if

if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
	Sys_BillNo1=right("00000000000000000"&trim(request("Sys_BillNo1")),16)
	Sys_BillNo2=right("00000000000000000"&trim(request("Sys_BillNo2")),16)

	strwhere=strwhere&" and ImageFileNameB between '"&trim(Sys_BillNo1)&"' and '"&trim(Sys_BillNo2)&"'"

elseif trim(request("Sys_BillNo1"))<>"" then
	Sys_BillNo1=right("00000000000000000"&trim(request("Sys_BillNo1")),16)

	strwhere=strwhere&" and ImageFileNameB between '"&trim(Sys_BillNo1)&"' and '"&trim(Sys_BillNo1)&"'"

elseif trim(request("Sys_BillNo2"))<>"" then
	Sys_BillNo2=right("00000000000000000"&trim(request("Sys_BillNo2")),16)

	strwhere=strwhere&" and ImageFileNameB between '"&trim(Sys_BillNo2)&"' and '"&trim(Sys_BillNo2)&"'"

end if
Sys_MailDate=request("MailDate")

strSQL="select distinct sn from billbase where "&strwhere
set rs=conn.execute(strSQL)
while Not rs.eof
	strSQL="Update StopBillMailHistory set MailDate="&funGetDate(gOutDT(Sys_MailDate),0)&" where BillSn="&trim(rs("sn"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close
strLog=""
If trim(sys_City)="台中市" Then
	If Sys_BatchNumber<>"" Then
		strLog="批號:"&replace(Sys_BatchNumber,"'","")&" , 郵寄日期:"&request("MailDate")
	elseif trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
		strLog="單號:"&trim(request("Sys_BillNo1"))&"~"&trim(request("Sys_BillNo2"))&" , 郵寄日期:"&request("MailDate")
	End if
	If Not ifnull(strLog) Then ConnExecute strLog,978
end if
%>