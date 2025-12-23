<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=60000
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
strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"

set rs=conn.execute(strSQL)

Sys_DeallineDate=request("Sys_DeallineDate")
while Not rs.eof
	'strSQL="select "
	strSQL="Update BillBase set DeallineDate="&funGetDate(gOutDT(Sys_DeallineDate),0)&",BillStatus=2 where Sn="&trim(rs("SN"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close

strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and BillStatus=2 and RecordStateId <> -1 and ImageFileNameB is null and DeallineDate is not null) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"
set rsfound=conn.execute(strSQL)
tempCarNo="":FileCount=1:StopCarNo=""
while Not rsfound.eof
	If trim(rsfound("CarNo"))<>trim(tempCarNo) Then
		FileCount=1
		tempCarNo=trim(rsfound("CarNo"))
		StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
	else
		FileCount=FileCount+1
		If FileCount>8 Then
			FileCount=1
			StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
		else
			StopCarNo="LPAD(StopCarNo.CurrVal,16,'0')"
		End if
	End if
	strSQL="Update BillBase set ImageFileNameB="&StopCarNo&" where SN="&trim(rsfound("SN"))
	conn.execute(strSQL)
	strSQL="select BillSN from StopBillMailHistory where BillSN="&trim(rsfound("SN"))
	set rsStop=conn.execute(strSQL)
	If rsStop.eof Then
		strSQL="Insert Into StopBillMailHistory(BillSN,BillNo,MailDate,MailNumber) values("&trim(rsfound("SN"))&",LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&",SubStr(LPAD(StopCarNo.CurrVal,16,'0'),11))"
		conn.execute(strSQL)
	else
		strSQL="Update StopBillMailHistory set MailDate="&funGetDate(date,0)&" where BillSN="&trim(rsfound("SN"))
		conn.execute(strSQL)
	End if
	rsStop.close
	rsfound.movenext
wend
rsfound.close
%>
<script language="JavaScript">
	alert ("Àx¦s§¹¦¨!!");
	opener.myForm.submit(); 
	self.close();
</script>