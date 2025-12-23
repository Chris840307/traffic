<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and BillStatus=2 and RecordStateId <> -1 and ImageFileNameB is null and DeallineDate is not null) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&request("SQLstr")&" order by a.CarNo,a.IllegalDate"
set rsfound=conn.execute(strSQL)
tempCarNo="":FileCount=1:StopCarNo=""
while Not rsfound.eof
	If trim(rsfound("CarNo"))<>trim(tempCarNo) Then
		FileCount=1
		tempCarNo=trim(rsfound("CarNo"))
		StopCarNo="StopCarNo.NextVal"
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
		strSQL="Insert Into StopBillMailHistory(BillSN,BillNo,MailDate) values("&trim(rsfound("SN"))&",LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&")"
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
funStopBillPrints_HuaLien();
