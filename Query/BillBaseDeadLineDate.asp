<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=60000

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

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
Sys_BillNo1=right("000000000000000000000"&trim(request("Sys_ImageFileNameB1")),16)
Sys_BillNo2=right("000000000000000000000"&trim(request("Sys_ImageFileNameB2")),16)

if trim(request("Sys_ImageFileNameB1"))<>"" and trim(request("Sys_ImageFileNameB2"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(Sys_BillNo1))&"' and '"&trim(UCase(Sys_BillNo2))&"'"
elseif trim(request("Sys_ImageFileNameB1"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(Sys_BillNo1))&"' and '"&trim(UCase(Sys_BillNo1))&"'"
elseif trim(request("Sys_ImageFileNameB2"))<>"" then
	strwhere=strwhere&" and a.ImageFileNameB between '"&trim(UCase(Sys_BillNo2))&"' and '"&trim(UCase(Sys_BillNo2))&"'"
end if
strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where BillNo is null and RecordStateId <> -1) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"

set rs=conn.execute(strSQL)

Sys_DeallineDate=request("Sys_DeallineDate")
while Not rs.eof
	'strSQL="select "
	strSQL="Update BillBase set DeallineDate="&funGetDate(gOutDT(Sys_DeallineDate),0)&",BillStatus=2 where Sn="&trim(rs("SN"))
	conn.execute(strSQL)
	rs.movenext
wend
rs.close
sys_maiNumberSN=""

If sys_City="台東縣" Then
sys_maiNumberSN="MailNumber_Stop_Sn"

elseif sys_City="花蓮縣" then
sys_maiNumberSN="stopcarbillprintno"

End if


strSQL="select distinct a.SN,a.CarNo,a.IllegalDate from (select * from BillBase where ImagePathName is not null and BillStatus=2 and RecordStateId <> -1 and ImageFileNameB is null and DeallineDate is not null) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.CarNo,a.IllegalDate"

set rsfound=conn.execute(strSQL)
tempCarNo="":FileCount=1:StopCarNo="":StopMailNumberA="":StopMailNumberB=""
while Not rsfound.eof
	SendAddrFlag=1
'	strSQL="select SendAddrFlag from StopCaseSendAddr where CarNo=(select CarNo from BillBase where SN="&trim(rsfound("SN"))&")"
'	set rsSend=conn.execute(strSQL)
'	If Not rsSend.eof Then SendAddrFlag=trim(rsSend("SendAddrFlag"))
'	rsSend.Close

	If trim(rsfound("CarNo"))<>trim(tempCarNo) Then
		FileCount=1
		tempCarNo=trim(rsfound("CarNo"))
		StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
		StopMailNumberA="LPAD("& sys_maiNumberSN &".NextVal,6,'0')"
		StopMailNumberB="LPAD("& sys_maiNumberSN &".NextVal,6,'0')"
	else
		FileCount=FileCount+1
		If FileCount>8 Then
			FileCount=1
			StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
			StopMailNumberA="LPAD("& sys_maiNumberSN &".NextVal,6,'0')"
			StopMailNumberB="LPAD("& sys_maiNumberSN &".NextVal,6,'0')"
		else
			StopCarNo="LPAD(StopCarNo.CurrVal,16,'0')"		
			StopMailNumberA="LPAD("& sys_maiNumberSN &".CurrVal,6,'0')"
			'If SendAddrFlag = 3 Then StopMailNumberA="LPAD((stopcarbillprintno.CurrVal-1),6,'0')"
			StopMailNumberB="LPAD("& sys_maiNumberSN &".CurrVal,6,'0')"
		End if
	End if
	strSQL="Update BillBase set ImageFileNameB="&StopCarNo&" where SN="&trim(rsfound("SN"))
	conn.execute(strSQL)

	strSQL="select BillSN from StopBillMailHistory where BillSN="&trim(rsfound("SN"))
	set rsStop=conn.execute(strSQL)
	If rsStop.eof Then
		If SendAddrFlag=1 Then
			strSQL="Insert Into StopBillMailHistory(BillSN,CarNo,BillNo,MailDate,MailNumber) values("&trim(rsfound("SN"))&",'"&trim(rsfound("CarNo"))&"',LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&","&StopMailNumberA&")"
			conn.execute(strSQL)
		elseif SendAddrFlag=2 then
			strSQL="Insert Into StopBillMailHistory(BillSN,CarNo,BillNo,MailDate,StoreAndSendMailNumber) values("&trim(rsfound("SN"))&",'"&trim(rsfound("CarNo"))&"',LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&","&StopMailNumberB&")"
			conn.execute(strSQL)
		elseif SendAddrFlag=3 then
			strSQL="Insert Into StopBillMailHistory(BillSN,CarNo,BillNo,MailDate,MailNumber) values("&trim(rsfound("SN"))&",'"&trim(rsfound("CarNo"))&"',LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&","&StopMailNumberA&")"
			conn.execute(strSQL)

			strSQL="Update StopBillMailHistory set StoreAndSendMailNumber="&StopMailNumberB&" where BillSN="&trim(rsfound("SN"))
			conn.execute(strSQL)
		End if
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
	alert ("儲存完成!!");
	opener.myForm.submit(); 
	self.close();
</script>