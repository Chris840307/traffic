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

		Ctrl=0:MailNumberKind=""
		ONfadd="":owradd="":Drhadd="":ONfaddID="":owraddID="":DrhaddID=""

		strSQL="select a.OwnerAddress,a.OwnerNotifyAddress,a.DriverHomeAddress,b.OwnerHomeAddress,b.DriverAddress from (select CarNo,nvl(OwnerAddress,' ') OwnerAddress,nvl(OwnerNotifyAddress,' ')OwnerNotifyAddress,nvl(DriverHomeAddress,' ') DriverHomeAddress from BillbaseDCIReturn where CarNo='"&trim(rsfound("CarNo"))&"' and ExchangetypeID='A') a,(select CarNo,OwnerAddress OwnerHomeAddress,DriverAddress from BillBase where sn="&trim(rsfound("SN"))&") b where a.CarNo=b.CarNo"

		set rsas=conn.execute(strSQL)

		If not ifnull(rsas("OwnerNotifyAddress")) Then
			ONfadd=mid(trim(rsas("OwnerNotifyAddress")),4)

		end if

		If not ifnull(rsas("OwnerHomeAddress")) Then
			owradd=trim(rsas("OwnerHomeAddress"))

		elseIf not ifnull(rsas("OwnerAddress")) Then
			owradd=trim(rsas("OwnerAddress"))

		End if

		If not ifnull(rsas("DriverAddress")) Then
			Drhadd=trim(rsas("DriverAddress"))

		elseIf not ifnull(rsas("DriverHomeAddress")) Then
			Drhadd=trim(rsas("DriverHomeAddress"))

		End if
		rsas.close

		If (ONfadd = owradd) and (ONfadd = Drhadd) and (owradd = Drhadd) Then
			If not ifnull(ONfadd) Then
				Ctrl=1
				MailNumberKind="MailNumber"
			End if

		elseIf (ONfadd <> owradd) and (ONfadd = Drhadd) and (owradd <> Drhadd) Then
			If (not ifnull(ONfadd)) and (not ifnull(owradd)) Then
				Ctrl=2
				MailNumberKind="MailNumber,StoreAndSendMailNumber"
			elseIf not ifnull(ONfadd) Then
				Ctrl=1
				MailNumberKind="MailNumber"
			elseIf not ifnull(owradd) Then
				Ctrl=1
				MailNumberKind="StoreAndSendMailNumber"
			End if

		elseIf (ONfadd = owradd) and (ONfadd <> Drhadd) and (owradd <> Drhadd) Then
			If (not ifnull(ONfadd)) and (not ifnull(Drhadd)) Then
				Ctrl=2
				MailNumberKind="MailNumber,DriverMailNumber"
			elseIf not ifnull(ONfadd) Then
				Ctrl=1
				MailNumberKind="MailNumber"
			elseIf not ifnull(Drhadd) Then
				Ctrl=1
				MailNumberKind="DriverMailNumber"
			End if

		elseIf (ONfadd <> owradd) and (ONfadd <> Drhadd) and (owradd = Drhadd) Then
			If (not ifnull(ONfadd)) and (not ifnull(owradd)) Then
				Ctrl=2
				MailNumberKind="MailNumber,StoreAndSendMailNumber"
			elseIf not ifnull(ONfadd) Then
				Ctrl=1
				MailNumberKind="MailNumber"
			elseIf not ifnull(owradd) Then
				Ctrl=1
				MailNumberKind="StoreAndSendMailNumber"
			End if

		elseIf (ONfadd <> owradd) and (ONfadd <> Drhadd) and (owradd <> Drhadd) Then
			If not ifnull(ONfadd) Then
				Ctrl=Ctrl+1
				MailNumberKind="MailNumber"
			end if

			If not ifnull(owradd) Then
				Ctrl=Ctrl+1
				If not ifnull(MailNumberKind) Then
					MailNumberKind=MailNumberKind&",StoreAndSendMailNumber"
				else
					MailNumberKind="StoreAndSendMailNumber"
				End if
			end if

			If not ifnull(Drhadd) Then
				Ctrl=Ctrl+1
				If not ifnull(Drhadd) Then
					MailNumberKind=MailNumberKind&",DriverMailNumber"
				else
					MailNumberKind="DriverMailNumber"
				End if
			end if

		End if

		MailNo=split(",,",",")
		For i = 1 to Ctrl
			strSQL="select LPAD(stopcarbillprintno.NextVal,6,'0') cmt from Dual"
			set rsnum=conn.execute(strSQL)
			MailNo(i-1)=trim(rsnum("cmt"))
			rsnum.close
		Next

		FileCount=1
		tempCarNo=trim(rsfound("CarNo"))
		StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
	else
		FileCount=FileCount+1
		If FileCount>8 Then
			
			MailNo=split(",,",",")
			For i = 1 to Ctrl
				strSQL="select LPAD(stopcarbillprintno.NextVal,6,'0') cmt from Dual"
				set rsnum=conn.execute(strSQL)
				MailNo(i-1)=trim(rsnum("cmt"))
				rsnum.close
			Next

			FileCount=1
			StopCarNo="LPAD(StopCarNo.NextVal,16,'0')"
		else
			FileCount=FileCount+1
			StopCarNo="LPAD(StopCarNo.CurrVal,16,'0')"		
		End if
	End if
	strSQL="Update BillBase set ImageFileNameB="&StopCarNo&" where SN="&trim(rsfound("SN"))
	conn.execute(strSQL)

	strSQL="select BillSN from StopBillMailHistory where BillSN="&trim(rsfound("SN"))
	set rsStop=conn.execute(strSQL)
	If rsStop.eof Then
		tmpMailNo=""
		For i = 1 to Ctrl
			If not ifnull(MailNo(i-1)) Then			
				If not ifnull(tmpMailNo) Then tmpMailNo=tmpMailNo&"','"
				tmpMailNo=tmpMailNo&MailNo(i-1)
			end if
		Next
		
		strSQL="Insert Into StopBillMailHistory(BillSN,CarNo,BillNo,MailDate,"&MailNumberKind&") values("&trim(rsfound("SN"))&",'"&trim(rsfound("CarNo"))&"',LPAD(StopCarNo.CurrVal,16,'0'),"&funGetDate(date,0)&",'"&tmpMailNo&"')"
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