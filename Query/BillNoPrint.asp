<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select a.BillSN,a.BillNo,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData<>'T' "&request("SQLstr")

If instr(request("Sys_BatchNumber"),"WT")>0 Then strSQL=strSQL&" and f.Note like '2%'"

strSQL=strSQL&" order by f.RecordDate"
if request("chk_MailNumKind")=1 then
	strSQL="select a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData<>'T' "&request("SQLstr")

	strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
end if
set rsfound=conn.execute(strSQL)
'-----------------------------------   保防標籤使用   -----寫入 maildate, mailnumber or storeandsendmailnumber--------------------------------------------------------
if trim(request("printStyle"))="99" then
	if sys_City="台中市" then
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			tmpSql="select BillSN,StoreAndSendMailNumber as MailNumber from BillMailHistory where BILLSN="
			tmpUSql="Update BillMailHistory set StoreAndSendMailNumber=MailNumber_Sn.NextVal where BillSN="
		end if
	else
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			tmpSql="select BillSN,StoreAndSendMailNumber as MailNumber from BillMailHistory where BILLSN="
			tmpUSql="Update BillMailHistory set StoreAndSendMailNumber=MailNumber_Sn.NextVal where BillSN="
		else
			tmpSql="select BillSN,MailNumber from BillMailHistory where BILLSN="
			tmpUSql="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=MailNumber_Sn.NextVal where BillSN="
		end if
	end if
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新
		strSQL=tmpSql&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			If ifnull(trim(rscnt("MailNumber"))) Then
				strSQL=tmpUSql&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	wend
end if
'----------------------------------------------------使用-----------寫入 maildate, mailnumber or storeandsendmailnumber--------------------------------------------------------
'77 雲林縣98新郵簡舉發單
'79 屏東縣98新郵簡舉發單

if instr(",0,18,19,25,26,33,34,53,62,77,79,179,270,",","&trim(request("printStyle"))&",") >0 then
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=MailNumber_Sn.NextVal where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	Wend
elseif trim(request("printStyle"))="81" then
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailDate,substr(BillNo,1,2) chkNo,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) and instr("BB,BC,BD",trim(rscnt("chkNo")))>0 then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=MailNumber_Sn.NextVal where BillSN="&trim(rsfound("BillSN"))

				conn.execute(strSQL)

			elseIf ifnull(rscnt("MailDate")) Then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&" where BillSN="&trim(rsfound("BillSN"))

				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	Wend
elseif trim(request("printStyle"))="24" then
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=TCMailNumber.NextVal where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	Wend
elseif sys_City="花蓮縣" and (not ifnull(trim(rsfound("BillNo")))) and (trim(Session("Unit_ID"))="Z000" or trim(Session("Unit_ID"))="A000") then
	while Not rsfound.eof
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber='"&trim(right("0000000"&right(trim(rsfound("BillNo")),6),6))&"' where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)

				sys_chkMailNumber=trim(right("0000000"&right(trim(rsfound("BillNo")),6),6))&" 970007 17"
				strSQL="Update BillMailHistory set MailChkNumber='"&sys_chkMailNumber&"' where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	wend
else
	while Not rsfound.eof
		strSQL="select MailDate from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rs=conn.execute(strSQL)
		If Not rs.eof Then
			If ifnull(rs("MailDate")) Then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&" where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
				
			End if
		end if
		rs.close
		rsfound.movenext
	wend
end if
rsfound.close
strLog=""
If trim(sys_City)="台中市" Then
	If Not Ifnull(request("Sys_BatchNumber")) Then
		strLog="批號:"&trim(request("Sys_BatchNumber"))&" , 郵寄日期:"&date
	End if
	If Not ifnull(strLog) Then ConnExecute strLog,978
end if
%>