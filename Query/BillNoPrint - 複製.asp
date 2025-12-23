<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select a.BillSN,a.BillNo,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("SQLstr")

If instr(request("Sys_BatchNumber"),"WT")>0 Then strSQL=strSQL&" and f.Note like '2%'"

If sys_City<>"高雄市" then strSQL=strSQL&" and a.DciErrorCarData<>'T'"

if Instr(request("Sys_BatchNumber"),"N")>0 then
	If sys_City<>"基隆市" and sys_City<>"高雄縣" then
		strSQL=strSQL&" and a.DciReturnStatusID<>'n'"
	end if
end if

strSQL=strSQL&" order by f.RecordDate"
if request("chk_MailNumKind")=1 then
	strSQL="select a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("SQLstr")

	If sys_City<>"高雄市" then strSQL=strSQL&" and a.DciErrorCarData<>'T'"

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		If sys_City<>"基隆市" and sys_City<>"高雄縣" then
			strSQL=strSQL&" and a.DciReturnStatusID<>'n'"
		end if
	end if

	strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
end if
set rsfound=conn.execute(strSQL)

sys_cnt=0
strSQL="select Count(1) cnt from ("&strSQL&")"
set rscnt=conn.execute(strSQL)
sys_cnt=rscnt("cnt")
rscnt.close

'-----------------------------------   保防標籤使用   -----寫入 maildate, mailnumber or storeandsendmailnumber--------------------------------------------------------
if trim(request("printStyle"))="99" then
	if sys_City="台中市" then
		If trim(Session("UnitLevelID"))="1" Then
			if Instr(request("Sys_BatchNumber"),"N")>0 then
				tmpSql="select BillSN,StoreAndSendMailNumber as MailNumber from BillMailHistory where BILLSN="
				tmpUSql="Update BillMailHistory set StoreAndSendMailNumber=MailNumber_Sn.NextVal where BillSN="
			else
				tmpSql="select BillSN,MailNumber from BillMailHistory where BILLSN="
				tmpUSql="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=MailNumber_Sn.NextVal where BillSN="
			end if
		else
			if Instr(request("Sys_BatchNumber"),"W")>0 then
				tmpSql="select BillSN,MailNumber from BillMailHistory where MailDate is null and BILLSN="
				tmpUSql="Update BillMailHistory set MailDate="&funGetDate(date,0)&" where BillSN="
			end if
		End if
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

if instr(",18,19,25,34,77,79,",","&trim(request("printStyle"))&",") >0 then
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
elseif instr(",0,26,",","&trim(request("printStyle"))&",") >0 then '基隆
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailNumber,MailDate from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
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
elseif trim(request("printStyle"))="81" or trim(request("printStyle"))="29" or trim(request("printStyle"))="31" then
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
elseif instr(",82,83,",","&trim(request("printStyle"))&",") >0 then 

	if cdbl(sys_cnt)>20 then 
		Mail_Cnt="18"
	else
		Mail_Cnt="16"
	end if
	UnitID=Session("Unit_ID")
	if UnitID="05GF" then      '集集
		UnitNum="54000918"

	elseif UnitID="05BA" then      '南投分局
		UnitNum="54000518"

	elseif UnitID="05FG" then
		'竹山
		UnitNum="540022"&Mail_Cnt
	else
		UnitNum="54000017"
	end if

	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=NTSMAILNUMBER.NextVal||'"&UnitNum&"' where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	Wend
elseif trim(request("printStyle"))="98" then
	while Not rsfound.eof
		'smith 修改 mailnumber 只有當該紀錄mailnumber為空值才更新才更新 
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=NTSMAILNUMBER.NextVal||'54000017' where BillSN="&trim(rsfound("BillSN"))
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
elseif sys_City="花蓮縣" and (trim(Session("Unit_ID"))="Z000" or trim(Session("Unit_ID"))="A000") then
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