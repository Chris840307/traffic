<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If trim(session("Unit_ID"))="08A7" Then
	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData<>'T' "&request("SQLstr")&" order by f.RecordDate"
	set rsfound=conn.execute(strSQL)
	while Not rsfound.eof
		strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&trim(rsfound("BillSN"))
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
				sys_chkMailNumber=""
				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber=LPAD(MailNumber_Sn.NextVal,6,'0') where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
				strSQL="select MailNumber from BillMailHistory where BillSN='"&trim(rsfound("BillSN"))&"'"
				set rsnumber=conn.execute(strSQL)
				If Not rsnumber.eof Then
					sys_chkMailNumber=trim(right("0000000"&rsnumber("MailNumber"),6))&" 830009 17"
				End if
				strSQL="Update BillMailHistory set MailChkNumber='"&sys_chkMailNumber&"' where BillSN="&trim(rsfound("BillSN"))
				conn.execute(strSQL)
			end if
		end if
		rscnt.close
		rsfound.movenext
	wend
	rsfound.close
end if
%>