<!--#include virtual="traffic/Common/DB.ini"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(UCase(request("Sys_BillNo1")))<>"" then
	BillStartNumber=trim(request("Sys_BillNo1")):BillEndNumber=trim(request("Sys_BillNo2"))
	if trim(BillEndNumber)="" then BillEndNumber=BillStartNumber

	if trim(BillStartNumber)<>"" then
		for i=1 to len(BillStartNumber)
			if IsNumeric(mid(BillStartNumber,i,1)) then
				Sno=MID(BillStartNumber,1,i-1)
				Tno=MID(BillStartNumber,i,len(BillStartNumber))
				exit for
			end if
		next
	end if
	if trim(BillEndNumber)<>"" then
		for i=1 to len(BillEndNumber)
			if IsNumeric(mid(BillEndNumber,i,1)) then
				Sno2=MID(BillEndNumber,1,i-1)
				Tno2=MID(BillEndNumber,i,len(BillEndNumber))
				exit for
			end if
		next
	end if
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,c.UserMarkDate from DCILog a,BillBase b,billmailhistory c where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and a.DciReturnStatusID not in('n','k') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.billsn=c.billsn and a.billno=c.billno and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by c.UserMarkDate"
	else
		strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate from DCILog a,BillBase b where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and a.BillSN=b.SN and a.BillNo=b.BillNo and a.DciReturnStatusID not in('N') and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by b.RecordDate"
		
		if sys_City="基隆市" then
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate,c.OwnerCounty from DCILog a,BillBase b,BillBaseDciReturn c where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and a.BillSN=b.SN and a.BillNo=b.BillNo and b.RecordStateID=0 and a.DciReturnStatusID not in('N') and a.BillNO=c.BillNo and a.CarNo=c.CarNO and c.ExchangeTypeID='W' and NVL(b.EquiPmentID,1)<>-1 order by c.OwnerCounty,b.RecordDate"
		end if
	end if
elseif trim(UCase(request("Sys_BatchNumber")))<>"" then
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
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,c.UserMarkDate from DCILog a,BillBase b,billmailhistory c where a.BatchNumber in('"&Sys_BatchNumber&"') and a.DciReturnStatusID not in('n','k') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.billsn=c.billsn and a.billno=c.billno and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by c.UserMarkDate"
	else
		strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate from DCILog a,BillBase b where a.BatchNumber in('"&Sys_BatchNumber&"') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.DciReturnStatusID not in('N') and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by b.RecordDate"

		if sys_City="基隆市" then
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate,c.OwnerCounty from DCILog a,BillBase b,BillBaseDciReturn c where a.BatchNumber in('"&Sys_BatchNumber&"') and a.BillSN=b.SN and a.BillNo=b.BillNo and b.RecordStateID=0 and a.DciReturnStatusID not in('N') and a.BillNO=c.BillNo and a.CarNo=c.CarNO and c.ExchangeTypeID='W' and NVL(b.EquiPmentID,1)<>-1 order by c.OwnerCounty,b.RecordDate"
		end if
	end if
end if
set rsload=conn.execute(strSQL)
cnt=0:chkcnt=cdbl(request("chkcnt")-1)
while Not rsload.eof
	If chkcnt<cdbl(cnt) Then
		chkcnt=chkcnt+30
		Response.Write "insertRow(fmyTable);"
	end if
	Response.Write "AddForm.item["&cnt&"].value='"&trim(rsload("BillNo"))&"';"
	cnt=cnt+1
	rsload.movenext
wend
rsload.close
conn.close
%>
