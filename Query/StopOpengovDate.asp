<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=60000

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strwhere=trim(request("SQLstr"))

opengovDate=gOutDT(request("Sys_opengovDate"))

strQuery="Update StopBillMailHistory set OpenGovDate="&funGetDate(opengovDate,0)& " where billsn in(select distinct a.sn from (select sn,imagefilenameb from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b,(select * from StopBillMailHistory where UserMarkResonID in('1','2','3','4','8','M','K','L','O','P','Q')) c where a.SN=b.BillSN and a.SN=c.BillSn and a.imagefilenameb=c.BillNo "&strwhere&")"

conn.execute(strQuery)

strSQL="update billbase set DealLineDate="&funGetDate(DateAdd("d",27,opengovDate),0)&" where sn in(select distinct a.sn from (select sn,imagefilenameb from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b,(select * from StopBillMailHistory where UserMarkResonID in('1','2','3','4','8','M','K','L','O','P','Q')) c where a.SN=b.BillSN and a.SN=c.BillSn and a.imagefilenameb=c.BillNo "&strwhere&")"

conn.execute(strSQL)

%>
<script language="JavaScript">
	alert ("Àx¦s§¹¦¨!!");
	opener.myForm.submit(); 
	self.close();
</script>