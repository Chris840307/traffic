<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


strSQL = "select b.BatchNumber,min(a.ImageFileNameB) BillNoA,max(a.ImageFileNameB) BillNoB from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN and b.BatchNumber='"&request("Sys_BatchNumber")&"' Group by b.BatchNumber"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	showBillNoA.innerHTML="起始號：<%=rscnt("BillNoA")%>";
	showBillNoB.innerHTML="結束號：<%=rscnt("BillNoB")%>";
<%else%>
	showBillNoA.innerHTML="";
	showBillNoB.innerHTML="";
<%end if
rscnt.close
conn.close
%>
