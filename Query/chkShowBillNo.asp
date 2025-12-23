<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
tmp_BatchNumber=UCase(replace(request("Sys_BatchNumber"),",","','"))

strSQL = "select min(BillNo) BillNoA,max(BillNo) BillNoB from DCILog where BillTypeID='2' and ExchangeTypeID in('W','N') and BatchNumber in('"&tmp_BatchNumber&"')"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	If Not ifnull(rscnt("BillNoA")) Then%>
		showBillNoA.innerHTML="起始號：<%=rscnt("BillNoA")%>";
		showBillNoB.innerHTML="結束號：<%=rscnt("BillNoB")%>";
<%	end if
else%>
	showBillNoA.innerHTML="";
	showBillNoB.innerHTML="";
<%end if
rscnt.close
conn.close
%>
