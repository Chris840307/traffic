<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

tmp_BatchNumber=UCase(replace(request("Sys_BatchNumber"),",","','"))


strSQL = "update billbase set rule4='"&request("Sys_Rule4")&"' where sn in(select billsn from DCILog where ExchangeTypeID='W' and BatchNumber in('"&tmp_BatchNumber&"')) and recordstateid=0"

conn.execute(strSQL)
%>
	alert("已更新完成!!");
<%
conn.close
%>
