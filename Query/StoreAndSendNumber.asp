<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=60000

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sDay1="To_Date('" & gOutDT(request("Day1"))&" 00:00:00','YYYY/MM/DD/HH24/MI/SS')"
sDay2="To_Date('" & gOutDT(request("Day2"))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

strSQL="update MailStationReturn set StoreAndSendNumber='"&Trim(Request("SeqNo"))&"' where UserMarkDate between "& sDay1 &" and "& sDay2

conn.execute(strSQL)
%>
<script language="JavaScript">
	alert ("Àx¦s§¹¦¨!!");
	opener.myForm.submit(); 
	self.close();
</script>