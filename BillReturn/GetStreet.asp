<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


strSQL = "select streetid,address from street where streetid='"&trim(Request("IllegalAddressID"))&"'"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	myForm.IllegalAddress[<%=Request("itemcnt")%>].value='<%=rscnt("address")%>';<%
else%>
	myForm.IllegalAddress[<%=Request("itemcnt")%>].value='';<%
end if
rscnt.close
conn.close
%>
