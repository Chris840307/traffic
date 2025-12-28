<!--#include virtual="traffic/Common/DB.ini"-->
<%

strSQL = "select Count(*) as cnt from PasserSend where BillSN="&request("BillSN")&" and BillNo='"&request("BillNo")&"'"
set rscnt=conn.execute(strSQL)
if Cint(rscnt("cnt"))>0 then%>
	myForm.DB_Add.value="Update";
	myForm.submit();
<%else%>
	myForm.DB_Add.value="ADD";
	myForm.submit();
<%end if
rscnt.close
conn.close
%>
