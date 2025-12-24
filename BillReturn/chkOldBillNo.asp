<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


strSQL = "select BillNo from Billbase where BillNo='"&trim(Ucase(request("BillNo")))&"'"
set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	response.write "alert("" !! 此單號為新系統案件，請勿在此進行送達註記 !! "");"
else
	response.write "myForm.CarNo[myForm.chkcnt.value-1].focus();"
end if
rscnt.close
conn.close
%>
