<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strSQL = "select BillTypeID,Driver,Owner from BillBaseView where BillNo='"&trim(request("BillNo"))&"'"
set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	if trim(rscnt("BillTypeID"))="1" then
		response.write "myForm.Sys_Arguer.value='"&trim(rscnt("Driver"))&"';"
	else
		response.write "myForm.Sys_Arguer.value='"&trim(rscnt("Owner"))&"';"
	end if
end if
rscnt.close
conn.close
%>
