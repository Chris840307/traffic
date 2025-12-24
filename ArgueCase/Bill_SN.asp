<!--#include virtual="traffic/Common/DB.ini"-->
<%
strSQL="select SN from BillBaseView where BillNo='"&trim(request("BillNo"))&"'"
set rsload=conn.execute(strSQL)
if Not rsload.eof then%>
	UrlStr="../Query/ViewBillBaseData_Car.asp?BillSn=<%=rsload("SN")%>"
	window.open(UrlStr,"WebPage1","left=0,top=0,location=0,width=850,height=700,status=yes,resizable=yes,scrollbars=yes")
<%else
	response.write "alert('無此舉發單號!!');"
end if
rsload.close
%>
