<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strSQL = "select Count(*) as cnt from BillBaseView where BillNo='"&trim(request("BillNo"))&"'"
set rscnt=conn.execute(strSQL)

if Cint(rscnt("cnt"))>0 then%>
	if(myForm.Sys_BillNo.value==''){
		alert("舉發單號不可空白");
	}else if(myForm.Sys_Arguer.value==''){
		alert("陳述人姓名不可空白");

	}else{
		myForm.DB_Add.value="ADD";
		myForm.submit();
	}
<%else%>
	alert('無此舉發案件');
<%end if
rscnt.close
conn.close
%>
