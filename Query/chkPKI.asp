<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt
if trim(request("PKICarchk"))<>"" then
	strSQL="select Count(*) as cnt from MemberData where AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null and PKI='"&trim(request("PKICarchk"))&"'"
	set rscnt=conn.execute(strSQL)
	if Cint(rscnt("cnt"))>0 then%>
		if(confirm('確定要入案到監理所？')){
			if (myForm.DB_Selt.value==""){
				alert("請先查詢欲入案的舉發單！");
			}else{
				myForm.kinds.value="BillToDCILog";
				myForm.submit();
			}
		}
	<%
	else
		response.write "alert('無此自然人憑證!!');"
	end if
else
	response.write "alert('無此自然人憑證!!');"
end if
rscnt.close
conn.close
%>
