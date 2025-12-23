<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strSQL = "select d.DCIRETURNSTATION,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo from BillBase a,Station b,UnitInfo c,BILLBASEDCIRETURN d where a.RecordStateID=0 and a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIRETURNSTATION=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+) and a.RecordStateID=0 and d.BillNo='"&trim(Ucase(request("BillNo")))&"'"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	var err=0,errmsg='';
	AddForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
	AddForm.mailNumber[myForm.chkcnt.value-1].focus();
	//AddForm.Sys_BackDate[myForm.chkcnt.value-1].focus();
<%
else
	response.write "alert('?L???渹!!');"
end if
rscnt.close
conn.close
%>
