<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

tb_BillBase="select Sn,BillNo,CarNo,BillStatus,BillUnitID from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0"

tb_BillbaseDCIReturn="select BillNo,CarNo,Owner,DCIReturnStation from BillbaseDCIReturn where ExChangeTypeID='W' and BillNo='"&trim(Ucase(request("BillNo")))&"'"

tb_BillMailHistory="select BillSn,BillNo,MailReturnDate from BillMailHistory where BillNo='"&trim(Ucase(request("BillNo")))&"'"


strSQL = "select a.BillStatus,d.DCIReturnStation,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo,d.Owner,e.MailReturnDate from ("&tb_BillBase&") a,Station b,UnitInfo c,("&tb_BillbaseDCIReturn&") d,("&tb_BillMailHistory&") e where a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIReturnStation=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+) and a.BillNo=e.BillNo(+) and a.Sn=e.BillSn(+)"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	var City='<%=sys_City%>';
	if(City=='花蓮縣'){
		myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("Owner"))%>';
		
		var billstatis='<%=trim(rscnt("BillStatus"))%>';
		//檢查是否沒有做過寄存送達
		if (billstatis != '4'){
			alert('沒有做過寄存送達!!');
		}
		
	}else{
		myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
	}
	if(City=='彰化縣'){
		myForm.EffectDate[myForm.chkcnt.value-1].value='<%=trim(gInitDT(rscnt("MailReturnDate")))%>';
	}
	if(myForm.chkcnt.value<myForm.item.length){
		myForm.item[myForm.chkcnt.value].focus();
	}
<%
else
	response.write "alert('無此單號!!');"
end if
rscnt.close
conn.close
%>
