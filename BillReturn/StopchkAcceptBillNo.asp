<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

'smith 停管修改 > 原本對billno. 現在對imagefilenameb (催繳單號) 也不用去對billbasedcireturn
strSQL = "select a.CarNo,a.DeallineDate,b.DCIStationName,b.StationID,c.UnitID,c.UnitName from BillBase a,Station b,UnitInfo c where a.BillUnitID=c.UnitID(+) and a.RecordStateID=0 and a.ImageFileNameB='"&trim(Ucase(request("BillNo")))&"'"


set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	var err=0,errmsg='',City='<%=sys_City%>';
	
	if(err=='0'){
			myForm.DeallineDate[myForm.chkcnt.value-1].value='<%=gInitDT(trim(rscnt("DeallineDate")))%>';

		if(City!='台南市'&&City!='嘉義市'&&City!='台中市'&&City!='台東縣'){
			myForm.Sys_BackDate[myForm.chkcnt.value-1].value='<%=gInitDT(DateAdd("d",getBillReturnDate,rscnt("DeallineDate")))%>';
		}
		if(City=='花蓮縣'){
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
			if(myForm.chkcnt.value<myForm.item.length){
				myForm.Sys_BackDate[myForm.chkcnt.value-1].select();
			}
		}else if(City=='台中市'){
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
			if(myForm.chkcnt.value<myForm.item.length){
				myForm.item[myForm.chkcnt.value].select();
			}
		}else{
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
			if(myForm.chkcnt.value<myForm.item.length){
				myForm.mailNumber[myForm.chkcnt.value-1].focus();
			}
		}
	}else{
		alert(errmsg);
	}
<%
else
	response.write "alert('無此催繳單號!!');"
end if
rscnt.close
conn.close
%>
