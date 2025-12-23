<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


strSQL = "select d.DCIRETURNSTATION,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo,d.Owner from BillBase a,Station b,UnitInfo c,BillbaseDCIReturn d where a.RecordStateID=0 and a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIRETURNSTATION=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+) and a.RecordStateID=0 and d.BillNo='"&trim(Ucase(request("BillNo")))&"'"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then%>
	var err=0,errmsg='',City='<%=sys_City%>';
	if(myForm.Sys_Station.value!=""){
		if(myForm.Sys_Station.value!='<%=trim(rscnt("StationID"))%>'){
			err=1;
			errmsg='監理站錯誤\n原監理所：<%=trim(rscnt("StationID"))%>，<%=trim(rscnt("DCIStationName"))%>';
		}
	}
	if(myForm.Sys_UnitID.value!=""){
		if(myForm.Sys_UnitID.value!='<%=trim(rscnt("UnitID"))%>'){
			err=1;
			errmsg=errmsg+'\n單位錯誤\n原單位：<%=trim(rscnt("UnitID"))%>，<%=trim(rscnt("UnitName"))%>';
		}
	}
	if(err=='0'){
		if(City=='花蓮縣'){
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("Owner"))%>';
		}else{
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
		}
		
		if(myForm.chkcnt.value<myForm.item.length){
			if(City=='基隆市'||City=='彰化縣'){
				myForm.item[myForm.chkcnt.value].focus();
			}else if(City=='花蓮縣'){
				myForm.Sys_BackCause[myForm.chkcnt.value-1].focus();
			}else{
				myForm.mailNumber[myForm.chkcnt.value-1].focus();
			}
		}
	}else{
		alert(errmsg);
	}
<%
else
	response.write "alert('無此單號!!');"
end if
rscnt.close
conn.close
%>
