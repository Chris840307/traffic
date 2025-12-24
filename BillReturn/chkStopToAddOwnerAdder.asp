<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

driverzip="":driveraddr=""

strSQL="select DriverHomeZip,DriverHomeAddress from BillbaseDCIReturn where CarNo=(select distinct CarNo from billbase where ImageFileNameB='"&trim(Request("Sys_BillNo"))&"') and ExchangetypeID='A'"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	driverzip=trim(rscnt("DriverHomeZip"))
	driveraddr=trim(rscnt("DriverHomeAddress"))
end if
rscnt.close

ownerzip="":owneraddr="":sys_carno=""

strSQL="select distinct Carno,OwnerZip,OwnerAddress from BillBase where ImageFileNameB='"&trim(Request("Sys_BillNo"))&"'"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	sys_carno=trim(rscnt("Carno"))
	ownerzip=trim(rscnt("OwnerZip"))
	owneraddr=trim(rscnt("OwnerAddress"))
end if
rscnt.close
%>
	myForm.Sys_CarNo[myForm.chkcnt.value-1].value="<%=sys_carno%>";
	myForm.OwnerZip[myForm.chkcnt.value-1].value="<%=ownerzip%>";
	myForm.OwnerAddress[myForm.chkcnt.value-1].value="<%=owneraddr%>";
	myForm.DriverZip[myForm.chkcnt.value-1].value="<%=driverzip%>";
	myForm.DirverAddress[myForm.chkcnt.value-1].value="<%=driveraddr%>";
	if('<%=owneraddr%>'!='<%=driveraddr%>'){
		myForm.SendKind[myForm.chkcnt.value-1].value='1';
	}else{
		myForm.SendKind[myForm.chkcnt.value-1].value='2';
	}
	myForm.SendKind[myForm.chkcnt.value-1].select();
<%
conn.close
%>
