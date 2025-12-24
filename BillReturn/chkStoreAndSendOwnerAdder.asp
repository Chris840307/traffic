<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

tb_BillBase="select BillNo,CarNo,BillUnitID,Note from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0"

tb_BillbaseDCIReturn="select BillNo,CarNo,DCIReturnStation,Owner from BillbaseDCIReturn where BillNo='"&trim(Ucase(request("BillNo")))&"'"


strSQL = "select d.DCIReturnStation,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo,d.Owner,a.Note from ("&tb_BillBase&") a,Station b,UnitInfo c,("&tb_BillbaseDCIReturn&") d where a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIReturnStation=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+)"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	If sys_City<>"彰化縣" then
		tmpSql="select BillNo,ExchangeTypeID,ReturnMarkType,DCIReturnStatusID,ExchangeDate from DCILog where billsn=(select sn from billbase where billno='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0) order by ExchangeDate Desc"
		strSQL="select * from ("&tmpSql&") DCILogTmp where rownum=1"

		set rschk=conn.execute(strSQL)
		If Not rschk.eof Then
			If trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="7" and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert(""該筆舉發單有做過收受註記，請先由舉發單維護系統進行撤銷送達!!"");"
			elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="4"  and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert(""該筆舉發單有做過寄存送達!!"");"

			elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="5"  and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert(""該筆舉發單有做過公示送達!!"");"

			elseif trim(rschk("ExchangeTypeID"))="N" and trim(rschk("DCIReturnStatusID"))="n" and sys_City<>"基隆市" Then
				response.write "alert(""該筆舉發單已結案請至上傳下載資料查詢確認!!"");"
			End if
		End if
		rschk.close
	end if%>
	var err=0,errmsg='',City='<%=sys_City%>';
	if(err==0){
		if(City=='花蓮縣'){
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("Owner"))%>';
		}else{
			myForm.CarNo[myForm.chkcnt.value-1].value='<%=trim(rscnt("CarNo"))%>';
		}
		myForm.OwnerZip[myForm.chkcnt.value-1].focus();
	}
<%
else
	response.write "alert('無此單號!!');"
end if
rscnt.close
conn.close
%>
