<!--#include virtual="traffic/Common/DB.ini"-->
<%
dim cnt

If trim(Request("itemcnt")) <>"" Then Response.Write "myForm.chkcnt.value="&trim(Request("itemcnt"))&";"

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

tb_BillBase="select BillNo,CarNo,BillUnitID from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0"

tb_BillbaseDCIReturn="select BillNo,CarNo,DCIReturnStation,Owner from BillbaseDCIReturn where BillNo='"&trim(Ucase(request("BillNo")))&"'"


strSQL = "select (select max(billstatus) from billbase where billno='"&trim(Ucase(request("BillNo")))&"') billstatus,d.DCIReturnStation,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,c.UnitTypeID,d.CarNo,d.Owner from ("&tb_BillBase&") a,Station b,UnitInfo c,("&tb_BillbaseDCIReturn&") d where a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIReturnStation=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+)"

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	If sys_City<>"彰化縣" then
		strSQL="select BillNo,ExchangeTypeID,ReturnMarkType,DCIReturnStatusID,ExchangeDate from DCILog where billsn=(select sn from billbase where billno='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0) and billno is not null order by ExchangeDate Desc"
		'strSQL="select * from ("&tmpSql&") DCILogTmp where rownum=1"

		set rschk=conn.execute(strSQL)
		If Not rschk.eof Then
			If trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="7" and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單有做過收受註記，請先由舉發單維護系統進行撤銷送達!!"");"

			elseIf trim(rschk("DCIReturnStatusID"))="h" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單已開裁決書，請洽監理機關查詢!!"");"

			elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="4"  and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單有做過寄存送達!!"");"

			elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="5"  and trim(rschk("DCIReturnStatusID"))="S" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單有做過公示送達!!"");"

			elseif trim(rschk("ExchangeTypeID"))="N" and trim(rschk("DCIReturnStatusID"))="n" and sys_City<>"基隆市" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單已結案請至上傳下載資料查詢確認!!"");"

			elseif trim(rscnt("billstatus"))="9" and sys_City<>"基隆市"  and sys_City<>"台中市" Then
				response.write "alert("""&request("BillNo")&"該筆舉發單已結案請至舉發單資料維護查詢確認!!"");"

			End If

			If sys_City="宜蘭縣" Then
				If instr(trim(Ucase(request("BillNo"))),"QZ")>0 Then 
				  If Trim(Session("User_ID"))<>"2424" and Trim(Session("User_ID"))<>"5682"  and Trim(Session("User_ID"))<>"6227" and Trim(Session("User_ID"))<>"6607" Then 
					response.write "alert("""&request("BillNo")&"該單號是不屬於停管場!!"");"
				  End If
				Else
				  If Trim(Session("User_ID"))="2424" or Trim(Session("User_ID"))="5682" or Trim(Session("User_ID"))="6227" or Trim(Session("User_ID"))="6607" Then 
					response.write "alert("""&request("BillNo")&"該單號是不屬於停管場!!"");"
				  End If				
				End if
			End if

		End if
		rschk.close
	end if%>
	var err=0,errmsg='',City='<%=sys_City%>';
	if(myForm.Sys_Station.value!=""){
		if(myForm.Sys_Station.value!='<%=trim(rscnt("StationID"))%>'){
			err=1;
			errmsg='監理站錯誤\n原監理所：<%=trim(rscnt("StationID"))%>，<%=trim(rscnt("DCIStationName"))%>';
		}
	}
	if(myForm.Sys_UnitID.value!=""){
		if(myForm.Sys_UnitID.value.search("<%=trim(rscnt("UnitID"))%>")<0 && myForm.Sys_UnitID.value.search("<%=trim(rscnt("UnitTypeID"))%>")<0){
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
			}else if(City=='雲林縣'){
				myForm.Sys_BackDate[myForm.chkcnt.value-1].focus();
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
	response.write "alert('"&request("BillNo")&"無此單號!!');"
end if
rscnt.close
conn.close
%>
