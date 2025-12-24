<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
If sys_City="高雄市" or sys_City="高港局" then

	strSQL="select CarNo,Owner,DriverHomeZip,DriverHomeAddress,decode(OwnerNotIfyAddress,null,DriverHomeZip,substr(OwnerNotIfyAddress,1,3)) OwnerHomeZip,decode(OwnerNotIfyAddress,null,DriverHomeAddress,substr(OwnerNotIfyAddress,4)) OwnerHomeAddress,OwnerZip,OwnerAddress from BillbaseDCIReturn where CarNo=(select CarNo from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0) and ExchangeTypeID='A'"

else
	tb_BillBase="select CarNo,OwnerZip OwnerHomeZip,OwnerAddress OwnerHomeAddress,Driverzip DriverHomeZip,DriverAddress DriverHomeAddress from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0"

	tb_BillbaseDCIReturn="select CarNo,Owner,DriverHomeZip,DriverHomeAddress,decode(OwnerNotIfyAddress,null,OwnerZip,substr(OwnerNotIfyAddress,1,3)) OwnerZip,decode(OwnerNotIfyAddress,null,OwnerAddress,substr(OwnerNotIfyAddress,4)) OwnerAddress from BillbaseDCIReturn where CarNo=(select CarNo from BillBase where BillNo='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0) and ExchangeTypeID='A'"

	
strSQL = "select b.CarNo,b.Owner,a.OwnerHomeZip,a.OwnerHomeAddress,decode(a.DriverHomeZip,null,b.DriverHomeZip,a.DriverHomeZip) DriverHomeZip,decode(a.DriverHomeAddress,null,b.DriverHomeAddress,a.DriverHomeAddress) DriverHomeAddress,b.OwnerZip,b.OwnerAddress from ("&tb_BillBase&") a,("&tb_BillbaseDCIReturn&") b where a.CarNo=b.CarNo(+)"
end if

set rscnt=conn.execute(strSQL)
if Not rscnt.eof then
	If sys_City<>"彰化縣" then
		strSQL="select BillNo,ExchangeTypeID,ReturnMarkType,DCIReturnStatusID,ExchangeDate from DCILog where billsn=(select sn from billbase where billno='"&trim(Ucase(request("BillNo")))&"' and RecordStateID=0) and billno is not null order by ExchangeDate Desc"
		'strSQL="select * from ("&tmpSql&") DCILogTmp where rownum=1"

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
	end If  

	carno="":ownerzip="":ownerAddr="":DriverZip="":DriverAddr=""

	If sys_City="花蓮縣" then
		carno=trim(rscnt("Owner"))
	else
		carno=trim(rscnt("CarNo"))
	end if
	If not ifnull(rscnt("OwnerHomeAddress")) Then
		ownerzip=trim(rscnt("OwnerHomeZip"))
		ownerAddr=trim(rscnt("OwnerHomeAddress"))

	elseIf not ifnull(rscnt("OwnerAddress")) Then

		ownerzip=trim(rscnt("OwnerZip"))
		ownerAddr=trim(rscnt("OwnerAddress"))
	End if

	If not ifnull(rscnt("DriverHomeAddress")) Then
		DriverZip=trim(rscnt("DriverHomeZip"))
		DriverAddr=trim(rscnt("DriverHomeAddress"))

	End if
	
	%>
	var CarNo='<%=carno%>';
	var City='<%=sys_City%>',OwnerZip='<%=ownerzip%>',OwnerAddress='<%=ownerAddr%>';
	var DriverZip='<%=DriverZip%>',DriverAddress='<%=DriverAddr%>';
	var BackIndex=myForm.Sys_BackCause[myForm.chkcnt.value-1].selectedIndex;

	myForm.CarNo[myForm.chkcnt.value-1].value=CarNo;
	
	myForm.OwnerZip[myForm.chkcnt.value-1].value=OwnerZip;
	myForm.OwnerAddress[myForm.chkcnt.value-1].value=OwnerAddress;

	myForm.DriverZip[myForm.chkcnt.value-1].value=DriverZip;
	myForm.DriverAddress[myForm.chkcnt.value-1].value=DriverAddress;

	if(OwnerAddress!=DriverAddress && DriverAddress!=''){
		myForm.ChkSend[myForm.chkcnt.value-1].value='是';
		
		myForm.Sys_BackCause[myForm.chkcnt.value-1].options[0]=new Option('其他寄存原因','T');
		myForm.Sys_BackCause[myForm.chkcnt.value-1].options[0].selected=true;
		myForm.Sys_BackCause[myForm.chkcnt.value-1].length=1;
	}else{
		myForm.ChkSend[myForm.chkcnt.value-1].value='否';
		<%
		strSQL="select ID,Content from DCICode where TypeID=7 and ID in('1','2','3','4','8','M','K','L','O','P','Q')"
		set rs1=conn.execute(strSQL)
		selectCmt=0
		while Not rs1.eof
			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('"&trim(rs1("Content"))&"','"&trim(rs1("ID"))&"');"& vbcrlf

			selectCmt=selectCmt+1
			rs1.movenext
		wend
		rs1.close

		if sys_City="南投縣" or sys_City="雲林縣" then
			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('車主死亡','AddReason1');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('拒收','AddReason2');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('公司無此車號','AddReason3');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('無此公司','AddReason4');"& vbcrlf

			selectCmt=selectCmt+1
		end if

		response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].length="&selectCmt&";"& vbcrlf
		response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options[BackIndex].selected=true;"& vbcrlf
		%>
	}

	myForm.DriverZip[myForm.chkcnt.value-1].select();
<%
else
	response.write "alert('無此單號!!');"
end if
rscnt.close
conn.close
%>
