<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<%
	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	
	
	chk_CarNo="":chk_IllegalDate="":chk_IllegalTime="":chk_BillUnitID="":chk_BillMemID1="":chk_CarSimpleID=""
	strSQL="select BillNo,CarNo,IllegalDate,CarSimpleID,BillUnitID,(select LoginID from MemberData where MemberID=billStopcaraccept.BillMemID1) BillMem1 from billStopcaraccept where BillNo='"&trim(Request("BillNo"))&"' and RecordStateID=0"
	set chkrs=conn.execute(strSQL)
	If not chkrs.eof Then
		chk_CarNo=trim(chkrs("CarNo"))
		chk_CarSimpleID=trim(chkrs("CarSimpleID"))
		chk_IllegalDate=gInitDT(trim(chkrs("IllegalDate")))
		chk_IllegalTime=right("00"&hour(chkrs("IllegalDate")),2)&right("00"&minute(chkrs("IllegalDate")),2)
		chk_BillUnitID=trim(chkrs("BillUnitID"))
		chk_BillMemID1=trim(chkrs("BillMem1"))
	End if 
	chkrs.close
	
	Response.Write "var errmsg="""";"
	If not ifnull(chk_CarNo) Then		
		Response.Write "if(myForm.CarNo.value!='"&chk_CarNo&"'){errmsg=errmsg+'車號與登記簿不同：'+myForm.CarNo.value+'（"&chk_CarNo&"）\n';}"
		Response.Write "if(myForm.CarSimpleID.value!='"&chk_CarSimpleID&"'){errmsg=errmsg+'簡式車種與登記簿不同：'+myForm.CarSimpleID.value+'（"&chk_CarSimpleID&"）\n';}"
		Response.Write "if(myForm.IllegalDate.value!='"&chk_IllegalDate&"'){errmsg=errmsg+'違規日與登記簿不同：'+myForm.IllegalDate.value+'（"&chk_IllegalDate&"）\n';}"
		Response.Write "if(myForm.IllegalTime.value!='"&chk_IllegalTime&"'){errmsg=errmsg+'違規時間與登記簿不同：'+myForm.IllegalTime.value+'（"&chk_IllegalTime&"）\n';}"
		Response.Write "if(myForm.BillUnitID.value!='"&chk_BillUnitID&"'){errmsg=errmsg+'舉發單位與登記簿不同：'+myForm.BillUnitID.value+'（"&chk_BillUnitID&"）\n';}"
		Response.Write "if(myForm.BillMem1.value!='"&chk_BillMemID1&"'){errmsg=errmsg+'舉發人1與登記簿不同：'+myForm.BillMem1.value+'（"&chk_BillMemID1&"）\n';}"
	Else
		response.write "errmsg='此單號不在登記簿內';"
	End if 
	Response.Write "if(errmsg!=""""){"
	Response.Write "if(confirm(errmsg+""\n是否確定要存檔？"")){"
	Response.Write "myForm.kinds.value=""DB_insert"";"
	Response.Write "myForm.submit();"
	Response.Write "}"
	Response.Write "}else{myForm.kinds.value=""DB_insert"";myForm.submit();}"
%>
