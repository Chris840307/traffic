<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<%

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQL="select a.*,b.UnitName,c.chname,c.LoginID from (select * from BillRunCarAccept where billunitid='"&trim(Request("UnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("AcceptDate")),0)&"  and recorddate="&funGetDate(Request("RecordDate"),1)&" and recordstateid=0) a,UnitInfo b,memberdata c where a.BillUnitID=b.UnitID and a.BillMemID1=c.MemberID order by a.billno"

	set rsmem=conn.execute(strSQL)
	cmt=0

	Response.Write "myForm.AcceptDate.value='"&gInitDT(rsmem("AcceptDate"))&"';"
	While not rsmem.eof
		
		Response.Write "myForm.item["&cmt&"].value='"&trim(rsmem("BillNo"))&"';"
				
		Response.Write "myForm.CarNo["&cmt&"].value='"&trim(rsmem("CarNo"))&"';"

		Response.Write "myForm.illegalDate["&cmt&"].value='"&gInitDT(rsmem("IllegalDate"))&"';"

		Response.Write "myForm.Rule1["&cmt&"].value='"&trim(rsmem("Rule1"))&"';"

		Response.Write "myForm.IllegalAddress["&cmt&"].value='"&trim(rsmem("IllegalAddress"))&"';"

		Response.Write "myForm.BillMemName["&cmt&"].value='"&trim(rsmem("LoginID"))&"';"

		Response.Write "myForm.BillMemID1["&cmt&"].value='"&trim(rsmem("BillMemID1"))&"';"

		Response.Write "myForm.BillUnitID["&cmt&"].value='"&trim(rsmem("BillUnitID"))&"';"

		Response.Write "BillMemName1["&cmt&"].innerHTML=""<font size=2>"&trim(rsmem("UnitName"))&":"&trim(rsmem("chname"))&"</font>"";"

		If rsmem("recordStateID")=-1 Then
			Response.Write "myForm.Sys_BackBillBase["&cmt&"].value='-1';"
			Response.Write "myForm.chkBackBillBase["&cmt&"].checked=true;"
			Response.Write "myForm.Note["&cmt&"].disabled=false;"
			Response.Write "myForm.Note["&cmt&"].value='"&trim(rsmem("Note"))&"';"
		End if
		

		cmt=cmt+1

		rsmem.movenext	
	Wend
	rsmem.close

	conn.close
	set conn=nothing
%>
