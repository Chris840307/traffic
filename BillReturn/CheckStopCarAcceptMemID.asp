<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<%

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQL="select a.chname,a.MemberID,b.UnitID,b.UnitName from (select chname,MemberID,unitid from memberdata where Loginid='"&Trim(Request("LoginID"))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

	set rsmem=conn.execute(strSQL)
	cmt=0
	sys_chname=""
	While not rsmem.eof
		If not ifnull(sys_chname) Then sys_chname=sys_chname&","

		sys_chname=sys_chname&trim(rsmem("UnitName"))&":"&trim(rsmem("chname"))
		cmt=cmt+1

		rsmem.movenext	
	Wend
	If cmt=1 Then
		rsmem.movefirst
		Response.Write "myForm.BillMemID1["&Request("itemcnt")&"].value="""&trim(rsmem("MemberID"))&""";"
		Response.Write "myForm.BillUnitID["&Request("itemcnt")&"].value="""&trim(rsmem("UnitID"))&""";"
	End if
	
	rsmem.close

	Response.Write "BillMemName1["&Request("itemcnt")&"].innerHTML=""<font size=2>"&sys_chname&"</font>"";"
	If cmt>1 Then response.write "alert('人員代碼重覆，請至人員管理確認!!');"

	conn.close
	set conn=nothing
%>
