<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	theIllegalDate=funGetDate(gOutDT(request("illegaldate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)

	str_err=""

	If Ifnull(Request("PBillSN")) Then 
		
		If Request("BillTypeID") <> 2 Then
			
			strWhere="where DriverID='"&trim(request("DriverID"))&"' and illegaldate="&theIllegalDate
		else

			strWhere="where CarNo='"&trim(request("Sys_CarNo"))&"' and illegaldate="&theIllegalDate
		End if 

		strSQL="select (select UnitName from Unitinfo where UnitID=passerbase.MemberStation) uitname,(select UnitName from Unitinfo where unitid=passerbase.billunitid) subName,Billno from PasserBase "&strWhere&" and Rule1='"&trim(request("Rule1"))&"' and recordstateid=0"

		strSQL=strSQL&" union all "
		strSQL=strSQL&"select (select UnitName from Unitinfo where UnitID=passerbase.MemberStation) uitname,(select UnitName from Unitinfo where unitid=passerbase.billunitid) subName,Billno from PasserBase "&strWhere&" and Rule2='"&trim(request("Rule1"))&"' and recordstateid=0"

		If trim(Request("rule2")) <>"" Then
			strSQL=strSQL&" union all "
			strSQL=strSQL&"select (select UnitName from Unitinfo where UnitID=passerbase.MemberStation) uitname,(select UnitName from Unitinfo where unitid=passerbase.billunitid) subName,Billno from PasserBase "&strWhere&" and Rule1='"&trim(request("Rule2"))&"' and recordstateid=0"

			strSQL=strSQL&" union all "
			strSQL=strSQL&"select (select UnitName from Unitinfo where UnitID=passerbase.MemberStation) uitname,(select UnitName from Unitinfo where unitid=passerbase.billunitid) subName,Billno from PasserBase "&strWhere&" and Rule1='"&trim(request("Rule2"))&"' and recordstateid=0"
		End if 

		tmpSQL="select distinct uitname,subName,Billno from ("&strSQL&")"

		set rscnt=conn.execute(tmpSQL)
		while Not rscnt.eof
			If str_err<>"" Then str_err=str_err&"\n"

			If str_err = "" Then str_err=str_err&"違規證號\\違規時間\\法條重覆舉發\n"

			str_err=str_err&"舉發單位："&rscnt("uitname")&"_"&rscnt("subName")
			str_err=str_err&"\n單號："&rscnt("Billno")

			rscnt.movenext
		wend
		rscnt.close
	End if 

	If str_err <> "" Then
		Response.Write "alert("""&str_err&""");"&vbCrLf
	else
		Response.Write "if(myForm.PBillSN.value!=''){"&vbCrLf
			Response.Write "myForm.kinds.value=""DB_Update"";"&vbCrLf
			Response.Write "myForm.submit();"&vbCrLf
		Response.Write "}else{"
			Response.Write "myForm.PBillSN.value='';"&vbCrLf
			Response.Write "myForm.kinds.value=""DB_insert"";"&vbCrLf
			Response.Write "myForm.submit();"&vbCrLf
		Response.Write "}"&vbCrLf
	End if 
%>

