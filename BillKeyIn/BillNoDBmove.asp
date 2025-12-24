<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	If trim(Session("Credit_ID"))="A000000000" Then
		strSQL="select * from (select rownum as Item,a.* from (Select * from PasserBase where BillNo = '"&request("Sys_Billno")&"' and RecorDStateID<>-1 order by RecordDate) a) "
	else
		strSQL="select * from (select rownum as Item,a.* from (Select * from PasserBase where (MemberStation in(select UnitID from Unitinfo where UnitTypeID=(select UnitTypeID from Unitinfo where UnitID='"&trim(Session("Unit_ID"))&"')) or RecordMemberID="&trim(Session("User_ID"))&") and RecorDStateID<>-1 order by RecordDate) a) where BillNo = '"&request("Sys_Billno")&"'"
	end If 

	set rs=conn.execute(strSQL)
	if Not rs.eof then
		response.write "myForm.Sys_DoubleCheckStatus.value='"&rs("DoubleCheckStatus")&"';"
		response.write "CreatDate.innerHTML='"&gInitDT(rs("RecordDate"))&"';"
		response.write "myForm.PBillSN.value='"&trim(rs("SN"))&"';"
		response.write "myForm.Billno1.value='"&trim(rs("BillNo"))&"';"
		response.write "myForm.Sys_BilltypeID.value='"&trim(rs("BilltypeID"))&"';"
		response.write "myForm.Insurance.value='"&trim(rs("Insurance"))&"';"
		response.write "myForm.Sys_CarNo.value='"&trim(rs("CarNo"))&"';"
		response.write "myForm.DriverName.value='"&trim(rs("Driver"))&"';"
		response.write "myForm.DriverBrith.value='"&ginitdt(trim(rs("DriverBirth")))&"';"
		response.write "myForm.DriverPID.value='"&trim(rs("Driverid"))&"';"
		response.write "myForm.DriverSEX.value='"&trim(rs("DriverSex"))&"';"
		response.write "myForm.DriverZip.value='"&trim(rs("DriverZip"))&"';"
		response.write "myForm.DriverAddress.value='"&trim(rs("DriverAddress"))&"';"
		response.write "myForm.IllegalDate.value='"&ginitdt(trim(rs("IllegalDate")))&"';"
		response.write "myForm.IllegalTime.value='"&right("00"&hour(rs("IllegalDate")),2)&right("00"&minute(rs("IllegalDate")),2)&"';"
		response.write "myForm.IllegalAddressID.value='"&trim(rs("IllegalAddressID"))&"';"
		response.write "myForm.IllegalAddress.value='"&trim(rs("IllegalAddress"))&"';"
		response.write "myForm.Rule1.value='"&trim(rs("Rule1"))&"';"

		strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs("Rule1"))&"' and Version='"&trim(rs("RuleVer"))&"'"
		set rsR1=conn.execute(strR1)
		if not rsR1.eof then 
			response.write "Layer1.innerHTML='"&trim(rsR1("IllegalRule"))&"';"
		end if
		rsR1.close
		set rsR1=nothing

		response.write "myForm.ForFeit1.value='"&trim(rs("ForFeit1"))&"';"
		response.write "myForm.Rule2.value='"&trim(rs("Rule2"))&"';"

		strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs("Rule2"))&"' and Version='"&trim(rs("RuleVer"))&"'"
		set rsR1=conn.execute(strR1)
		if not rsR1.eof then 
			response.write "Layer2.innerHTML='"&trim(rsR1("IllegalRule"))&"';"
		end if
		rsR1.close
		set rsR1=nothing

		response.write "myForm.ForFeit2.value='"&trim(rs("ForFeit2"))&"';"

		response.write "myForm.RuleSpeed.value='"&trim(rs("RuleSpeed"))&"';"
		response.write "myForm.IllegalSpeed.value='"&trim(rs("IllegalSpeed"))&"';"

		response.write "myForm.DealLineDate.value='"&ginitdt(trim(rs("DealLineDate")))&"';"
		response.write "myForm.MemberStation.value='"&trim(rs("MemberStation"))&"';"
		
		strS="select UnitName from UnitInfo where UnitID='"&trim(rs("MemberStation"))&"'"
		set rsS=conn.execute(strS)
		if not rsS.eof then
			response.write "Layer5.innerHTML='"&trim(rsS("UnitName"))&"';"
		end if
		rsS.close
		set rsS=nothing

		if trim(rs("BillMemID1"))<>"" then
			strMeb="select LoginID from MemberData where MemberID="&trim(rs("BillMemID1"))
			set rsmeb=conn.execute(strMeb)
			if not rsmeb.eof then Sys_Meb1=rsmeb("LoginID")
			rsmeb.close
		end if
		if sys_City="高雄縣" or sys_City="高雄市" then
			response.write "myForm.BillMem1.value='"&Sys_Meb1&"';"
			response.write "Layer12.innerHTML='"&trim(rs("BillMem1"))&"';"
		else
			response.write "myForm.BillMem1.value='"&trim(rs("BillMem1"))&"';"
			response.write "Layer12.innerHTML='"&Sys_Meb1&"';"
		end if
		response.write "myForm.BillMemID1.value='"&trim(rs("BillMemID1"))&"';"
		response.write "myForm.BillMemName1.value='"&trim(rs("BillMem1"))&"';"
		
		if trim(rs("BillMemID2"))<>"" then
			strMeb="select LoginID from MemberData where MemberID="&trim(rs("BillMemID2"))
			set rsmeb=conn.execute(strMeb)
			if not rsmeb.eof then Sys_Meb2=rsmeb("LoginID")
			rsmeb.close
		end if
		if sys_City="高雄縣" then
			response.write "myForm.BillMem2.value='"&Sys_Meb2&"';"
			response.write "Layer13.innerHTML='"&trim(rs("BillMem2"))&"';"
		else
			response.write "myForm.BillMem2.value='"&trim(rs("BillMem2"))&"';"
			response.write "Layer13.innerHTML='"&Sys_Meb2&"';"
		end if	
		response.write "myForm.BillMemID2.value='"&trim(rs("BillMemID2"))&"';"
		response.write "myForm.BillMemName2.value='"&trim(rs("BillMem2"))&"';"
		
		if trim(rs("BillMemID3"))<>"" then
			strMeb="select LoginID from MemberData where MemberID="&trim(rs("BillMemID3"))
			set rsmeb=conn.execute(strMeb)
			if not rsmeb.eof then Sys_Meb3=rsmeb("LoginID")
			rsmeb.close
		end if
		if sys_City="高雄縣" then
			response.write "myForm.BillMem3.value='"&Sys_Meb3&"';"
			response.write "Layer14.innerHTML='"&trim(rs("BillMem3"))&"';"
		else
			response.write "myForm.BillMem3.value='"&trim(rs("BillMem3"))&"';"
			response.write "Layer14.innerHTML='"&Sys_Meb3&"';"
		end if
		response.write "myForm.BillMemID3.value='"&trim(rs("BillMemID3"))&"';"
		response.write "myForm.BillMemName3.value='"&trim(rs("BillMem3"))&"';"

		if trim(rs("BillMemID4"))<>"" then
			strMeb="select LoginID from MemberData where MemberID="&trim(rs("BillMemID4"))
			set rsmeb=conn.execute(strMeb)
			if not rsmeb.eof then Sys_Meb4=rsmeb("LoginID")
			rsmeb.close
		end if
		if sys_City="高雄縣" then
			response.write "myForm.BillMem4.value='"&Sys_Meb4&"';"
			response.write "Layer17.innerHTML='"&trim(rs("BillMem4"))&"';"
		else
			response.write "myForm.BillMem4.value='"&trim(rs("BillMem4"))&"';"
			response.write "Layer17.innerHTML='"&Sys_Meb4&"';"
		end if
		response.write "myForm.BillMemID4.value='"&trim(rs("BillMemID4"))&"';"
		response.write "myForm.BillMemName4.value='"&trim(rs("BillMem4"))&"';"
		response.write "myForm.SignType.value='"&trim(rs("SignType"))&"';"

		strConfiscate=""
		strFas="select ConfiscateID from PasserConfiscate where BillSN="&trim(rs("SN"))
		set rsFas=conn.execute(strFas)
		While Not rsFas.Eof
			if strConfiscate="" then
				strConfiscate=trim(rsFas("ConfiscateID"))
			else
				strConfiscate=strConfiscate&","&trim(rsFas("ConfiscateID"))
			end if
			rsFas.MoveNext
		Wend
		rsFas.close
		set rsFas=nothing
		response.write "for(var i=0;i<myForm.Fastener1.length;i++){"
		response.write "myForm.Fastener1[i].checked=false;"
		response.write "}"
		strItem="select ID,Content from Code where TypeID=2 and Not(ID<478 or ID=479) order by ID"
		set rsItem=conn.execute(strItem)
		ItemCunt=-1
		While Not rsItem.Eof
			ItemCunt=ItemCunt+1
			if strConfiscate<>"" then
				if instr(strConfiscate,trim(rsItem("ID")))>0 then
					response.write "myForm.Fastener1["&ItemCunt&"].checked=true;"
				end if
			end if
			rsItem.MoveNext
		Wend
		rsItem.close
		set rsItem=nothing
		
		response.write "myForm.BillUnitID.value='"&trim(rs("BillUnitID"))&"';"

		strU="select UnitName from UnitInfo where UnitID='"&trim(rs("BillUnitID"))&"'"
		set rsU=conn.execute(strU)
		if not rsU.eof then
			response.write "Layer6.innerHTML='"&trim(rsU("UnitName"))&"';"
		end if
		rsU.close
		set rsU=nothing

		response.write "myForm.ProjectID.value='"&trim(rs("ProjectID"))&"';"

		strU="select name from Project where projectid='"&trim(rs("ProjectID"))&"'"
		set rsU=conn.execute(strU)
		if not rsU.eof then
			response.write "Layer001.innerHTML='"&trim(rsU("name"))&"';"
		end if
		rsU.close
		set rsU=nothing

		response.write "myForm.BillFillDate.value='"&ginitdt(trim(rs("BillFillDate")))&"';"
		response.write "myForm.Note.value='"&trim(rs("Note"))&"';"
	end if
%>
