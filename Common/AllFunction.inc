<%
Function chkValue(ValueId) 
        Dim ValueTmp 
		ValueTmp = ""
        ValueTmp = Replace(Trim(ValueId & ""), "--", "")
        ValueTmp = Replace(ValueTmp, "<", "")
        ValueTmp = Replace(ValueTmp, ">", "")
        ValueTmp = Replace(ValueTmp, "&", "")
        ValueTmp = Replace(ValueTmp, "\", "")
        ValueTmp = Replace(ValueTmp, "sleep", "")
        ValueTmp = Replace(ValueTmp, "delay", "")
        ValueTmp = Replace(ValueTmp, "script", "")
        chkValue = ValueTmp
End Function

'------------------------------------------------
Function LawRule(LawID,illRule,CarSimpleID,CarAddID)
	 
	if lawID="2110101" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="未領有駕駛執照駕駛小型車"

	elseif lawID="2110101" then
		LawRule="未領有駕駛執照駕駛機車"

	elseif lawID="2110102" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="未滿十八歲之人駕駛小型車"

	elseif lawID="2110102" then
		LawRule="未滿十八歲之人駕駛機車"

	elseif lawID="2110301" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="使用偽造之駕駛執照駕駛小型車"

	elseif lawID="2110301" then
		LawRule="使用偽造之駕駛執照駕駛機車"

	elseif lawID="2110302" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="未滿十八歲之人使用偽造之駕駛執照駕駛小型車"

	elseif lawID="2110302" then
		LawRule="未滿十八歲之人使用偽造之駕駛執照駕駛機車"

	elseif lawID="2110303" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="使用變造之駕駛執照駕駛小型車"

	elseif lawID="2110303" then
		LawRule="使用變造之駕駛執照駕駛機車"

	elseif lawID="2110303" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="使用變造之駕駛執照駕駛小型車"

	elseif lawID="2110303" then
		LawRule="使用變造之駕駛執照駕駛機車"

	elseif lawID="2110303" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="使用變造之駕駛執照駕駛小型車"

	elseif lawID="2110303" then
		LawRule="使用變造之駕駛執照駕駛機車"

	elseif lawID="2110304" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="未滿十八歲之人使用變造之駕駛執照駕駛小型車"

	elseif lawID="2110304" then
		LawRule="未滿十八歲之人使用變造之駕駛執照駕駛機車"

	elseif lawID="2110305" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="使用矇領之駕駛執照駕駛小型車"

	elseif lawID="2110305" then
		LawRule="使用矇領之駕駛執照駕駛機車"

	elseif lawID="2110306" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="未滿十八歲之人使用矇領之駕駛執照駕駛小型車"

	elseif lawID="2110306" then
		LawRule="未滿十八歲之人使用矇領之駕駛執照駕駛機車"

	elseif lawID="2110401" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="駕駛執照業經吊銷仍駕駛小型車"

	elseif lawID="2110401" then
		LawRule="駕駛執照業經吊銷仍駕駛機車"

	elseif lawID="2110402" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="駕駛執照業經註銷仍駕駛小型車"

	elseif lawID="2110402" then
		LawRule="駕駛執照業經註銷仍駕駛機車"

	elseif lawID="2110501" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="駕駛執照吊扣期間駕駛小型車"

	elseif lawID="2110501" then
		LawRule="駕駛執照吊扣期間駕駛機車"

	elseif lawID="2210401" and (CarAddID=7) then
		LawRule="領有聯結車駕駛執照駕駛大型重型機車"

	elseif lawID="2210401" then
		LawRule="領有聯結車駕駛執照駕駛重型機車"

	elseif lawID="2210402" and (CarAddID=7) then
		LawRule="領有大客車駕駛執照駕駛大型重型機車"

	elseif lawID="2210402" then
		LawRule="領有大客車駕駛執照駕駛重型機車"

	elseif lawID="2210403" and (CarAddID=7) then
		LawRule="領有大貨車駕駛執照駕駛大型重型機車"

	elseif lawID="2210403" then
		LawRule="領有大貨車駕駛執照駕駛重型機車"

	elseif lawID="2210404" and (CarAddID=7) then
		LawRule="領有小型車駕駛執照駕駛大型重型機車"

	elseif lawID="2210404" then
		LawRule="領有小型車駕駛執照駕駛重型機車"

	elseif lawID="2210601" and (CarAddID=7) then
		LawRule="領有輕型機車駕駛執照駕駛大型重型機車"

	elseif lawID="2210601" then
		LawRule="領有輕型機車駕駛執照駕駛重型機車"

	elseif lawID="3310101" and (CarAddID=7) then
		LawRule="領有輕型機車駕駛執照駕駛大型重型機車"

	elseif lawID="3310101" then
		LawRule="領有輕型機車駕駛執照駕駛重型機車"

	elseif lawID="4310102" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="以危險方式在道路上駕駛汽車"

	elseif lawID="4310102" then
		LawRule="以危險方式在道路上駕駛機車"

	elseif lawID="4310103" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="駕駛汽車在道路上蛇行因而肇事者"

	elseif lawID="4310103" then
		LawRule="駕駛機車在道路上蛇行因而肇事者"

	elseif lawID="4310104" and (CarSimpleID = 1 or CarSimpleID = 2) then
		LawRule="以危險方式在道路上駕駛汽車因而肇事者"

	elseif lawID="4310104" then
		LawRule="以危險方式在道路上駕駛機車因而肇事者"

	else
		LawRule=illRule

	end if

End Function

Function gInitDT(iDate)
'Date 格式 轉換 字串
'系統所要顯示的日期
Dim iTemp 
	if trim(iDate)<>"" then
		iTemp=right(year(iDate)-1911,3) &_
		right("0"&month(iDate),2) &_
		right("0"&day(iDate),2)
	    gInitDT = iTemp
	else
		gInitDT = ""
	end if
End Function

Function gArrMT(iDate)
'Date 格式 轉換 字串
'系統所要顯示的日期
Dim iTemp 
	if trim(iDate)<>"" then
		iTemp=left(iDate,len(iDate)-4) &"-"&_
		Mid(iDate,len(iDate)-3,2) &"-"&_
		right(iDate,2)
		gArrMT = iTemp
	else
		gArrMT = "- -"
	end if
End Function

Function gArrDT(iDate)
'Date 格式 轉換 字串
'系統所要顯示的日期
Dim iTemp 
	if trim(iDate)<>"" then
		iTemp=right(year(iDate)-1911,3) &"-"&_
		right("0"&month(iDate),2) &"-"&_
		right("0"&day(iDate),2)
	    gArrDT = iTemp
	else
		gArrDT = "- -"
	end if
End Function

Function gOutDT(iDate)
'Date 格式 轉換 字串
'系統所要顯示的日期
	if trim(iDate)<>"" then
		iDate=trim(iDate)
		Dim tTemp
		tTemp=DateSerial(Left(iDate,len(iDate)-4)+1911,Mid(iDate,len(iDate)-3,2),right(iDate,2))
		gOutDT = tTemp
	else
		gOutDT = ""
	end if
End Function
'------------------------------------------------
'光棒
'	response.write "<tr align='center' bgcolor='#FFFFFF'"
'	lightbarstyle 0
'	response.write ">"
'
function lightbarstyle(styleindex)
   if styleindex=0 then
	response.write "onMouseOver=""this.style.backgroundColor='#CCFFFF'"" onMouseOut=""this.style.backgroundColor='#FFFFFF'"""
   elseif styleindex=1 then
	response.write "onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='#FFFFFF'"""
   end if
end function
'------------------------------------------------
Function funBillType(ObjName,DefValue,reValue)
	BillTypeID=split("1,2,3,4,5,A,B,N",",")
	BillTypeName=split("當場,逕行,肇事,拖吊,載運砂石士方,違規營業,違規重標,未知",",")
	if trim(ObjName)<>"" then
		for billcunt=0 to Ubound(BillTypeID)
			response.write "<option value="""&BillTypeID(billcunt)&""""
			if trim(BillTypeID(billcunt))=trim(DefValue) then response.write " selected"
			response.write ">"&BillTypeName(billcunt)&"</option>"
		next
	end if
	if trim(reValue)<>"" then
		for idtoname=0 to Ubound(BillTypeID)
			if trim(BillTypeID(idtoname))=trim(reValue) then response.write BillTypeName(idtoname)
		next
	end if
end function

Function funTableSeq(tableName,filedName)
	dim cnt
	strSQL="select Max("&filedName&")+1 as cnt from "&tableName
	set rscnt=conn.execute(strSQL)
	cnt=rscnt("cnt")
	rscnt.close
	if ISNULl(cnt) then cnt=1
	funTableSeq=cnt
end function

Function funGetDate(strDay,indx)
	if trim(strDay)<>"" then
		if indx=1 then
			dbDay="TO_DATE('"&formatDatetime(strDay,2)&" "&hour(strDay)&":"&Minute(strDay)&":"&Second(strDay)&"','YYYY/MM/DD HH24:MI:SS')"
		else
			dbDay="TO_DATE('"&strDay&"','YYYY/MM/DD')"
		end if
		funGetDate=dbDay
	else
		funGetDate="''"
	end if
end function

Function funTnumber(strValue)
	if ISNull(strValue) or trim(strValue)="" or IsEmpty(strValue) then
		funTnumber=0
	else
		funTnumber=strValue
	end if
end function

Function ConnExecute(Log_Content,Log_TypeID)
	if trim(Log_Content)<>"" or trim(Log_TypeID)<>"" then
		if trim(Log_TypeID)="" then Log_TypeID=0
		userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

		userip=replace(userip,"::1","127.0.0.1")

		ContentSQL=replace(Log_Content,"'","")
		logsql="Insert Into traffic.Log(SN,TypeID,ActionMemBerID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(sn),0)+1 from traffic.log),"&Log_TypeID&","&Session("User_ID")&",'"&Session("Ch_Name") &"','"&userip&"',sysdate,'"&ContentSQL&"')"
		conn.execute(logsql)

		strCity="select value from traffic.Apconfigure where id=31"
		set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
		rsCity.close

		if sys_City="高雄市" and (Log_TypeID=361 or Log_TypeID=362) then
			dim ConnTmp
			Set ConnTmp = Server.CreateObject("ADODB.Connection")
			Provider="Provider=MSDAORA;Data Source=trafficorcl;User Id=traffic;Password=joly902f;"
			ConnTmp.Open Provider
				logsql="Insert Into traffic.Log(SN,TypeID,ActionMemBerID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(sn),0)+1 from traffic.log),"&Log_TypeID&","&Session("User_ID")&",'"&Session("Ch_Name") &"','"&userip&"',sysdate,'"&ContentSQL&"')"
				ConnTmp.execute(logsql)
			ConnTmp.close
			set ConnTmp=nothing
		end if 
	end if
end function

Function Crt_Log(Log_User,Log_Reason,Log_Content,Log_TypeID,Log_Tmp1,LogTmp2)
	if trim(Log_Content)<>"" or trim(Log_TypeID)<>"" then
		if trim(Log_TypeID)="" then Log_TypeID=0
		userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

		userip=replace(userip,"::1","127.0.0.1")

		ContentSQL=replace(Log_Content,"'","")

		if trim(Log_User)="" then Log_User=Session("Ch_Name")

		logsql="Insert Into traffic.Log(SN,TypeID,MainQryUser,ActionMemBerID,ActionChName,ActionIP,ActionDate,Log_Reason,ActionContent) values((select nvl(max(sn),0)+1 from traffic.log),"&Log_TypeID&",'"&Log_User&"',"&Session("User_ID")&",'"&Session("Ch_Name") &"','"&userip&"',sysdate,'"&Log_Reason&"','"&ContentSQL&"')"
		conn.execute(logsql)

		strCity="select value from traffic.Apconfigure where id=31"
		set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
		rsCity.close

		if sys_City="高雄市" and (Log_TypeID=361 or Log_TypeID=362) then
			dim ConnTmp
			Set ConnTmp = Server.CreateObject("ADODB.Connection")
			Provider="Provider=MSDAORA;Data Source=trafficorcl;User Id=traffic;Password=joly902f;"
			ConnTmp.Open Provider
				logsql="Insert Into traffic.Log(SN,TypeID,MainQryUser,ActionMemBerID,ActionChName,ActionIP,ActionDate,Log_Reason,ActionContent) values((select nvl(max(sn),0)+1 from traffic.log),"&Log_TypeID&",'"&Log_User&"',"&Session("User_ID")&",'"&Session("Ch_Name") &"','"&userip&"',sysdate,'"&Log_Reason&"','"&ContentSQL&"')"
				ConnTmp.execute(logsql)
			ConnTmp.close
			set ConnTmp=nothing
		end if 
	end if
end function

Function showMap(MapSN)
	tmpstr=""
	strSql="select SN,IMAGEDIRECTORYNAME,IMAGEFILENAME from PASSERSENDARRIVED where SN="&MapSN
	set rsmap=conn.execute(strsql)
	if Not rsmap.eof then
		tmpstr="onMouseOver=""popup(1,'"&rsmap("IMAGEDIRECTORYNAME")&rsmap("IMAGEFILENAME")&"');"" onMouseOut=""popup(0,'');"""
	end if
	showMap=tmpstr
end function

Function SelectUnitOption(UnitName,MemberName)
	strtmp=""
	strUnitID=""
	if trim(session("ManagerPower"))="1" then
		if trim(MemberName)<>"" then
			strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitMan('"&UnitName&"','"&MemberName&"');"">"
		else
			strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"">"
		end if

		if trim(Session("UnitLevelID"))="1" then
			strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder,UnitTypeID,UnitName"
			strtmp=strtmp+"<option value="""">所有單位</option>"
		elseif trim(Session("UnitLevelID"))="2" then
			strSQL="select UnitID,UnitName from UnitInfo where UnitTypeid in(select UnitTypeid from Unitinfo where Unitid='"&Session("Unit_ID")&"') order by UnitOrder,UnitTypeID,UnitName"

			set rs1=conn.execute(strSQL)
			while Not rs1.eof
				if trim(strUnitID)<>"" then strUnitID=trim(strUnitID)&","
				if trim(strUnitID)="" then
					strUnitID=strUnitID&trim(rs1("UnitID"))
				else
					strUnitID=strUnitID&"'"&trim(rs1("UnitID"))
				end if
				rs1.movenext
				if Not rs1.eof then strUnitID=strUnitID&"'"
			wend
			rs1.close
			strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
		elseif trim(Session("UnitLevelID"))="3" then
			strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitName"
		end if
		set rs1=conn.execute(strSQL)
		while Not rs1.eof
			strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
			if isEmpty(request(UnitName)) then
				if trim(rs1("UnitID"))=trim(Session("Unit_ID")) then strtmp=strtmp+" selected"
			elseif trim(rs1("UnitID"))=trim(request(UnitName)) then
				strtmp=strtmp+" selected"
			end if
			strtmp=strtmp+">"&rs1("UnitName")&"</option>"
			rs1.movenext
		wend
		rs1.close
		strtmp=strtmp+"</select>"
		
	elseif trim(session("ManagerPower"))="0" then
		strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID"
		set rs1=conn.execute(strSQL)

		if trim(MemberName)<>"" then
			strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitMan('"&UnitName&"','"&MemberName&"');"">"
		else
			strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"">"
		end if

		strtmp=strtmp+"<option Value="""&rs1("UnitID")&""">"&rs1("UnitName")&"</option>"
		strtmp=strtmp+"</select>"
		rs1.close
	end if
	SelectUnitOption=strtmp
end function
Function SelectMemberOption(UnitName,MemberName)
	strtmp=""
	if trim(session("ManagerPower"))="1" then
		strtmp="<input type=""text"" name=""chekChMemID"" size=2 onkeyup=""ListItem('"&UnitName&"','"&MemberName&"')""  class=""btn1"">"
		strtmp=strtmp+"<select name="""&MemberName&""" class=""btn1"">"
		strtmp=strtmp+"<option Value="""">所有人</option>"
		strtmp=strtmp+"</select>"
	elseif trim(session("ManagerPower"))="0" then
		strtmp="<select name="""&MemberName&""" class=""btn1"">"
		strtmp=strtmp+"<option Value="""&Session("User_ID")&""">"&Session("Ch_Name")&"</option>"
		strtmp=strtmp+"</select>"
	end if
	SelectMemberOption=strtmp
end function

Function UnSelectUnitOption(UnitName,MemberName)
	strtmp=""
	strUnitID=""
	strUnitName=""
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUnit=conn.execute(strSQL)
	If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
	rsUnit.close
	if trim(MemberName)<>"" then
		strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitMan('"&UnitName&"','"&MemberName&"');"">"
	else
		strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"">"
	end if
	if trim(Session("UnitLevelID"))="1" or Instr(strUnitName,"交通隊")>0 then
		strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder,UnitTypeID,UnitID"
		strtmp=strtmp+"<option value="""">所有單位</option>"
	elseif trim(Session("UnitLevelID"))="2" or Instr(strUnitName,"分局")>0 then
		strSQL="select UnitID,UnitName from UnitInfo where UnitTypeid in(select UnitTypeid from Unitinfo where Unitid='"&Session("Unit_ID")&"') order by UnitOrder,UnitTypeID,UnitID"

		set rs1=conn.execute(strSQL)
		while Not rs1.eof
			if trim(strUnitID)<>"" then strUnitID=trim(strUnitID)&","
			if trim(strUnitID)="" then
				strUnitID=strUnitID&trim(rs1("UnitID"))
			else
				strUnitID=strUnitID&"'"&trim(rs1("UnitID"))
			end if
			rs1.movenext
			if Not rs1.eof then strUnitID=strUnitID&"'"
		wend
		rs1.close
		strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
	elseif trim(Session("UnitLevelID"))="3" then
		strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitName"
	end if
	set rs1=conn.execute(strSQL)
	while Not rs1.eof
		strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
		if isEmpty(request(UnitName)) then
			if trim(rs1("UnitID"))=trim(Session("Unit_ID")) then strtmp=strtmp+" selected"
		elseif trim(rs1("UnitID"))=trim(request(UnitName)) then
			strtmp=strtmp+" selected"
		end if
		strtmp=strtmp+">"&rs1("UnitID")&" - "&rs1("UnitName")&"</option>"
		rs1.movenext
	wend
	rs1.close
	strtmp=strtmp+"</select>"
	UnSelectUnitOption=strtmp
end function

Function UnSelectMemberOption(UnitName,MemberName)
	strtmp="<input type=""text"" name=""chekChMemID"" id=""chekChMemID"" size=2 onkeyup=""ListItem('"&UnitName&"','"&MemberName&"')""  class=""btn1"">"
	strtmp=strtmp+"<select name="""&MemberName&""" id="""&MemberName&""" class=""btn1"">"
	strtmp=strtmp+"<option Value="""">所有人</option>"
	strtmp=strtmp+"</select>"
	UnSelectMemberOption=strtmp
end function

Function UnEventSelectUnitOption(EventUnitName,UnitName,MemberName)
	strtmp=""
	strUnitID=""
	strUnitName=""
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUnit=conn.execute(strSQL)
	If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
	rsUnit.close
	if trim(MemberName)<>"" then
		strtmp="<select name="""&EventUnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitMan('"&UnitName&"','"&MemberName&"');"">"
	else
		strtmp="<select name="""&EventUnitName&""" ID="""&UnitName&""" class=""btn1"">"
	end if
	if trim(Session("UnitLevelID"))="1" or Instr(strUnitName,"交通隊")>0 then
		strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder,UnitTypeID,UnitName"
		strtmp=strtmp+"<option value="""">所有單位</option>"
	elseif trim(Session("UnitLevelID"))="2" or Instr(strUnitName,"分局")>0 then
		strSQL="select UnitID,UnitName from UnitInfo where UnitTypeid in(select UnitTypeid from Unitinfo where Unitid='"&Session("Unit_ID")&"') order by UnitOrder,UnitTypeID,UnitName"

		set rs1=conn.execute(strSQL)
		while Not rs1.eof
			if trim(strUnitID)<>"" then strUnitID=trim(strUnitID)&","
			if trim(strUnitID)="" then
				strUnitID=strUnitID&trim(rs1("UnitID"))
			else
				strUnitID=strUnitID&"'"&trim(rs1("UnitID"))
			end if
			rs1.movenext
			if Not rs1.eof then strUnitID=strUnitID&"'"
		wend
		rs1.close
		strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
	elseif trim(Session("UnitLevelID"))="3" then
		strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitName"
	end if
	set rs1=conn.execute(strSQL)
	while Not rs1.eof
		strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
		if isEmpty(request(UnitName)) then
			if trim(rs1("UnitID"))=trim(Session("Unit_ID")) then strtmp=strtmp+" selected"
		elseif trim(rs1("UnitID"))=trim(request(UnitName)) then
			strtmp=strtmp+" selected"
		end if
		strtmp=strtmp+">"&rs1("UnitID")&" - "&rs1("UnitName")&"</option>"
		rs1.movenext
	wend
	rs1.close
	strtmp=strtmp+"</select>"
	UnEventSelectUnitOption=strtmp
end function

Function UnEventSelectMemberOption(UnitName,EventMemberName,MemberName)
	strtmp="<input type=""text"" name=""chekChMemID"" id=""chekChMemID"" size=2 onkeydown=""funTextControl(this);"" onkeyup=""ListItem('"&UnitName&"','"&MemberName&"')""  class=""btn1"">"
	strtmp=strtmp+"<select name="""&EventMemberName&""" id="""&MemberName&""" class=""btn1"">"
	strtmp=strtmp+"<option Value="""">所有人</option>"
	strtmp=strtmp+"</select>"
	UnEventSelectMemberOption=strtmp
end function

Function UnLaverSelectUnitOption(UnitName,MemberName)
	strtmp=""
	strUnitID=""
	strUnitName=""
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUnit=conn.execute(strSQL)
	If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
	rsUnit.close
	if trim(MemberName)<>"" then
		strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitLaverMan('"&UnitName&"','"&MemberName&"');"">"
	else
		strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"">"
	end if

	strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder,UnitTypeID,UnitName"
	strtmp=strtmp+"<option value="""">所有單位</option>"

	set rs1=conn.execute(strSQL)
	while Not rs1.eof
		strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
		if trim(rs1("UnitID"))=trim(request(UnitName)) then
			strtmp=strtmp+" selected"
		end if
		strtmp=strtmp+">"&rs1("UnitID")&" - "&rs1("UnitName")&"</option>"
		rs1.movenext
	wend
	rs1.close
	strtmp=strtmp+"</select>"
	UnLaverSelectUnitOption=strtmp
end function
Function UnLaverSelectMemberOption(UnitName,MemberName)
	strtmp="<input type=""text"" name=""chekLaverChMemID"" id=""chekLaverChMemID"" size=2 onkeyup=""ListItemLaver('"&UnitName&"','"&MemberName&"')""  class=""btn1"">"
	strtmp=strtmp+"<select name="""&MemberName&""" id="""&MemberName&""" class=""btn1"">"
	strtmp=strtmp+"<option Value="""">所有人</option>"
	strtmp=strtmp+"</select>"
	UnLaverSelectMemberOption=strtmp
end function

Function IfNull(strNull)
	if ISNull(strNull) or trim(strNull)="" or IsEmpty(strNull) then
		IfNull=true
	else
		IfNull=false
	end if
end function

'--------------------------------------
'檢查登入者是否是上傳人
function checkUpLoadUser(BatchNumber)
	strSQL="select RecordMemberID from DciLog where BatchNumber='"&BatchNumber&"'"

	set rs1=conn.execute(strSQL)
	if Not rs1.eof then
		if trim(rs1("RecordMemberID"))=trim(Session("User_ID")) then
			checkUpLoadUser=true
		else
			checkUpLoadUser=false
		end if
	else
		checkUpLoadUser=false
	end if
	rs1.close
end function
'-----------------------------------
'檢查能不能刪除
function checkIsAllowDel(s_City,BillType)
	if s_City="花蓮縣" then
		'交通隊, 保安隊
'		if Session("Unit_ID")="G000" or Session("Unit_ID")="A000" then
'			if Trim(Session("Credit_ID"))="A001001" then
'				checkIsAllowDel=true
'			else
'				checkIsAllowDel=false
'			end if
'		else
			if trim(rsfound("RecordMemberID"))=trim(session("User_ID")) then
				checkIsAllowDel=true
			else
				checkIsAllowDel=false
			end if
'		end if
	elseif s_City="雲林縣" then
		if trim(rsfound("RecordMemberID"))=trim(session("User_ID")) or Session("Group_ID")="200" then
			checkIsAllowDel=true
		else
			checkIsAllowDel=false
		end if
	elseif s_City="台中市" then
		if trim(rsfound("RecordMemberID"))=trim(session("User_ID")) then
			checkIsAllowDel=true
		else
			checkIsAllowDel=false
		end if
	elseif s_City="高雄縣" then
		if Session("Unit_ID")="08A7" and BillType="2" then
			if Session("SpecUser")="1" then
				checkIsAllowDel=true
			else
				checkIsAllowDel=false
			end if
		else
			if trim(rsfound("RecordMemberID"))=trim(session("User_ID")) then
				checkIsAllowDel=true
			else
				checkIsAllowDel=false
			end if
		end if
	else
		if trim(rsfound("RecordMemberID"))=trim(session("User_ID")) then
			checkIsAllowDel=true
		else
			checkIsAllowDel=false
		end if
	end if
end function

function InstrAdd(str,cnt)
	restr=""
	sumcnt=fix(len(str)/cnt)
	if len(str) mod cnt =0 then sumcnt=sumcnt-1
	If len(str)>cnt Then
		For chk=0 to sumcnt
			If chk>0 Then restr=restr&"<br>"
			restr=restr&Mid(str,chk*cnt+1,cnt)
		Next
		InstrAdd=restr
	else
		InstrAdd=str
	end if
end function

Function to_Money(strMoney)
	strUnit=split(",拾,佰,仟,萬",",")
	strNum=split(",壹,貳,參,肆,伍,陸,柒,捌,玖",",")
	ch_Money=""
	For chMon=0 to Len(strMoney)-1
		If Mid(strMoney,chMon+1,1)<>0 Then ch_Money=ch_Money&strNum(Mid(strMoney,chMon+1,1))&strUnit(Len(strMoney)-1-chMon)
	Next
	to_Money=ch_Money
End Function

Function toImageDir(nowDate)
	If Datediff("d","2013/07/01",nowDate)<0 and Datediff("d","2005/06/01",nowDate)>0 Then
		toImageDir="/Imgfix/"
	elseif Datediff("d","2016/07/01",nowDate)<0 and Datediff("d","2013/07/01",nowDate)>=0 Then
		toImageDir="/Imgfix1/"
	else
		toImageDir="/Imgfix/"
	End if
End Function

'-----檢查整本單子是不是有未建檔的-----
Function chkBillNoIsAllKeyIn(sysBillNo)
	strTestBegin="QWERTYUIOPASDFGHJKLZXCVBNM"
	strNotKeyIn=""
	StartNo=0
	BillNoTmp=right(sysBillNo,3)

	if BillNoTmp mod 50=0 then
		for i=1 to len(sysBillNo) 
			
			if InStr(strTestBegin,Mid(sysBillNo,i,1))>0 then
				StartNo=i
			end if
		next
	
		for j=1 to 49
			strSql="select * from BillBaseView where BillNo='"&left(sysBillNo,StartNo)&right("000000000"&right(sysBillNo,9-StartNo)-j,9-StartNo)&"'"
			set rsSql=conn.execute(strSql)
			if rsSql.eof then
				if strNotKeyIn="" then
					strNotKeyIn=left(sysBillNo,StartNo)&right("000000000"&right(sysBillNo,9-StartNo)-j,9-StartNo)
				else
					strNotKeyIn=strNotKeyIn&","&left(sysBillNo,StartNo)&right("000000000"&right(sysBillNo,9-StartNo)-j,9-StartNo)
				end if
			end if
			rsSql.close
			set rsSql=nothing
		next
	end if

	chkBillNoIsAllKeyIn=strNotKeyIn
End Function

Function MidX(str,scn,ecn)
	strmid="":uscn=0:uecn=0:lenx=0
	If not ifnull(str) Then
		For i=1 to len(str)
			If asc(mid(str,i,1))>0 Then
				lenx=lenx+1
			else
				lenx=lenx+2
			End if
			If lenx>=scn and uscn=0 Then
				uscn=i
			End if
			If lenx>=scn+ecn and uecn=0 Then
				uecn=i-uscn
				If uecn<=0 Then uecn=1
				exit for
			End if
		Next
		If uscn=0 and uecn=0 Then
			MidX=strmid
		else
			Midx=Mid(str,uscn,uecn)
		End if
	else
		MidX=strmid
	End if
End Function

'-----------NPA 單位層級-----
function funcNpaUnitLevelSelect(UnitID1,UnitID2,UnitID3,UnitID4,UnitID5,UserMemID)
	strUnit1="select * from E0DT_Unit where E0_Unit_CD='"&trim(UserMemID)&"'"
	set rsUnit1=conn.execute(strUnit1)
	if not rsUnit1.eof then
		UserUnitLevel=trim(rsUnit1("E0_Unit_Level"))
		UserDeptUnitID=trim(rsUnit1("E0_Dept_CD"))
		UserBranchUnitID=trim(rsUnit1("E0_Branch_CD"))
	end if
	rsUnit1.close
	'最終單位代碼
	response.write "<input type=""hidden"" name="""&UnitID5&""" value="""&UserMemID&""">"
	'第一層
	if UserUnitLevel="1" then
		response.write "<select name="""&UnitID1&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','1')"">"
		response.write "<option value="""">請選擇</option>"
		strU1="select * from E0DT_Unit where E0_Unit_Level='1'"
		set rsU1=conn.execute(strU1)
		If Not rsU1.Bof Then rsU1.MoveFirst 
		While Not rsU1.Eof
			response.write "<option value="""&trim(rsU1("E0_Unit_CD"))&""" "
			if trim(rsU1("E0_Unit_CD"))=trim(UserMemID) then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU1("E0_Unit_NM"))&"</option>"
		rsU1.MoveNext
		Wend
		rsU1.close
		set rsU1=nothing
		response.write "</select>"
	else
		'response.write "<input type=""hidden"" name="""&UnitID1&""" value=""A000"">"
		response.write "<select name="""&UnitID1&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','1')"">"
		response.write "<option value="""">請選擇</option>"
		strU1="select * from E0DT_Unit where E0_Unit_Level='1'"
		set rsU1=conn.execute(strU1)
		If Not rsU1.Bof Then rsU1.MoveFirst 
		While Not rsU1.Eof
			response.write "<option value="""&trim(rsU1("E0_Unit_CD"))&""" "
			if trim(rsU1("E0_Unit_CD"))="A000" then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU1("E0_Unit_NM"))&"</option>"
		rsU1.MoveNext
		Wend
		rsU1.close
		set rsU1=nothing
		response.write "</select>"
	end if

	'第二層
	if UserUnitLevel="1" then
		response.write "<select name="""&UnitID2&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','2')"">"
		response.write "<option value="""">請選擇</option>"
		strU2="select * from E0DT_Unit where E0_Unit_Level='2'"
		set rsU2=conn.execute(strU2)
		If Not rsU2.Bof Then rsU2.MoveFirst 
		While Not rsU2.Eof
			response.write "<option value="""&trim(rsU2("E0_Unit_CD"))&""" "
			response.write ">"
			response.write trim(rsU2("E0_Unit_NM"))&"</option>"
		rsU2.MoveNext
		Wend
		rsU2.close
		set rsU2=nothing
		response.write "</select>"
	elseif UserUnitLevel="2" then
		response.write "<select name="""&UnitID2&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','2')"">"
		response.write "<option value="""">請選擇</option>"
		strU3="select * from E0DT_Unit where E0_Unit_Level='2' and E0_Dept_CD='"&UserDeptUnitID&"'"
		set rsU3=conn.execute(strU3)
		If Not rsU3.Bof Then rsU3.MoveFirst 
		While Not rsU3.Eof
			response.write "<option value="""&trim(rsU3("E0_Unit_CD"))&""" "
			if trim(rsU3("E0_Unit_CD"))=trim(UserMemID) then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU3("E0_Unit_NM"))&"</option>"
		rsU3.MoveNext
		Wend
		rsU3.close
		set rsU3=nothing
		response.write "</select>"
	else
		'response.write "<input type=""hidden"" name="""&UnitID2&""" value="""&UserDeptUnitID&""">"
		response.write "<select name="""&UnitID2&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','2')"">"
		response.write "<option value="""">請選擇</option>"
		strU3="select * from E0DT_Unit where E0_Unit_Level='2' and E0_Dept_CD='"&UserDeptUnitID&"'"
		set rsU3=conn.execute(strU3)
		If Not rsU3.Bof Then rsU3.MoveFirst 
		While Not rsU3.Eof
			response.write "<option value="""&trim(rsU3("E0_Unit_CD"))&""" "
			if trim(rsU3("E0_Unit_CD"))=UserDeptUnitID then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU3("E0_Unit_NM"))&"</option>"
		rsU3.MoveNext
		Wend
		rsU3.close
		set rsU3=nothing
		response.write "</select>"
	end if

	response.write "<br>"
	'第三層
	if UserUnitLevel="1" then
		response.write "<select name="""&UnitID3&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','3')"">"
		response.write "<option value="""">請選擇</option>"
		response.write "</select>"
	elseif UserUnitLevel="2" then
		response.write "<select name="""&UnitID3&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','3')"">"
		response.write "<option value="""">請選擇</option>"
		strU4="select * from E0DT_Unit where E0_Unit_Level='3' and E0_Dept_CD='"&trim(UserMemID)&"'" 
		set rsU4=conn.execute(strU4)
		If Not rsU4.Bof Then rsU4.MoveFirst 
		While Not rsU4.Eof
			response.write "<option value="""&trim(rsU4("E0_Unit_CD"))&""" "
			response.write ">"
			response.write trim(rsU4("E0_Unit_NM"))&"</option>"
		rsU4.MoveNext
		Wend
		rsU4.close
		set rsU4=nothing
		response.write "</select>"
	elseif UserUnitLevel="3" then
		response.write "<select name="""&UnitID3&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','3')"">"
		response.write "<option value="""">請選擇</option>"
		strU5="select * from E0DT_Unit where E0_Unit_Level='3' and E0_Dept_CD='"&UserDeptUnitID&"' " &_
			" and E0_Branch_CD='"&UserBranchUnitID&"'"
		set rsU5=conn.execute(strU5)
		If Not rsU5.Bof Then rsU5.MoveFirst 
		While Not rsU5.Eof
			response.write "<option value="""&trim(rsU5("E0_Unit_CD"))&""" "
			if trim(rsU5("E0_Unit_CD"))=trim(UserMemID) then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU5("E0_Unit_NM"))&"</option>"
		rsU5.MoveNext
		Wend
		rsU5.close
		set rsU5=nothing
		response.write "</select>"
	else
		'response.write "<input type=""hidden"" name="""&UnitID3&""" value="""&UserBranchUnitID&""">"
		response.write "<select name="""&UnitID3&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','3')"">"
		response.write "<option value="""">請選擇</option>"
		strU5="select * from E0DT_Unit where E0_Unit_Level='3' and E0_Dept_CD='"&UserDeptUnitID&"' " 
		set rsU5=conn.execute(strU5)
		If Not rsU5.Bof Then rsU5.MoveFirst 
		While Not rsU5.Eof
			response.write "<option value="""&trim(rsU5("E0_Unit_CD"))&""" "
			if trim(rsU5("E0_Unit_CD"))=UserBranchUnitID then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU5("E0_Unit_NM"))&"</option>"
		rsU5.MoveNext
		Wend
		rsU5.close
		set rsU5=nothing
		response.write "</select>"
	end if
	'第四層
	if UserUnitLevel="1" or UserUnitLevel="2" then
		response.write "<select name="""&UnitID4&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','4')"">"
		response.write "<option value="""">請選擇</option>"
		response.write "</select>"
	elseif UserUnitLevel="3" then
		response.write "<select name="""&UnitID4&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','4')"">"
		response.write "<option value="""">請選擇</option>"
		strU6="select * from E0DT_Unit where E0_Unit_Level='4' and E0_Dept_CD='"&UserDeptUnitID&"' " &_
			" and E0_Branch_CD='"&trim(UserMemID)&"'"
		set rsU6=conn.execute(strU6)
		If Not rsU6.Bof Then rsU6.MoveFirst 
		While Not rsU6.Eof
			response.write "<option value="""&trim(rsU6("E0_Unit_CD"))&""" "
			response.write ">"
			response.write trim(rsU6("E0_Unit_NM"))&"</option>"
		rsU6.MoveNext
		Wend
		rsU6.close
		set rsU6=nothing
		response.write "</select>"
	else
		response.write "<select name="""&UnitID4&""" onchange=""NpaUnitLevelSelectAJAX('"&UnitID1&"','"&UnitID2&"','"&UnitID3&"','"&UnitID4&"','"&UnitID5&"','4')"">"
		response.write "<option value="""">請選擇</option>"
'		strU6="select * from E0DT_Unit where E0_Unit_Level='4' and E0_Dept_CD='"&UserDeptUnitID&"' " &_
'			" and E0_Branch_CD='"&UserBranchUnitID&"' and E0_Unit_CD='"&trim(UserMemID)&"'"
		strU6="select * from E0DT_Unit where E0_Unit_Level='4' and E0_Dept_CD='"&UserDeptUnitID&"' " &_
			" and E0_Branch_CD='"&UserBranchUnitID&"'"
		set rsU6=conn.execute(strU6)
		If Not rsU6.Bof Then rsU6.MoveFirst 
		While Not rsU6.Eof
			response.write "<option value="""&trim(rsU6("E0_Unit_CD"))&""" "
			if trim(rsU6("E0_Unit_CD"))=trim(UserMemID) then
				response.write "selected"
			end if
			response.write ">"
			response.write trim(rsU6("E0_Unit_NM"))&"</option>"
		rsU6.MoveNext
		Wend
		rsU6.close
		set rsU6=nothing
		response.write "</select>"
	end if
end function

function funcNpaUnitLevelSelectAJAX()
	response.write "function NpaUnitLevelSelectAJAX(UnitID1,UnitID2,UnitID3,UnitID4,UnitID5,UnitType){"
	response.write "runServerScript(""/traffic/Common/TD09A01M.asp?UnitID1=""+UnitID1+""&UnitID2=""+UnitID2+""&UnitID3=""+UnitID3+""&UnitID4=""+UnitID4+""&UnitID5=""+UnitID5+""&UnitType=""+UnitType+""&UnitValue1=""+eval(""myForm.""+UnitID1+"".value"")+""&UnitValue2=""+eval(""myForm.""+UnitID2+"".value"")+""&UnitValue3=""+eval(""myForm.""+UnitID3+"".value"")+""&UnitValue4=""+eval(""myForm.""+UnitID4+"".value""));"
	response.write "}"
end function

function funcCheckFont(strFont,strSize,strFILTER)
	if instr(strFont,"@@")>0 then
		arrFont=split(" "&strFont&" ","@@")
		strTmp=""
		for FontRoop=1 to ubound(arrFont)+1
			if FontRoop mod 2 =0 then
				if strFILTER="1" then	'1正常
					strTmp=strTmp&"<img src='/traffic/FONTPIC/"&trim(arrFont(FontRoop-1))&"' width='"&strSize&"'>"
				elseif strFILTER="2" then	'2轉90度
					strTmp=strTmp&"<img style=""FILTER: progid:DXImageTransform.Microsoft.BasicImage(rotation=3)"" src='/traffic/FONTPIC/"&trim(arrFont(FontRoop-1))&"' width='"&strSize&"'  border=0>"
				elseif strFILTER="3" then	'3轉270度
					strTmp=strTmp&"<img style=""FILTER: progid:DXImageTransform.Microsoft.BasicImage(rotation=1)"" src='/traffic/FONTPIC/"&trim(arrFont(FontRoop-1))&"' width='"&strSize&"'  border=0>"
				elseif strFILTER="0" then	'變空白
					strTmp=strTmp&" &nbsp; "
				elseif strFILTER="4" then	'變空白 轉V
					strTmp=strTmp&"<img src='/traffic/FONTPIC_H/"&trim(arrFont(FontRoop-1))&"' width='"&strSize&"'>"
				end if
			else
				strTmp=strTmp&trim(arrFont(FontRoop-1))
			end if
		next

		funcCheckFont=strTmp
	else
		funcCheckFont=replace(strFont&""," ","　")
	end if
end function

function getZip(straddr)
	zipid="null"
	if not ifnull(straddr) then
		straddr=replace(straddr,"臺","台")
		straddr=replace(straddr,"台北縣","新北市")
		straddr=replace(straddr,"台中縣","台中市")
		straddr=replace(straddr,"台南縣","台南市")
		straddr=replace(straddr,"高雄縣","高雄市")

		If instr(straddr,"台中市")>0 or instr(straddr,"台南市")>0 then
			straddr=replace(straddr,"區","市")
			straddr=replace(straddr,"區","鎮")
			straddr=replace(straddr,"區","村")
		end if

		If instr(straddr,"新北市")>0 or instr(straddr,"高雄市")>0 then
			straddr=replace(straddr,"區","市")
			straddr=replace(straddr,"區","鎮")
			straddr=replace(straddr,"區","村")
		end if

		strSQL = "select ZipID,ZipName from Zip where ZipName like '"&left(trim(straddr),5)&"%'"

		set rszip=conn.execute(strSQL)
		If not rszip.eof Then
			zipid=rszip("zipid")
		else
			rszip.close
			
			strSQL = "select ZipID,ZipName from Zip where ZipName like '"&left(trim(straddr),3)&"%'"
			set rszip=conn.execute(strSQL)
			if Not rszip.eof then zipid=rszip("zipid")
		end if
		rszip.close
	end if
	if not ifnull(zipid) then
		getZip=zipid
	else
		getZip="null"
	end if
end function

function ChangeOldCity(Sys_OwnerZip,Sys_OwnerZipName)
	If cdbl(Sys_OwnerZip)> 408 and cdbl(Sys_OwnerZip)< 440 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台中市","台中縣")

		If Sys_OwnerZip=411 or Sys_OwnerZip=412 or Sys_OwnerZip=420 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=423 or Sys_OwnerZip=433 or Sys_OwnerZip=435 or Sys_OwnerZip=436 or Sys_OwnerZip=437 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if
	elseIf cdbl(Sys_OwnerZip)> 709 and cdbl(Sys_OwnerZip)< 746 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台南市","台南縣")

		If Sys_OwnerZip=710 or Sys_OwnerZip=730 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=712 or Sys_OwnerZip=721 or Sys_OwnerZip=722 or Sys_OwnerZip=726 or Sys_OwnerZip=732 or Sys_OwnerZip=737 or Sys_OwnerZip=741 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if

	elseIf cdbl(Sys_OwnerZip)> 813 and cdbl(Sys_OwnerZip)< 853 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"高雄市","高雄縣")

		If Sys_OwnerZip=830 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=820 or Sys_OwnerZip=842 or Sys_OwnerZip=843 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if

	elseIf cdbl(Sys_OwnerZip)> 206 and cdbl(Sys_OwnerZip)< 254 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"新北市","台北縣")

		If Sys_OwnerZip=220 or Sys_OwnerZip=221 or Sys_OwnerZip=231 or Sys_OwnerZip=234 or Sys_OwnerZip=235 or Sys_OwnerZip=236 or Sys_OwnerZip=238 or Sys_OwnerZip=241 or Sys_OwnerZip=242 or Sys_OwnerZip=247 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","市")

		elseif Sys_OwnerZip=224 or Sys_OwnerZip=237 or Sys_OwnerZip=239 or Sys_OwnerZip=251 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","鎮")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"區","村")
		End if
		
	end if
	ChangeOldCity=Sys_OwnerZipName
end function

function ChangeNewCity(Sys_OwnerZip,Sys_OwnerZipName)
	if not ifnull(Sys_OwnerZipName) then
		Sys_OwnerZipName=replace(Sys_OwnerZipName,"臺","台")
	end if
	If cdbl(Sys_OwnerZip)> 408 and cdbl(Sys_OwnerZip)< 440 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台中縣","台中市")

		If Sys_OwnerZip=411 or Sys_OwnerZip=412 or Sys_OwnerZip=420 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"市","區")

		elseif Sys_OwnerZip=423 or Sys_OwnerZip=433 or Sys_OwnerZip=435 or Sys_OwnerZip=436 or Sys_OwnerZip=437 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"鎮","區")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"村","區")
		End if
	elseIf cdbl(Sys_OwnerZip)> 709 and cdbl(Sys_OwnerZip)< 746 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台南縣","台南市")

		If Sys_OwnerZip=710 or Sys_OwnerZip=730 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"市","區")

		elseif Sys_OwnerZip=712 or Sys_OwnerZip=721 or Sys_OwnerZip=722 or Sys_OwnerZip=726 or Sys_OwnerZip=732 or Sys_OwnerZip=737 or Sys_OwnerZip=741 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"鎮","區")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"村","區")
		End if

	elseIf cdbl(Sys_OwnerZip)> 813 and cdbl(Sys_OwnerZip)< 853 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"高雄縣","高雄市")

		If Sys_OwnerZip=830 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"市","區")

		elseif Sys_OwnerZip=820 or Sys_OwnerZip=842 or Sys_OwnerZip=843 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"鎮","區")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"村","區")
		End if

	elseIf cdbl(Sys_OwnerZip)> 206 and cdbl(Sys_OwnerZip)< 254 then

		Sys_OwnerZipName=replace(Sys_OwnerZipName,"台北縣","新北市")

		If Sys_OwnerZip=220 or Sys_OwnerZip=221 or Sys_OwnerZip=231 or Sys_OwnerZip=234 or Sys_OwnerZip=235 or Sys_OwnerZip=236 or Sys_OwnerZip=238 or Sys_OwnerZip=241 or Sys_OwnerZip=242 or Sys_OwnerZip=247 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"市","區")

		elseif Sys_OwnerZip=224 or Sys_OwnerZip=237 or Sys_OwnerZip=239 or Sys_OwnerZip=251 Then
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"鎮","區")

		else
			Sys_OwnerZipName=replace(Sys_OwnerZipName,"村","區")
		End if
		
	end if
	ChangeNewCity=Sys_OwnerZipName
end function

Function chstr(istr) ' 半形轉全形
Dim strtmp
min=0:max=0
If instr(istr,"<") >0 Then
	min=instr(istr,"<")
	max=instr(istr,">")-min+1
	img_word=mid(istr,min,max)
	istr=replace(istr,img_word,"")
End if
if trim(istr)<>"" then
	strtmp = Replace(istr, "(", "（") 
	strtmp = Replace(strtmp, ")", "）") 
	strtmp = Replace(strtmp, "[", "〔") 
	strtmp = Replace(strtmp, "]", "〕") 
	strtmp = Replace(strtmp, "{", "｛") 
	strtmp = Replace(strtmp, "}", "｝") 
	strtmp = Replace(strtmp, ".", "。") 
	strtmp = Replace(strtmp, ",", "，") 
	strtmp = Replace(strtmp, ";", "；") 
	strtmp = Replace(strtmp, ":", "：") 
	strtmp = Replace(strtmp, "-", "－") 
	strtmp = Replace(strtmp, "?", "？") 
	strtmp = Replace(strtmp, "!", "！") 
	strtmp = Replace(strtmp, "@", "＠") 
	strtmp = Replace(strtmp, "#", "＃") 
	strtmp = Replace(strtmp, "$", "＄") 
	strtmp = Replace(strtmp, "%", "％") 
	strtmp = Replace(strtmp, "&", "＆") 
	strtmp = Replace(strtmp, "|", "｜") 
	strtmp = Replace(strtmp, "", "＼") 
	strtmp = Replace(strtmp, "/", "／") 
	strtmp = Replace(strtmp, "+", "＋") 
	strtmp = Replace(strtmp, "=", "＝") 
	strtmp = Replace(strtmp, "*", "＊") 
	strtmp = Replace(strtmp, "0", "０") 
	strtmp = Replace(strtmp, "1", "１") 
	strtmp = Replace(strtmp, "2", "２") 
	strtmp = Replace(strtmp, "3", "３") 
	strtmp = Replace(strtmp, "4", "４") 
	strtmp = Replace(strtmp, "5", "５") 
	strtmp = Replace(strtmp, "6", "６") 
	strtmp = Replace(strtmp, "7", "７") 
	strtmp = Replace(strtmp, "8", "８") 
	strtmp = Replace(strtmp, "9", "９") 
	strtmp = Replace(strtmp, "a", "ａ") 
	strtmp = Replace(strtmp, "b", "ｂ") 
	strtmp = Replace(strtmp, "c", "ｃ") 
	strtmp = Replace(strtmp, "d", "ｄ") 
	strtmp = Replace(strtmp, "e", "ｅ") 
	strtmp = Replace(strtmp, "f", "ｆ") 
	strtmp = Replace(strtmp, "g", "ｇ") 
	strtmp = Replace(strtmp, "h", "ｈ") 
	strtmp = Replace(strtmp, "i", "ｉ") 
	strtmp = Replace(strtmp, "j", "ｊ") 
	strtmp = Replace(strtmp, "k", "ｋ") 
	strtmp = Replace(strtmp, "l", "ｌ") 
	strtmp = Replace(strtmp, "m", "ｍ") 
	strtmp = Replace(strtmp, "n", "ｎ") 
	strtmp = Replace(strtmp, "o", "ｏ") 
	strtmp = Replace(strtmp, "p", "ｐ") 
	strtmp = Replace(strtmp, "q", "ｑ") 
	strtmp = Replace(strtmp, "r", "ｒ") 
	strtmp = Replace(strtmp, "s", "ｓ") 
	strtmp = Replace(strtmp, "t", "ｔ") 
	strtmp = Replace(strtmp, "u", "ｕ") 
	strtmp = Replace(strtmp, "v", "ｖ") 
	strtmp = Replace(strtmp, "w", "ｗ") 
	strtmp = Replace(strtmp, "x", "ｘ") 
	strtmp = Replace(strtmp, "y", "ｙ") 
	strtmp = Replace(strtmp, "z", "ｚ") 
	strtmp = Replace(strtmp, "A", "Ａ") 
	strtmp = Replace(strtmp, "B", "Ｂ") 
	strtmp = Replace(strtmp, "C", "Ｃ") 
	strtmp = Replace(strtmp, "D", "Ｄ") 
	strtmp = Replace(strtmp, "E", "Ｅ") 
	strtmp = Replace(strtmp, "F", "Ｆ") 
	strtmp = Replace(strtmp, "G", "Ｇ") 
	strtmp = Replace(strtmp, "H", "Ｈ") 
	strtmp = Replace(strtmp, "I", "Ｉ") 
	strtmp = Replace(strtmp, "J", "Ｊ") 
	strtmp = Replace(strtmp, "K", "Ｋ") 
	strtmp = Replace(strtmp, "L", "Ｌ") 
	strtmp = Replace(strtmp, "M", "Ｍ") 
	strtmp = Replace(strtmp, "N", "Ｎ") 
	strtmp = Replace(strtmp, "O", "Ｏ") 
	strtmp = Replace(strtmp, "P", "Ｐ") 
	strtmp = Replace(strtmp, "Q", "Ｑ") 
	strtmp = Replace(strtmp, "R", "Ｒ") 
	strtmp = Replace(strtmp, "S", "Ｓ") 
	strtmp = Replace(strtmp, "T", "Ｔ") 
	strtmp = Replace(strtmp, "U", "Ｕ") 
	strtmp = Replace(strtmp, "V", "Ｖ") 
	strtmp = Replace(strtmp, "W", "Ｗ") 
	strtmp = Replace(strtmp, "X", "Ｘ") 
	strtmp = Replace(strtmp, "Y", "Ｙ") 
	strtmp = Replace(strtmp, "Z", "Ｚ") 
	strtmp = Replace(strtmp, " ", "　")
end if

If min >0 Then strtmp=mid(strtmp,1,min-1)&img_word&mid(strtmp,min)

chstr = strtmp 
End Function

function decrypt(text) 
	if trim(text & "")="" then
		decrypt=trim(text)
	else
		const key="HD" 
		dim str : str=key & text 
		dim str1
		dim str2
		for i=1 to len(str)
			
			str2=hex(AscW(mid(str,i,1)))
			for j=1 to 4-len(str2)
				str2="0" & str2 
			next
			str1=str1 & str2
		next
		decrypt=strreverse(str1) 
	end if 
end function

function encrypt(text)
	if trim(text & "")="" then
		encrypt=trim(text)
	else
		const key="HD" 
		dim str : str=text
		dim str1
		dim str2 : str2=strreverse(str) 
		for i=1 to len(str2) step 4 
			str1=str1 & ChrW(Cint("&H" & mid(str2,i,4)))
		next
		encrypt=mid(str1,len(key)+1,len(str)-len(key))
	end if
end function

TaiChungNotInUnitID=" and BillUnitID not in ('9000','4E11','Z001')"
%>
