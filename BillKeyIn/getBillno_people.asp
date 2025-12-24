<%
function funcBillToDCICaseIn()
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)
	
	'若案件是逕舉且無單號則產生單號

	if sys_City="高雄縣" or sys_City="花蓮縣" or sys_City="高雄市" or sys_City="彰化縣" Or sys_City=ApconfigureCityName then	'依單位分別取號
		if sys_City="高雄市" then
			'thirdNo=right(year(date)-1911,1)
			thirdNo="H"
			thirdNoSubUnit="D"
			'response.write thirdNo
			if trim(Session("Credit_ID"))="T220933992" then
				UserStartNo="BD"&thirdNo&"0"
				UserSeq="bill0807BD0"
			elseif trim(Session("Credit_ID"))="E223625931" then
				UserStartNo="BD"&thirdNo&"3"
				UserSeq="bill0807BD3"
			elseif trim(Session("Credit_ID"))="E121011955" then
				UserStartNo="BD"&thirdNo&"5"
				UserSeq="bill0807BD5"
			elseif trim(Session("Credit_ID"))="E120003931" then
				UserStartNo="BB"&thirdNo&"0"
				UserSeq="bill0807BB0"
			elseif trim(Session("Credit_ID"))="T220359567" then
				UserStartNo="BB"&thirdNo&"3"
				UserSeq="bill0807BB3"
			elseif trim(Session("Credit_ID"))="E220912204" then
				UserStartNo="BB"&thirdNo&"5"
				UserSeq="bill0807BB5"				
			elseif trim(Session("Credit_ID"))="T220988040" then
				UserStartNo="BB"&thirdNo&"7"
				UserSeq="bill0807BB7"
			elseif trim(Session("Credit_ID"))="S220060233" then
				UserStartNo="BB"&thirdNo&"9"
				UserSeq="bill0807BB9"
			elseif trim(Session("Credit_ID"))="E220182233" then
				UserStartNo="BC"&thirdNo&"0"
				UserSeq="bill0807BC0"
			elseif trim(Session("Credit_ID"))="T121457177" then
				UserStartNo="BC"&thirdNo&"3"
				UserSeq="bill0807BC3"
			elseif trim(Session("Credit_ID"))="R120228634" Or trim(Session("Credit_ID"))="E221201933" Or trim(Session("Credit_ID"))="S221552347" then
				UserStartNo="BC"&thirdNo&"5"
				UserSeq="bill0807BC5"
			else
				if trim(Session("Unit_ID"))="0807" then
					sysUserUnit="0807"
				else
					strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
					set rsUserUnit=conn.execute(strUserUnit)
					if not rsUserUnit.eof then
						sysUserUnit=trim(rsUserUnit("UnitTypeID"))
					end if
					rsUserUnit.close
					set rsUserUnit=nothing
				end if
				'抓 起始碼 跟 Seq Name
				strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
				set rsSeq=conn.execute(strSeq)
				if not rsSeq.eof Then
					If Len(trim(rsSeq("BillStartVocab")))=3 then
						UserStartNo=trim(rsSeq("BillStartVocab"))
					Else
						UserStartNo=trim(rsSeq("BillStartVocab"))&thirdNoSubUnit
					End If 
					UserSeq=trim(rsSeq("SeqNoName"))
				end if
				rsSeq.close
				set rsSeq=nothing	
			end If
		elseif sys_City="花蓮縣" Then
			if trim(Session("Credit_ID"))="A06" Or trim(Session("Credit_ID"))="A07" Then
				UserStartNo="PB"
				UserSeq="billA06PB"
			Else
				strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
				set rsUserUnit=conn.execute(strUserUnit)
				if not rsUserUnit.eof then
					sysUserUnit=trim(rsUserUnit("UnitTypeID"))
				end if
				rsUserUnit.close
				set rsUserUnit=Nothing

				'抓 起始碼 跟 Seq Name
				strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
				set rsSeq=conn.execute(strSeq)
				if not rsSeq.eof then
					UserSeq=trim(rsSeq("SeqNoName"))
					UserStartNo=trim(rsSeq("BillStartVocab"))
				end if
				rsSeq.close
				set rsSeq=nothing
			End If 
		else
			if sys_City="高雄縣" then
				if Session("Unit_ID")="8H00" then
					sysUserUnit="8H00"
				else
					strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
					set rsUserUnit=conn.execute(strUserUnit)
					if not rsUserUnit.eof then
						sysUserUnit=trim(rsUserUnit("UnitTypeID"))
					end if
					rsUserUnit.close
					set rsUserUnit=nothing
				end if
			elseif sys_City="彰化縣" then
				strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
				set rsUserUnit=conn.execute(strUserUnit)
				if not rsUserUnit.eof then
					sysUserUnit=trim(rsUserUnit("UnitTypeID"))
				end if
				rsUserUnit.close
				set rsUserUnit=Nothing
			elseif sys_City=ApconfigureCityName then

					strUserUnit="select UnitID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
					set rsUserUnit=conn.execute(strUserUnit)
					if not rsUserUnit.eof then
						sysUserUnit=trim(rsUserUnit("UnitID"))
					end if
					rsUserUnit.close
					set rsUserUnit=nothing

			end if
			'抓 起始碼 跟 Seq Name
			strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
			set rsSeq=conn.execute(strSeq)
			if not rsSeq.eof then
				UserSeq=trim(rsSeq("SeqNoName"))
				UserStartNo=trim(rsSeq("BillStartVocab"))
			end if
			rsSeq.close
			set rsSeq=nothing
		end if

		'抓單號最大值
		if len(UserStartNo)=2 then
			strSQLSn = "select LPAD("&UserSeq&".nextval,7,'0') as SN from Dual"
		elseif len(UserStartNo)=3 then
			strSQLSn = "select LPAD("&UserSeq&".nextval,6,'0') as SN from Dual"
		elseif len(UserStartNo)=4 then
			strSQLSn = "select LPAD("&UserSeq&".nextval,5,'0') as SN from Dual"
		elseif len(UserStartNo)=5 then
			strSQLSn = "select LPAD("&UserSeq&".nextval,4,'0') as SN from Dual"
		end if
		set rsBNo = Conn.execute(strSQLSn)
		if not rsBNo.EOF then
			sMaxSN = rsBNo("SN")
		end if
		rsBNo.close
		set rsBNo = nothing

		NewBillNo=UserStartNo&sMaxSN

	else
		'到ApConfigure抓縣市代碼 ID=2
		CityID=""
		strCity="select Value from ApConfigure where ID=2"
		set rsCity=conn.execute(strCity)
		if not rsCity.eof then
			CityID=trim(rsCity("Value"))
		end if
		rsCity.close
		set rsCity=Nothing

		strSQLSn = "select LPAD(test_sn.nextval,6,'0') as SN from Dual"
		set rsBNo = Conn.execute(strSQLSn)
		if not rsBNo.EOF then
			sMaxSN = rsBNo("SN")
		end if
		rsBNo.close
		set rsBNo = nothing
		'到ApConfigure抓單號第三碼 ID=39
		BillNo3=""
		strBillNo3="select Value from ApConfigure where ID=39"
		set rsBillNo3=conn.execute(strBillNo3)
		if not rsBillNo3.eof then
			BillNo3=trim(rsBillNo3("Value"))
		end if
		rsBillNo3.close
		set rsBillNo3=nothing
		NewBillNo=CityID&BillNo3&sMaxSN
		if sMaxSN="999999" then
			NewBillNo3=""
			if ASC(BillNo3)>47 and ASC(BillNo3)<57 then	'0~8
				NewBillNo3=Chr(ASC(BillNo3)+1)
			elseif ASC(BillNo3)=57 then	'9
				NewBillNo3="A"
			elseif ASC(BillNo3)>64 and ASC(BillNo3)<90 then	'A~Y
				NewBillNo3=Chr(ASC(BillNo3)+1)
			elseif ASC(BillNo3)=90 then	'Z
				NewBillNo3="0"
			end if
			strUpdBillNo3="update Apconfigure set Value='"&NewBillNo3&"' where ID=39"
			conn.execute strUpdBillNo3
		end If

	end if

	funcBillToDCICaseIn=NewBillNo
end function
%>