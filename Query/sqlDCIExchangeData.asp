<%
'conn=Connection名稱
'車籍查詢
function funcCarDataCheck(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'更新該筆紀錄的 BILLSTATUS 更新為 1
	strUpdCar="Update BillBase set billstatus='1' where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	strInsCar="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'A','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
	conn.BeginTrans
		conn.execute strUpdCar
		conn.execute strInsCar
	if err.number = 0 then
       conn.CommitTrans
    else            
       conn.RollbackTrans
    end if   
end function

'入案
function funcBillToDCICaseIn(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime,sysCity)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)
	strDUnit="select DciUnitID from UnitInfo where UnitID='"&trim(BillUnitID)&"'"
		set rsDUnit=conn.execute(strDUnit)
		if not rsDUnit.eof then
			if trim(rsDUnit("DciUnitID"))="" or isnull(trim(rsDUnit("DciUnitID"))) then
				theDciUnitID=BillUnitID
			else
				theDciUnitID=trim(rsDUnit("DciUnitID"))
			end if
		else
			theDciUnitID=BillUnitID
		end if
	'若案件是逕舉且無單號則產生單號
	if BillTypeID="2" and (isnull(BillNo) or trim(BillNo)="") then
		if sysCity="高雄縣" or sysCity="花蓮縣" or sys_City="高雄市" or sys_City="彰化縣" Or sys_City=ApconfigureCityName then	'依單位分別取號
			if sysCity="高雄市" then
				'thirdNo=right(year(date)-1911,1)
				thirdNo="G"
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
			elseif sysCity="花蓮縣" Then
				if trim(Session("Credit_ID"))="A06" Then
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
				if sysCity="高雄縣" then
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
				elseif sysCity="彰化縣" then
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

			strUpdSN="update BillBase set BillNo='"&NewBillNo&"' where SN="&BillSN
			conn.execute strUpdSN
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
'			If sysCity="台中市" Then
'				'抓單號最大值
'				strSQLSn = "select LPAD(test_sn.nextval,5,'0') as SN from Dual"
'				set rsBNo = Conn.execute(strSQLSn)
'				if not rsBNo.EOF then
'					sMaxSN = rsBNo("SN")
'				end if
'				rsBNo.close
'				set rsBNo = nothing
'				'到ApConfigure抓單號第三碼 ID=39
'				BillNo3=""
'				strBillNo3="select Value from ApConfigure where ID=39"
'				set rsBillNo3=conn.execute(strBillNo3)
'				if not rsBillNo3.eof then
'					BillNo3=trim(rsBillNo3("Value"))
'				end if
'				rsBillNo3.close
'				set rsBillNo3=nothing
'				NewBillNo=CityID&BillNo3&"A"&sMaxSN
'			else
				'抓單號最大值
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
			'End If 
			strUpdSN="update BillBase set BillNo='"&NewBillNo&"' where SN="&BillSN
			conn.execute strUpdSN
		end if
	end if
	if BillTypeID="2" and (isnull(BillNo) or trim(BillNo)="") then
		strSN="select BillNo from BillBase where SN="&BillSN
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			BillNotemp=trim(rsSN("BillNo"))
		end if
		rsSN.close
		set rsSN=nothing
	else
		BillNotemp=BillNo
	end if
	'把select 出來的紀錄寫入到DCILog
	strInsCaseIn="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,DCIwindowName,BatchNumber,DciUnitID)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNotemp&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'W','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		",'"&theDciUnitID&"')" 
	conn.execute strInsCaseIn

	'更新該筆紀錄的 BILLSTATUS 更新為 2
	if sysCity="高雄市" Or sysCity="苗栗縣" then
		strUpdCaseIn="Update BillBase set billstatus='2',caseindate=sysdate where SN="&BillSN
	else
		strUpdCaseIn="Update BillBase set billstatus='2' where SN="&BillSN
	end if
	conn.execute strUpdCaseIn

	'新增BillMailHistory
	strMailChk="select * from BillMailHistory where BillSN="&BillSN
	set rsMailChk=conn.execute(strMailChk)
	if rsMailChk.eof then
		'if BillTypeID="2" then
			strMail="Insert into BillMailHistory(BillSN,BillNo,CarNo,MailDate,MailNumber) "&_
				" values("&BillSN&",'"&BillNotemp&"','"&CarNo&"',null,null)"		
			conn.execute strMail
		'end if
	end if
	rsMailChk.close
	set rsMailChk=nothing
	funcBillToDCICaseIn=BillNotemp
end function

'退件
function funcBillReturn(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'更新該筆紀錄的 BILLSTATUS 更新為 3
	strUpdRet="Update BillBase set billstatus='3' where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	'UpLoadTimeTmp=DateAdd("n",5,now)
	UpLoadTimeTmp=now
	UpLoadTime=year(UpLoadTimeTmp)&"/"&month(UpLoadTimeTmp)&"/"&day(UpLoadTimeTmp)&" "&hour(UpLoadTimeTmp)&":"&minute(UpLoadTimeTmp)&":"&second(UpLoadTimeTmp)
	strInsRet="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",TO_DATE('"&UpLoadTime&"','YYYY/MM/DD/HH24/MI/SS')" &_
		",'N','3','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
	conn.BeginTrans
		conn.execute strUpdRet
		conn.execute strInsRet
	if err.number = 0 then
       conn.CommitTrans
    else            
       conn.RollbackTrans
    end if   
end function

'寄存
function funcSafeKeep(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'更新該筆紀錄的 BILLSTATUS 更新為 4
	strUpdStore="Update BillBase set billstatus='4' where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	strInsStore="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'N','4','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
		conn.execute strUpdStore
		conn.execute strInsStore
end function

'公示
function funcPublic(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'更新該筆紀錄的 BILLSTATUS 更新為 5
	strUpdGov="Update BillBase set billstatus='5' where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	strInsGov="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'N','5','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
		conn.execute strUpdGov
		conn.execute strInsGov
end function

'刪除
function funcDCIdel(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,BillStatus,RecordDate,RecordMemberID,DelReason,DelNote,CaseInStatus,theBatchTime)

	'該筆紀錄的打驗資料表的 BILLSTATUS 更新為 6
	strUpdDelTemp="Update BillBaseTmp set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where BillNo='"&BillNo&"'"
	conn.execute strUpdDelTemp

	'寫入刪除原因(判斷是否已經有 有的話就update)
	strReaCheck="select * from BillDeleteReason where BillSN="&BillSN
	set rsReaCheck=conn.execute(strReaCheck)
	if rsReaCheck.eof then
		strReaDel="Insert into BillDeleteReason(BillSN,DelDate,DelReason,Note)" &_
			" values("&BillSN&",sysdate,'"&DelReason&"','"&DelNote&"')" 
	else
		strReaDel="Update BillDeleteReason set DelDate=sysdate,DelReason='"&DelReason&"'" &_
			",Note='"&DelNote&"' where BillSN="&BillSN
	end if
	'更新該筆紀錄的 BILLSTATUS 更新為 6
	strUpdDel="Update BillBase set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&BillSN
	'刪除入案舉發單時，只處理刪除人
	strDelMem="Update BillBase set DelMemberID="&Session("User_ID")&" where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	strInsDel="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',sysdate" &_
		","&Session("User_ID")&",sysdate,'E','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 

	'檢查是否已刪除過
	DelCheckFlag01="0"
	strDelCheck="select BillSN from DciLog where BillSN="&BillSN&" and ExchangeTypeID='E' and DciReturnStatusID is null"
	set rsDelCheck=conn.execute(strDelCheck)
	if not rsDelCheck.eof then
		DelCheckFlag01="1"
	end if
	rsDelCheck.close
	set rsDelCheck=nothing
if DelCheckFlag01="1" then
	funcDCIdel="N"
else
	'若該案件狀態為車籍查詢或未處理，則不需要寫入DCILOG，其餘的要
	if BillStatus="0" or BillStatus="1" or CaseInStatus="0" then
		if trim(DelReason)="" or isnull(DelReason) then
			conn.execute strUpdDel
			conn.execute strUpdDelTemp
		else
			conn.BeginTrans
				conn.execute strReaDel
				conn.execute strUpdDel
				conn.execute strUpdDelTemp
			if err.number = 0 then
			   conn.CommitTrans
			else            
			   conn.RollbackTrans
			end if   
		end if
		funcDCIdel="Y"
	else
		conn.BeginTrans
			conn.execute strReaDel
			conn.execute strUpdDel
			conn.execute strInsDel
			conn.execute strDelMem
		if err.number = 0 then
		   conn.CommitTrans
		else            
		   conn.RollbackTrans
		end if   
		funcDCIdel="Y"
	end if
end if
end function

'寄存轉公示撤銷（Ｙ）
function funcStoreAndSendToGov(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'把select 出來的紀錄寫入到DCILog
	strInsStoG="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'N','Y','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
	conn.execute strInsStoG

	strUpdStoG="update billmailhistory set MailTypeID=null where BillSN="&BillSN
	conn.execute strUpdStoG
end function

'取得BatchNumber時分秒部份
function getBatchTimeFormat(BatchTime)
	getBatchTimeFormat=Right("00"&hour(BatchTime),2)&Right("00"&minute(BatchTime),2)&Right("00"&second(BatchTime),2)
end function

'收受
function funcBillGet(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

	'更新該筆紀錄的 BILLSTATUS 更新為 7
	strUpdGet="Update BillBase set billstatus='7' where SN="&BillSN
	'把select 出來的紀錄寫入到DCILog
	strInsGet="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
		",RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber)"&_
		"values(DCILOG_SEQ.nextval,"&BillSN&",'"&BillNo&"',"&BillTypeID&",'"&CarNo&"'" &_
		",'"&BillUnitID&"',TO_DATE('"&RecordDateTemp&"','YYYY/MM/DD/HH24/MI/SS')" &_
		","&Session("User_ID")&",sysdate,'N','7','"&Session("DCIwindowName")&"','"&theBatchTime&"'" &_
		")" 
	conn.BeginTrans
		conn.execute strUpdGet
		conn.execute strInsGet
	if err.number = 0 then
       conn.CommitTrans
    else            
       conn.RollbackTrans
    end if   
end function
%>