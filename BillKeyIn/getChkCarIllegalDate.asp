<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
' 檔案名稱： getChkCarIllegalDate.asp
	'ChkCarNoIllDateHour檢查同車號、法條、違規日期的舉發單的違規時間幾小時內有重複

		'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	'檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
	CarCnt=0
	'日期時間
	illegalDateTmp=gOutDT(request("IllDate"))&" "&left(trim(request("IllTime")),2)&":"&right(trim(request("IllTime")),2)&":00"
	'雲林縣要檢查當天整天
	if sys_City="雲林縣" then
		illegalDate1=gOutDT(request("IllDate"))&" 00:00:00"
		illegalDate2=gOutDT(request("IllDate"))&" 23:59:59"
	else
		illegalDate1=DateAdd("h",-2,illegalDateTmp)
		illegalDate2=DateAdd("h",2,illegalDateTmp)
	end if
	strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"

	theCarSimpleID=trim(request("CarSimpleID"))
	'法條一
	Rule1Flag=0
	FlagRuleDetail=0
	strRule1=""
	if trim(request("IllRule1"))<>"" then
		strRule1=" and Rule1='"&trim(request("IllRule1"))&"'"
		Rule1Flag=1
		'檢查車種跟法條內容相不相符
		if left(request("IllRule1"),2)="31" And Trim(request("IllRule1"))<>"31300011" then
			strRuleDetail1="select IllegalRule from Law where ItemID='"&trim(request("IllRule1"))&"'"
			set rsRuleDetail1=conn.execute(strRuleDetail1)
			If Not rsRuleDetail1.eof Then
				if (InStr(rsRuleDetail1("IllegalRule"),"機器腳踏車")>0 Or InStr(rsRuleDetail1("IllegalRule"),"機車")>0) and (theCarSimpleID="1" or theCarSimpleID="2") then
					FlagRuleDetail=1
				elseif (InStr(rsRuleDetail1("IllegalRule"),"小客車")>0 or InStr(rsRuleDetail1("IllegalRule"),"汽車")>0) and (theCarSimpleID="3" or theCarSimpleID="4") then
					FlagRuleDetail=1
				elseif InStr(rsRuleDetail1("IllegalRule"),"安全帶")>0 and (theCarSimpleID="3" or theCarSimpleID="4") then
					FlagRuleDetail=1
				elseif InStr(rsRuleDetail1("IllegalRule"),"安全帽")>0 and (theCarSimpleID="1" or theCarSimpleID="2") then
					FlagRuleDetail=1
				end if
			end if
			rsRuleDetail1.close
			set rsRuleDetail1=nothing
		end if
	end if

	'法條二
	Rule2Flag=0
	strRule2=""
	if trim(request("IllRule2"))<>"" then
		strRule2=" and Rule2='"&trim(request("IllRule2"))&"'"
		Rule2Flag=1
		'檢查車種跟法條內容相不相符
		if left(request("IllRule2"),2)="31" And Trim(request("IllRule2"))<>"31300011" then
			strRuleDetail2="select IllegalRule from Law where ItemID='"&trim(request("IllRule2"))&"'"
			set rsRuleDetail2=conn.execute(strRuleDetail2)
			If Not rsRuleDetail2.eof Then
				if InStr(rsRuleDetail2("IllegalRule"),"機器腳踏車")>0 and (theCarSimpleID="1" or theCarSimpleID="2") then
					FlagRuleDetail=1
				elseif (InStr(rsRuleDetail2("IllegalRule"),"小客車")>0 or InStr(rsRuleDetail2("IllegalRule"),"汽車")>0) and (theCarSimpleID="3" or theCarSimpleID="4") then
					FlagRuleDetail=1
				end if
			end if
			rsRuleDetail2.close
			set rsRuleDetail2=nothing
		end if
	end if

	if Rule1Flag=1 or Rule2Flag=1 Then
		if sys_City="宜蘭縣" Or sys_City="基隆市" Or sys_City="台南市" Or sys_City="台東縣" Or sys_City="雲林縣" Then '宜蘭不用判斷同法條
			strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDate
		Else
			strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDate&strRule1&strRule2
		End If 
		set rsRep=conn.execute(strRep)
		If Not rsRep.eof Then 
			CarCnt=1
			theOldIllDate1=year(rsRep("IllegalDate"))-1911&"年"&month(rsRep("IllegalDate"))&"月"&day(rsRep("IllegalDate"))&"日"&" "&hour(rsRep("IllegalDate"))&"點"&minute(rsRep("IllegalDate"))&"分"

			ConnExecute trim(request("CarID"))&","&rsRep("IllegalDate")&" "&"，同車號同違規時間",223 
		else
			CarCnt=0
		end if
		rsRep.close
		set rsRep=nothing
	else
		CarCnt=0
	end if

	strUChk="select UnitName from UnitInfo where UnitID='"&trim(request("BillUnitID"))&"'"
	set rsUChk=conn.execute(strUChk)
	if rsUChk.eof then
		FlagRuleDetail=2
	end if
	rsUChk.close
	set rsUChk=nothing

	if sys_City="高雄市" Then
		If FlagRuleDetail<>2 then
			strVIP="select * from SpecCar where CarNo='"&trim(request("CarID"))&"' and RecordStateID<>-1"
			set rsVIP=conn.execute(strVIP)
			if not rsVIP.eof Then
				If FlagRuleDetail=1 then
					FlagRuleDetail=3
				Else
					FlagRuleDetail=4
				End If 
			end if
			rsVIP.close
			set rsVIP=Nothing
		End if
	End If
	
	if sys_City="苗栗縣" Then
		If Trim(request("BillCheck"))="1" Then
			If Left(Trim(request("IllRule1")),2)="40" Or Left(Trim(request("IllRule2")),2)="40" Or Left(Trim(request("IllRule2")),5)="33101" Or Left(Trim(request("IllRule2")),5)="43102" Then
				If Trim(request("IllSpeed"))="" Then
					chkIllegalSpeed="null"
				Else 
					chkIllegalSpeed=Trim(request("IllSpeed"))
				End If 
				If Trim(request("RuleSpeed"))="" Then
					chkRuleSpeed="null"
				Else 
					chkRuleSpeed=Trim(request("RuleSpeed"))
				End If 

				strBChk="select count(*) as cnt from billruncaraccept " &_
				" where CarNo='"&Trim(request("CarID"))&"'" &_
				" and IllegalDate=to_date('"&illegalDateTmp&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0" &_
				" and IllegalSpeed="&chkIllegalSpeed&" and RuleSpeed="&chkRuleSpeed
			Else
				strBChk="select count(*) as cnt from billruncaraccept " &_
				" where CarNo='"&Trim(request("CarID"))&"'" &_
				" and IllegalDate=to_date('"&illegalDateTmp&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0" 
			End If 
			Set rsBchk=conn.execute(strBChk)
			If Trim(rsBchk("cnt"))="0" Then
				FlagRuleDetail=7
			End If 
			rsBchk.close
			Set rsBchk=Nothing 
		ElseIf Trim(request("BillCheck"))="2" Then
			strBChk="select count(*) as cnt from billStopcaraccept " &_
				" where CarNo='"&Trim(request("CarID"))&"'" &_
				" and IllegalDate=to_date('"&illegalDateTmp&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0" &_
				" and BillUnitID='"&trim(Request("BillUnitID"))&"' and BillMemID1 in(select MemberID from memberdata where LoginID='"&trim(Request("BillMemID1"))&"' and recordstateid=0)" &_
				" and CarSimpleID="&trim(Request("CarSimpleID"))&_
				" and BillNo='"&Trim(request("BillNO"))&"'"
			' and DriverID='"&Trim(request("CreditID"))&"'
			Set rsBchk=conn.execute(strBChk)
			If Trim(rsBchk("cnt"))="0" Then
'				strCChk="select count(*) as cnt from billStopcaraccept " &_
'					" where RecordStateID=0" &_
'					" and BillNo='"&Trim(request("BillNO"))&"'"
'				Set rsCchk=conn.execute(strCChk)
'				If Trim(rsCchk("cnt"))<>"0" Then
					FlagRuleDetail=7
'				End If 
'				rsCchk.close
'				Set rsCchk=Nothing 
			End If 
			rsBchk.close
			Set rsBchk=Nothing 		
		End If 
	End If 

	if sys_City="南投縣" Then
		If FlagRuleDetail<>2 Then
			strNTChk="select * from apconfigure where id=777"
			Set rsNTChk=conn.execute(strNTChk)
			If Not rsNTChk.eof Then 
				If Trim(rsNTChk("value"))="1" then 
					illegalDate1b=DateAdd("n",-6,illegalDateTmp)
					illegalDate2b=DateAdd("n",6,illegalDateTmp)
					strIllDateb=" and IllegalDate between TO_DATE('"&year(illegalDate1b)&"/"&month(illegalDate1b)&"/"&day(illegalDate1b)&" "&Hour(illegalDate1b)&":"&minute(illegalDate1b)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2b)&"/"&month(illegalDate2b)&"/"&day(illegalDate2b)&" "&Hour(illegalDate2b)&":"&minute(illegalDate2b)&":59','YYYY/MM/DD/HH24/MI/SS')"
					strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateb&strRule1&strRule2
					set rsRep=conn.execute(strRep)
					If Not rsRep.eof Then 
						FlagRuleDetail=5
					end if
					rsRep.close
					set rsRep=Nothing

					If (Left(trim(request("IllRule1")),2)="40" Or Left(trim(request("IllRule2")),2)="40") And FlagRuleDetail<>5 Then
						strRep2="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateb&" and (Rule1 like '40%' or Rule2 like '40%')"
						Set rsRep2=conn.execute(strRep2)
						If Not rsRep2.eof Then 
							FlagRuleDetail=5
						End If 
						rsRep2.close
						Set rsRep2=Nothing 
					End if
				End If 
			End If 
			rsNTChk.close
			Set rsNTChk=Nothing 
		End If 
	elseif sys_City="宜蘭縣" Then
		If FlagRuleDetail<>2 Then
			strNTChk="select * from apconfigure where id=777"
			Set rsNTChk=conn.execute(strNTChk)
			If Not rsNTChk.eof Then 
				If Trim(rsNTChk("value"))="1" then 

					strIllDateb=" and IllegalDate between TO_DATE('"&year(illegalDateTmp)&"/"&month(illegalDateTmp)&"/"&day(illegalDateTmp)&" "&"0:0:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDateTmp)&"/"&month(illegalDateTmp)&"/"&day(illegalDateTmp)&" "&"23:59:59','YYYY/MM/DD/HH24/MI/SS')"

					strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateb

					set rsRep=conn.execute(strRep)
					If Not rsRep.eof Then 
						FlagRuleDetail=5
					end if
					rsRep.close
					set rsRep=Nothing
				End If 
			End If 
			rsNTChk.close
			Set rsNTChk=Nothing 
		End If 
	End If 

	'if sys_City="花蓮縣" Then
		If FlagRuleDetail<>2 Then
			If Trim(request("IllegalAddress"))="" Then
				IllegalAddressTmp1="HaveNoAddress.."
			Else
				IllegalAddressTmp1=Trim(request("IllegalAddress"))
			End If 
			strIllDateb=" and IllegalDate=TO_DATE('"&year(illegalDateTmp)&"/"&month(illegalDateTmp)&"/"&day(illegalDateTmp)&" "&Hour(illegalDateTmp)&":"&minute(illegalDateTmp)&":00','YYYY/MM/DD/HH24/MI/SS') " 
			if sys_City="台中市" Then
				strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateb&strRule1&strRule2
			Else
				strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and IllegalAddress='"&IllegalAddressTmp1&"' and RecordstateID=0 "&strIllDateb&strRule1&strRule2
			End If 
			set rsRep=conn.execute(strRep)
			If Not rsRep.eof Then 
				FlagRuleDetail=6
			end if
			rsRep.close
			set rsRep=Nothing

			if sys_City="台中市" Then
				If ((Left(trim(request("IllRule1")),2)="55" And Len(trim(request("IllRule1")))=7) Or (Left(trim(request("IllRule2")),2)="55" And Len(trim(request("IllRule2")))=7) Or (Left(trim(request("IllRule1")),2)="56"  And Len(trim(request("IllRule1")))=7) Or (Left(trim(request("IllRule2")),2)="56") And Len(trim(request("IllRule2")))=7) And FlagRuleDetail<>5 Then
					illegalDate1c=DateAdd("h",-2,illegalDateTmp)
					illegalDate2d=DateAdd("h",2,illegalDateTmp)
					strIllDateC=" and IllegalDate between TO_DATE('"&year(illegalDate1c)&"/"&month(illegalDate1c)&"/"&day(illegalDate1c)&" "&Hour(illegalDate1c)&":"&minute(illegalDate1c)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2d)&"/"&month(illegalDate2d)&"/"&day(illegalDate2d)&" "&Hour(illegalDate2d)&":"&minute(illegalDate2d)&":59','YYYY/MM/DD/HH24/MI/SS')"
					strRep2="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateC&" and ((Rule1 like '55%' and length(Rule1)=7) or (Rule2 like '55%' and length(Rule1)=7) or (Rule1 like '56%' and length(Rule1)=7) or (Rule2 like '56%' and length(Rule1)=7))"
					Set rsRep2=conn.execute(strRep2)
					If Not rsRep2.eof Then 
						FlagRuleDetail=6
					End If 
					rsRep2.close
					Set rsRep2=Nothing 
				End If
			Else
				If ((Left(trim(request("IllRule1")),2)="55" And Len(trim(request("IllRule1")))=7) Or (Left(trim(request("IllRule2")),2)="55" And Len(trim(request("IllRule2")))=7) Or (Left(trim(request("IllRule1")),2)="56" And Len(trim(request("IllRule1")))=7) Or (Left(trim(request("IllRule2")),2)="56") And Len(trim(request("IllRule2")))=7) And FlagRuleDetail<>5 Then
					strIllDateC=" and IllegalDate=TO_DATE('"&year(illegalDateTmp)&"/"&month(illegalDateTmp)&"/"&day(illegalDateTmp)&" "&Hour(illegalDateTmp)&":"&minute(illegalDateTmp)&":00','YYYY/MM/DD/HH24/MI/SS')"
					strRep2="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 "&strIllDateC&" and ((Rule1 like '55%' and length(Rule1)=7) or (Rule2 like '55%' and length(Rule1)=7) or (Rule1 like '56%' and length(Rule1)=7) or (Rule2 like '56%' and length(Rule1)=7))"
					Set rsRep2=conn.execute(strRep2)
					If Not rsRep2.eof Then 
						FlagRuleDetail=6
					End If 
					rsRep2.close
					Set rsRep2=Nothing 
				End If
			End If 
			if sys_City="台中市" Then
				If FlagRuleDetail=0 then
					If (trim(request("IllRule1"))="1210401" Or trim(request("IllRule1"))="1210402" Or trim(request("IllRule2"))="1210401" Or trim(request("IllRule2"))="1210402") And FlagRuleDetail<>6 Then
						strIllChk="select count(*) as cnt from billbase where CarNo='"&trim(request("CarID"))&"' and (Rule1 in ('1210401','1210402') or Rule2 in ('1210401','1210402')) and RecordStateID=0"
						Set rsIllChk=conn.execute(strIllChk)
						If Not rsIllChk.eof Then
							If CDbl(rsIllChk("cnt"))>0 Then
								FlagRuleDetail=8
							End If 
						End If
						rsIllChk.close
						Set rsIllChk=Nothing 
					End If 
				End If 
			End if
		End If 
	
	'End If 
	'FlagRuleDetail=1,違規事實與簡式車種不符
	'FlagRuleDetail=2,舉發單位代號輸入錯誤
	'FlagRuleDetail=3,舉發單位代號輸入錯誤+車號為業管車輛 (高)
	'FlagRuleDetail=4,車號為業管車輛 (高)
	'FlagRuleDetail=5,6分鐘內有同車號同法條案件 (南頭)
	'FlagRuleDetail=6,同時間、違規地點、法條(台中 同時間、法條)
	'FlagRuleDetail=7,登記簿檢核無此筆違規資料(苗栗)
	'FlagRuleDetail=8,同車牌有開過1210401或1210402(台中103/4/8)

'檢舉案件檢查如果同一天有相同違規不可建檔
strJurgeDayError=""
'If Trim(request("JurgeDay"))<>"" Then
'	illegalDate1J=gOutDT(request("IllDate"))&" 00:00:00"
'	illegalDate2K=gOutDT(request("IllDate"))&" 23:59:59"
'	strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1J)&"/"&month(illegalDate1J)&"/"&day(illegalDate1J)&" 0:0:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2K)&"/"&month(illegalDate2K)&"/"&day(illegalDate2K)&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
'	strRep="select IllegalDate from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordstateID=0 and JurgeDay is not null "&strIllDate
'
'	set rsRep=conn.execute(strRep)
'	If Not rsRep.eof Then 
'		strJurgeDayError="此車號於"&gOutDT(request("IllDate"))&"已有民眾檢舉案件"
'
'	end if
'	rsRep.close
'	set rsRep=nothing
'End If 

If strJurgeDayError<>"" Then
%>
alert("<%=strJurgeDayError%>");
<%
else
	if sys_City="台中市" Then
		AcceptErrorFlag=0
		If Trim(request("AcceptBatchNumberChk"))="1" Then
			If Trim(request("BillCheck"))="1" Then	'逕舉
				If Trim(request("BillNO"))="" Then 
					strBRC="select * from BillRunCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
						" and CarNo='"&trim(request("CarID"))&"' and RecordStateID=0"
					set rsBRC=conn.execute(strBRC)
					if not rsBRC.eof then
						If Trim(rsBRC("CarSimpleID"))<>Trim(request("CarSimpleID")) Then
							AcceptErrorFlag=2
						End If 
					Else
						AcceptErrorFlag=3
					end if
					rsBRC.close
					set rsBRC=Nothing
				Else
					strSQL1="select * from BillStopCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
					" and BillNo='"&Trim(request("BillNO"))&"' and RecordStateID=0"
					Set rs1=conn.execute(strSQL1)
					If rs1.eof Then
						AcceptErrorFlag=1
					End If
					rs1.close
					Set rs1=Nothing 
				End If 
			ElseIf Trim(request("BillCheck"))="2" Then	'攔停
				strSQL1="select * from BillStopCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
					" and BillNo='"&Trim(request("BillNO"))&"' and RecordStateID=0"
				Set rs1=conn.execute(strSQL1)
				If rs1.eof Then
					AcceptErrorFlag=1
				End If
				rs1.close
				Set rs1=Nothing 
			End if 

		End if 
		If AcceptErrorFlag=1 Then
	%>
			alert("此單號，登記簿沒有登打資料！！");
	<%	elseIf AcceptErrorFlag=2 Then
	%>
			alert("輸入的車種與登記簿車種不符！！");
	<%	elseIf AcceptErrorFlag=3 Then
	%>
			alert("此車號，登記簿沒有登打記錄！！");
	<%
		else
			response.write "setChkCarIllegalDate(""" & CarCnt & """,""" & theOldIllDate1 & """,""" & FlagRuleDetail & """);"
		End If 
	else
		response.write "setChkCarIllegalDate(""" & CarCnt & """,""" & theOldIllDate1 & """,""" & FlagRuleDetail & """);"
	End If
End if
%>
//alert("<%=FlagRuleDetail%>");

<%
conn.close
set conn=nothing
%>

