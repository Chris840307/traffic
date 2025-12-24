<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">

<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'on error resume next
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
end if
' 告發類別
' theBilltype=1  1 攔停  2 逕舉
if trim(request("Billtype"))="" then
	theBilltype=""
else
	theBilltype=trim(request("Billtype"))
end if
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
gCh_Name = session("CH_Name")
gUnit_ID = Session("Unit_ID")

checkF2Flag=0	'檢查F2是否有效
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)

end function

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
'==========================
'是否要放大鏡功能(Y/N)
if sys_City="台東縣" then
	isBig="N" 
else
	isBig="Y" 
end if

'============================================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing
'==========================================
If Trim(request("Speed"))="1" then
	BillBaseName="billbaseTmp3"
	BILLILLEGALIMAGEName="BILLILLEGALIMAGETemp3"
	
else
	BillBaseName="billbaseTmp"
	BILLILLEGALIMAGEName="BILLILLEGALIMAGETemp2"
End If 
'==========================================
'新增告發單
if trim(request("kinds"))="DB_insert" Then
	'違規地點處理
	If trim(request("IllegalAddress") &"")="" Then
		theIllegalAddress=""
	Else
		theIllegalAddress=Replace(Replace(Replace(trim(request("IllegalAddress") &"")," ",""),"'",""),"|","")
	End If 

	chkIsExistBillNumFlag=0
	if trim(request("Billno1"))<>"" then
		strchkno="select BillNo from BillBase where BillNo='"&trim(request("Billno1"))&"' and RecordStateID=0"
		set rschkno=conn.execute(strchkno)
		if not rschkno.eof then
			chkIsExistBillNumFlag=1
		end if
		rschkno.close
		set rschkno=Nothing
	end If

	If Trim(request("RuleSpeed"))<>"" And Trim(request("IllegalSpeed"))<>"" Then
		If Trim(request("RuleSpeed"))>300 Or Trim(request("IllegalSpeed"))>300 Then
			chkIsSpeedTooOver=1
		Else
			chkIsSpeedTooOver=0
		End If 
	Else
		chkIsSpeedTooOver=0
	End If 

	
	chkIillegalDataDouble=0
	If Trim(request("Speed"))="1" Then
		strIllDate=" and StartIllegalDate=TO_DATE('"&gOutDT(request("StartIllegalDate"))&" "&left(trim(request("StartIllegalTime")),2)&":"&mid(trim(request("StartIllegalTime")),3,2)&":"&right(trim(request("StartIllegalTime")),2) &"','YYYY/MM/DD/HH24/MI/SS')"
		strChk="select BillNo,Rule1 " &_
			" from Billbase where CarNo='"&UCase(trim(request("CarNo")))&"' "&strIllDate&" and Rule1='"&trim(request("Rule1"))&"' and RecordStateID=0" 
		Set rsChk=conn.execute(strChk)
		If Not rsChk.eof Then
			chkIillegalDataDouble=1
			
		End If 
		rsChk.close
		Set rsChk=Nothing 
		'response.write strChk
	End If

	chkIllegalDateAndCar_KS=0
	chkAlertString=""
	If sys_City="高雄市" Then
		illegalDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
		illegalDate2=gOutDT(request("IllegalDate"))&" 23:59:59"
		strIllDate=" and IllegalDate between TO_DATE('"&illegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&illegalDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,BillNo,Rule1,Rule2 " &_
			" from Billbase where CarNo='"&UCase(trim(request("CarNo")))&"' and RecordStateID=0 " &_
			" " & strIllDate
		Set rsChk=conn.execute(strChk)
		If Not rsChk.eof Then
			chkIllegalDateAndCar_KS=1
			chkAlertString="此車號在此違規日有違規紀錄，舉發單位:"&Trim(rsChk("UnitName"))&"，單號:"&Trim(rsChk("BillNo"))&"，法條:"&Trim(rsChk("Rule1"))
			If Trim(rsChk("Rule2"))<>"" Then
				chkAlertString=chkAlertString & "/" & Trim(rsChk("Rule2"))
			End If 
		End If 
		rsChk.close
		Set rsChk=Nothing 
	End If

	theIllegalDate="null"
	If trim(Request("IllegalDate"))<>"" Then
		theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
		If Trim(request("Speed"))="1" Then
			theCheckIllegalDate="to_date('" & gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2) & "','YYYY/MM/DD/HH24/MI/SS')"
		Else
			theCheckIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
		End If 
	End If 

	chkReKeyInImgBill=0
	If sys_City="彰化縣" Or Trim(request("Speed"))="1" Then ' Or sys_City="台中市"
		If (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) then
			strChk="select count(*) as cnt from Billbase where CarNo='"&Trim(request("CarNo"))&"'" & _
			" and IllegalDate="& theCheckIllegalDate & _
			" and ((Rule1 like '55%' and length(Rule1)=7) or (Rule2 like '55%' and length(Rule2)=7) or (Rule1 like '56%' and length(Rule1)=7) or (Rule2 like '56%' and length(Rule2)=7)) and Recordstateid=0"

			'response.write strChk
			Set rsChk=conn.execute(strChk)
			If not rsChk.eof Then	
				If CInt(rsChk("cnt"))>0 then
					chkReKeyInImgBill=1
				End If 
			End If 
			rsChk.close
			Set rsChk=Nothing 
		Else
			strChk="select count(*) as cnt from Billbase where CarNo='"&Trim(request("CarNo"))&"'" & _
			" and IllegalDate="& theCheckIllegalDate & _
			" and Rule1='"&Trim(request("Rule1"))&"' and Recordstateid=0"


			'response.write strChk
			Set rsChk=conn.execute(strChk)
			If not rsChk.eof Then	
				If CInt(rsChk("cnt"))>0 then
					chkReKeyInImgBill=1
				End If 
			End If 
			rsChk.close
			Set rsChk=Nothing 
		End If 

		
	End If 
	
	'科技執法
	chkIsSpeedRuleFlag_TC=0
	chkIsIllegalTimeNoRuleFlag_TC=0
	chkIsRule5620002Flag_TC=0
	chkIsDoubleFlag_TC=0
	if Trim(request("Speed"))="1" Then
		if sys_City="台中市" then
			'當天同違規地點超速
			If left(trim(request("Rule1")),2)="40" Or left(trim(request("Rule2")),2)="40" Or left(trim(request("Rule1")),5)="33101" Or left(trim(request("Rule2")),5)="33101" Or left(trim(request("Rule1")),5)="43102" Or left(trim(request("Rule2")),5)="43102" Then
				illegalDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
				illegalDate2=gOutDT(request("IllegalDate"))&" 23:59:59"
				strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
				strChk="select count(*) as cnt " &_
					" from Billbase where (Rule1 like '40%' or Rule2 like '40%' or Rule1 like '33101%' or Rule2 like '33101%' or Rule1 like '43102%' or Rule2 like '43102%') " &_
					" and carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Recordstateid=0 " & strIllDate & " and IllegalAddress='" & theIllegalAddress & "'"
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
					If CDbl(rsChk("cnt"))>0 Then
						chkIsSpeedRuleFlag_TC=1
					End If 
				End If 
				rsChk.close
				Set rsChk=Nothing
			End If 

			If (left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) Then
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)
				illegalDate1=DateAdd("h",-2,illegalDateTmp)
				illegalDate2=DateAdd("h",2,illegalDateTmp)
				strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
				strChk="select count(*) as cnt " &_
					" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and ((Rule1 like '55%' and length(Rule1)=7) or (Rule1 like '56%' and length(Rule1)=7))" &_
					" and Recordstateid=0 " & strIllDate 
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
					If CDbl(rsChk("cnt"))>0 Then
						chkIsDoubleFlag_TC=1
					End If 
				End If 
				rsChk.close
				Set rsChk=Nothing
			ElseIf trim(request("Rule1"))="5310001" or trim(request("Rule1"))="5320001" or trim(request("Rule1"))="6020302" Then
				illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)
				illegalDate1=DateAdd("h",-2,illegalDateTmp)
				illegalDate2=DateAdd("h",2,illegalDateTmp)
				strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
				strChk="select count(*) as cnt " &_
					" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and (Rule1='5310001' or Rule1='5320001' or Rule1='6020302')" &_
					" and Recordstateid=0 " & strIllDate 
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
					If CDbl(rsChk("cnt"))>0 Then
						chkIsDoubleFlag_TC=1
					End If 
				End If 
				rsChk.close
				Set rsChk=Nothing
			Else
				illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)
				illegalDate1=DateAdd("h",-2,illegalDateTmp)
				illegalDate2=DateAdd("h",2,illegalDateTmp)
				strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
				strChk="select count(*) as cnt " &_
					" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Rule1='"&trim(request("Rule1"))&"'" &_
					" and Recordstateid=0 " & strIllDate 
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
					If CDbl(rsChk("cnt"))>0 Then
						chkIsDoubleFlag_TC=1
					End If 
				End If 
				rsChk.close
				Set rsChk=Nothing
			End If 
				
			'當天5620002
			If trim(request("Rule1"))="5620002" Or trim(request("Rule2"))="5620002" Or trim(request("Rule3"))="5620002" Then
				illegalDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
				illegalDate2=gOutDT(request("IllegalDate"))&" 23:59:59"
				strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
				strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate" &_
					" from Billbase where (Rule1='5620002' or Rule2='5620002' or Rule3='5620002') " &_
					" and carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Recordstateid=0 " & strIllDate
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
					chkIsRule5620002Flag_TC=1
					chkIsRule5620002Unit=Trim(rsChk("UnitName"))
					chkIsRule5620002IllegalTime=Trim(rsChk("IllegalDate"))
				End If 
				rsChk.close
				Set rsChk=Nothing
			End If 
			
			'同違規時間、車號
			illegalDateTmpTC=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)

			strIllDateTC=" and IllegalDate=TO_DATE('"&year(illegalDateTmpTC)&"/"&month(illegalDateTmpTC)&"/"&day(illegalDateTmpTC)&" "&Hour(illegalDateTmpTC)&":"&minute(illegalDateTmpTC)&":"&Second(illegalDateTmpTC)&"','YYYY/MM/DD/HH24/MI/SS')"
			strChkTC="select count(*) as cnt " &_
				" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
				" and Recordstateid=0 " & strIllDateTC 
			'response.write strChk
			Set rsChk=conn.execute(strChkTC)
			If Not rsChk.eof Then	
				If CDbl(rsChk("cnt"))>0 Then
					chkIsIllegalTimeNoRuleFlag_TC=1
				End If 
			End If 
			rsChk.close
			Set rsChk=Nothing
		end if 
	end if 
	
	chkReKeyInBill=0
	If sys_City<>"台中市" Or Trim(request("Speed"))="1" Then
		strChk="select count(*) as cnt from Billbase where CarNo='"&Trim(request("CarNo"))&"'" & _
		" and IllegalAddress='"&theIllegalAddress&"'" & _
		" and IllegalDate="& theCheckIllegalDate & _
		" and Rule1='"&Trim(request("Rule1"))&"' and Recordstateid=0 "

		'response.write strChk
		Set rsChk=conn.execute(strChk)
		If not rsChk.eof Then	
			If CInt(rsChk("cnt"))>0 then
				chkReKeyInBill=1
			End If 
		End If 
		rsChk.close
		Set rsChk=Nothing 
	End If 

	'response.write strChk & "<br>"
	'response.write chkReKeyInBill & "<br>"
	'response.end
If chkIsSpeedTooOver=0 And chkIsExistBillNumFlag=0 And chkReKeyInImgBill=0 And chkIillegalDataDouble=0 And chkReKeyInBill=0 and chkIsSpeedRuleFlag_TC=0 and chkIsIllegalTimeNoRuleFlag_TC=0 and chkIsRule5620002Flag_TC=0 and chkIsDoubleFlag_TC=0 then
	 
	strSqlA="select * from "&BillBaseName&" where Sn=" & Trim(request("CheckSn"))
	set rsA=conn.execute(strSqlA)
	If Not rsA.eof then
	ReportSn=trim(rsA("Sn"))
	
	If Trim(request("Speed"))="1" Then
		'SN抓最大值
		sSQL = "select BillBase_seq.nextval as SN from Dual"
		set oRST = Conn.execute(sSQL)
		if not oRST.EOF then
			sMaxSN = oRST("SN")
		end if
		oRST.close
		set oRST = Nothing
	Else
		sMaxSN=trim(rsA("Sn"))
	End If

	If sys_City="基隆市" And Trim(request("Speed"))="1" Then 
		If trim(request("Rule4"))<>"" Then
			Session("SpeedKeyIn_Rule4")=trim(request("Rule4"))
		Else
			Session("SpeedKeyIn_Rule4")=""
		End If 
	elseIf (sys_City="雲林縣" Or sys_City="台東縣" Or sys_City="屏東縣" Or sys_City="南投縣") And Trim(request("Speed"))="1" Then 
		If trim(request("BillMem1"))<>"" Then
			Session("SpeedKeyIn_BillMem1")=trim(request("BillMem1"))
		Else
			Session("SpeedKeyIn_BillMem1")=""
		End If 
	Else
		Session("SpeedKeyIn_Rule4")=""
		Session("SpeedKeyIn_BillMem1")=""
	End If 
	
	'BillBase
	'If sys_City="高雄市" Then
		ColAdd=",IllegalZip"
		valueAdd=",'"&trim(rsA("IllegalZip"))&"'"
	'End if	
		If Trim(request("Speed"))="1" Then
			theCarSimpleID="null"
			If trim(Request("CarSimpleID"))<>"" Then
				theCarSimpleID=trim(Request("CarSimpleID"))
			End If 
			theCarAddID="null"
			If trim(Request("CarAddID"))<>"" Then
				theCarAddID=trim(Request("CarAddID"))
			End If 
			
			theIllegalSpeed="null"
			If Trim(Request("IllegalSpeed"))<>"" Then
				theIllegalSpeed=trim(Request("IllegalSpeed"))
			End If 
			theRuleSpeed="null"
			If Trim(Request("RuleSpeed"))<>"" Then
				theRuleSpeed=trim(Request("RuleSpeed"))
			End If 
			theForFeit1="null"
			If Trim(Request("ForFeit1"))<>"" Then
				theForFeit1=trim(Request("ForFeit1"))
			End If 
			theForFeit2="null"
			If Trim(Request("ForFeit2"))<>"" Then
				theForFeit2=trim(Request("ForFeit2"))
			End If 
			theForFeit3="null"
			If Trim(Request("ForFeit3"))<>"" Then
				theForFeit3=trim(Request("ForFeit3"))
			End If 
			theForFeit4="null"
			If Trim(Request("ForFeit4"))<>"" Then
				theForFeit4=trim(Request("ForFeit4"))
			End If 
			theInsurance="null"
			If Trim(Request("Insurance"))<>"" Then
				theInsurance=trim(Request("Insurance"))
			End If
			theUseTool="null"
			If Trim(Request("UseTool"))<>"" Then
				theUseTool=trim(Request("UseTool"))
			End If
			theBillFillDate="null"
			If trim(Request("BillFillDate"))<>"" Then
				theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
			End If 
			theDealLineDate="null"
			If trim(Request("DealLineDate"))<>"" Then
				theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
			End If 

			theJurgeDay="null"
			If trim(Request("JurgeDay"))<>"" Then
				theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
			End If 
			
			theStartIllegalDate="null"
			If trim(Request("StartIllegalDate"))<>"" Then
				theStartIllegalDate="to_date('" & gOutDT(request("StartIllegalDate") ) &" "&left(trim(request("StartIllegalTime")),2)&":"&mid(trim(request("StartIllegalTime")),3,2)&":"&right(trim(request("StartIllegalTime")),2) & "','YYYY/MM/DD/HH24/MI/SS')"
			End If 

			theEndIllegalDate="null"
			If trim(Request("IllegalDate"))<>"" Then
				theEndIllegalDate="to_date('" & gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2) & "','YYYY/MM/DD/HH24/MI/SS')"
			End If 
			
			theDistance="null"
			If trim(Request("Distance"))<>"" Then
				theDistance=trim(Request("Distance"))
			End If 

			strUpd="Update BillBaseTmp3 set " &_
			"CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&theCarSimpleID &_
			",CarAddID="&theCarAddID&",IllegalDate="&theEndIllegalDate &_
			",IllegalAddressID='"&trim(request("IllegalAddressID"))&"'" &_
			",IllegalAddress='"&theIllegalAddress&"'" &_
			",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed &_
			",RuleSpeed="&theRuleSpeed&",Rule2='"&trim(request("Rule2"))&"'" &_
			",ForFeit1="&theForFeit1&",ForFeit2="&theForFeit2 &_
			",Rule4='"&trim(request("Rule4"))&"'" &_
			",Insurance="&theInsurance&",UseTool="&theUseTool &_
			",ProjectID='"&trim(request("ProjectID"))&"'" &_
			",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
			",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
			",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
			",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
			",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
			",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
			",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
			",Note='"&trim(request("Note"))&"',EquipmentID='1'" &_
			",StartIllegalDate="&theStartIllegalDate&",Distance="&theDistance &_
			",CheckDate=sysdate" &_
			" where Sn=" & Trim(request("CheckSn"))
'response.write strUpd
			'response.end
			conn.execute strUpd  

			strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID4,BillMem4,BillMemID2,BillMem2,BillMemID3,BillMem3" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,ImageFileName"&ColAdd&",JurgeDay" &_
				",StartIllegalDate,Distance" &_
				")" &_
				" values("&sMaxSN&",'"&trim(rsA("BillTypeId"))&"','"&UCase(trim(rsA("Billno")))&"'" &_
				",'"&UCase(trim(request("CarNo")))&"',"&theCarSimpleID &_						          
				","&theCarAddID&","&theEndIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
				",'"&theIllegalAddress&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
				","&theRuleSpeed&","&theForFeit1&",'"&trim(request("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
				","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
				",'',null,''" &_
				",'','','"&trim(request("MemberStation"))&"'" &_
				",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
				",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
				",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
				",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'0',0,sysdate,'" & trim(Session("User_ID")) &"'" &_
				",'"&trim(request("Note"))&"','1','"&trim(rsA("RuleVer"))&"'" &_
				",'"&trim(rsA("DriverSex"))&"','"&trim(rsA("ImageFileName"))&"'" &_
				""&valueAdd&"," & theJurgeDay &_
				","& theStartIllegalDate &"," & theDistance &_
				")"
				'response.write strInsert
				'response.end
				conn.execute strInsert  

			'ConnExecute "區間速率審核通過:"&strInsert,371
		elseIf sys_City="台中市" Then
			
			theCarSimpleID="null"
			If trim(Request("CarSimpleID"))<>"" Then
				theCarSimpleID=trim(Request("CarSimpleID"))
			End If 
			theCarAddID="null"
			If trim(Request("CarAddID"))<>"" Then
				theCarAddID=trim(Request("CarAddID"))
			End If 
			
			theIllegalSpeed="null"
			If Trim(Request("IllegalSpeed"))<>"" Then
				theIllegalSpeed=trim(Request("IllegalSpeed"))
			End If 
			theRuleSpeed="null"
			If Trim(Request("RuleSpeed"))<>"" Then
				theRuleSpeed=trim(Request("RuleSpeed"))
			End If 
			theForFeit1="null"
			If Trim(Request("ForFeit1"))<>"" Then
				theForFeit1=trim(Request("ForFeit1"))
			End If 
			theForFeit2="null"
			If Trim(Request("ForFeit2"))<>"" Then
				theForFeit2=trim(Request("ForFeit2"))
			End If 
			theForFeit3="null"
			If Trim(Request("ForFeit3"))<>"" Then
				theForFeit3=trim(Request("ForFeit3"))
			End If 
			theForFeit4="null"
			If Trim(Request("ForFeit4"))<>"" Then
				theForFeit4=trim(Request("ForFeit4"))
			End If 
			theInsurance="null"
			If Trim(Request("Insurance"))<>"" Then
				theInsurance=trim(Request("Insurance"))
			End If
			theUseTool="null"
			If Trim(Request("UseTool"))<>"" Then
				theUseTool=trim(Request("UseTool"))
			End If
			theBillFillDate="null"
			If trim(Request("BillFillDate"))<>"" Then
				theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
			End If 
			theDealLineDate="null"
			If trim(Request("DealLineDate"))<>"" Then
				theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
			End If 
			theRecordDate="null"
			If trim(rsA("RecordDate"))<>"" Then
				theRecordDate="to_date('"&Year(rsA("RecordDate"))&"/"&month(rsA("RecordDate"))&"/"&day(rsA("RecordDate"))&" "&Hour(rsA("RecordDate"))&":"&Minute(rsA("RecordDate"))&":"&Second(rsA("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
			End If 
			theJurgeDay="null"
			If trim(Request("JurgeDay"))<>"" Then
				theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
			End If 

			ColAddTC=",IllegalZip='"&trim(request("IllegalZip"))&"'"
			
			strUpd="Update BillBaseTmp set " &_
			"CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&theCarSimpleID &_
			",CarAddID="&theCarAddID&",IllegalDate="&theIllegalDate&_
			",IllegalAddressID='"&trim(request("IllegalAddressID"))&"'" &_
			",IllegalAddress='"&theIllegalAddress&"'" &_
			",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed &_
			",RuleSpeed="&theRuleSpeed&",Rule2='"&trim(request("Rule2"))&"'" &_
			",ForFeit1="&theForFeit1&",ForFeit2="&theForFeit2 &_
			",Rule4='"&trim(request("Rule4"))&"'" &_
			",Insurance="&theInsurance&",UseTool="&theUseTool &_
			",ProjectID='"&trim(request("ProjectID"))&"'" &_
			",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
			",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
			",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
			",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
			",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
			",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
			",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
			",Note='"&trim(request("Note"))&"',EquipmentID='1'" &_
			""&ColAddTC &_
			",JurgeDay="&theJurgeDay &",ReportCreditID='"&Trim(request("ReportCreditID"))&"'" &_
			",ReportCaseNo='"&Trim(request("ReportCaseNo"))&"'" &_
			",CheckDate=sysdate" &_
			" where Sn=" & Trim(request("CheckSn"))
'response.write strUpd
			'response.end
			conn.execute strUpd  
			
			strDelR="delete from BILLREPORTNOTemp where BillSN="&Trim(request("CheckSn"))
			conn.execute strDelR
			If Trim(request("ReportNo"))<>"" Then
				strReportNo="insert into BillReportNoTemp(BillSN,ReportNo)" &_
					" values("&Trim(request("CheckSn"))&",'"&trim(request("ReportNo"))&"')"
				conn.execute strReportNo
			End If 
				
'			response.write "-----------------------"
			ConnExecute "影像建檔分局審核通過:"&strInsert,371
		Else
			theCarSimpleID="null"
			If trim(Request("CarSimpleID"))<>"" Then
				theCarSimpleID=trim(Request("CarSimpleID"))
			End If 
			theCarAddID="null"
			If trim(Request("CarAddID"))<>"" Then
				theCarAddID=trim(Request("CarAddID"))
			End If 

			theIllegalSpeed="null"
			If Trim(Request("IllegalSpeed"))<>"" Then
				theIllegalSpeed=trim(Request("IllegalSpeed"))
			End If 
			theRuleSpeed="null"
			If Trim(Request("RuleSpeed"))<>"" Then
				theRuleSpeed=trim(Request("RuleSpeed"))
			End If 

			theForFeit1="null"
			If Trim(Request("ForFeit1"))<>"" Then
				theForFeit1=trim(Request("ForFeit1"))
			End If 
			theForFeit2="null"
			If Trim(Request("ForFeit2"))<>"" Then
				theForFeit2=trim(Request("ForFeit2"))
			End If 
			theForFeit3="null"
			If Trim(Request("ForFeit3"))<>"" Then
				theForFeit3=trim(Request("ForFeit3"))
			End If 
			theForFeit4="null"
			If Trim(Request("ForFeit4"))<>"" Then
				theForFeit4=trim(Request("ForFeit4"))
			End If 
			theInsurance="null"
			If Trim(Request("Insurance"))<>"" Then
				theInsurance=trim(Request("Insurance"))
			End If
			theUseTool="null"
			If Trim(Request("UseTool"))<>"" Then
				theUseTool=trim(Request("UseTool"))
			End If

			theBillFillDate="null"
			If trim(Request("BillFillDate"))<>"" Then
				theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
			End If 
			theDealLineDate="null"
			If trim(Request("DealLineDate"))<>"" Then
				theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
			End If 
			theRecordDate="null"
			If trim(rsA("RecordDate"))<>"" Then
				theRecordDate="to_date('"&Year(rsA("RecordDate"))&"/"&month(rsA("RecordDate"))&"/"&day(rsA("RecordDate"))&" "&Hour(rsA("RecordDate"))&":"&Minute(rsA("RecordDate"))&":"&Second(rsA("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
			End If 
			theJurgeDay="null"
			If trim(Request("JurgeDay"))<>"" Then
				theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
			End If 

			strUpd="Update BillBaseTmp set " &_
			"CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&theCarSimpleID &_
			",CarAddID="&theCarAddID&",IllegalDate="&theIllegalDate&_
			",IllegalAddressID='"&trim(request("IllegalAddressID"))&"'" &_
			",IllegalAddress='"&theIllegalAddress&"'" &_
			",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed &_
			",RuleSpeed="&theRuleSpeed&",Rule2='"&trim(request("Rule2"))&"'" &_
			",ForFeit1="&theForFeit1&",ForFeit2="&theForFeit2 &_
			",Rule4='"&trim(request("Rule4"))&"'" &_
			",Insurance="&theInsurance&",UseTool="&theUseTool &_
			",ProjectID='"&trim(request("ProjectID"))&"'" &_
			",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
			",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
			",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
			",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
			",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
			",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
			",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
			",Note='"&trim(request("Note"))&"',EquipmentID='1'" &_
			",JurgeDay="&theJurgeDay &",ReportCreditID='"&Trim(request("ReportCreditID"))&"'" &_
			",ReportCaseNo='"&Trim(request("ReportCaseNo"))&"'" &_
			",CheckDate=sysdate" &_
			" where Sn=" & Trim(request("CheckSn"))
'response.write strUpd
			'response.end
			conn.execute strUpd  
			
			RecordMemberID_Temp=""
			If sys_City="彰化縣" Then
				RecordMemberID_Temp=theRecordMemberID
			Else
				RecordMemberID_Temp=trim(rsA("RecordMemberID"))
			End if

			strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID4,BillMem4,BillMemID2,BillMem2,BillMemID3,BillMem3" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,ImageFileName"&ColAdd&",JurgeDay" &_
				")" &_
				" values("&sMaxSN&",'"&trim(rsA("BillTypeId"))&"','"&UCase(trim(rsA("Billno")))&"'" &_
				",'"&UCase(trim(request("CarNo")))&"',"&theCarSimpleID &_						          
				","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
				",'"&theIllegalAddress&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
				","&theRuleSpeed&","&theForFeit1&",'"&trim(request("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
				","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
				",'',null,''" &_
				",'','','"&trim(request("MemberStation"))&"'" &_
				",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
				",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
				",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
				",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'0',0,"&theRecordDate&"," & RecordMemberID_Temp &"" &_
				",'"&trim(request("Note"))&"','1','"&trim(rsA("RuleVer"))&"'" &_
				",'"&trim(rsA("DriverSex"))&"','"&trim(rsA("ImageFileName"))&"'" &_
				""&valueAdd&"," & theJurgeDay &_
				")"
				'response.write strInsert
				'response.end
				conn.execute strInsert  

			'ConnExecute "民眾檢舉審核通過:"&strInsert,371
		End if
	End If
	rsA.close
	Set rsA=Nothing 
	'寫入BILLILLEGALIMAGE
	strSqlB="select * from "&BILLILLEGALIMAGEName&" where BillSn=" & Trim(request("CheckSn"))
	set rsB=conn.execute(strSqlB)
	If Not rsB.eof Then
		'只將有效照片寫到舉發資料
		fileTemp1=""
		fileTemp2=""
		fileTemp3=""
		fileTemp4=""
		If Trim(request("chkImgNoUseA"))="1" Then
			If trim(request("ImageFileNameA"))<>"" Then
				fileTemp1=trim(request("ImageFileNameA"))
			End If 
		End If 
		If Trim(request("chkImgNoUseB"))="1" Then
			If trim(request("ImageFileNameB"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameB"))
				Else
					fileTemp2=trim(request("ImageFileNameB"))
				End If 				
			End If 
		End If 
		If Trim(request("chkImgNoUseC"))="1" Then
			If trim(request("ImageFileNameC"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameC"))
				ElseIf fileTemp2="" Then
					fileTemp2=trim(request("ImageFileNameC"))
				Else 
					fileTemp3=trim(request("ImageFileNameC"))
				End If 				
			End If 
		End If
		If Trim(request("chkImgNoUseD"))="1" Then
			If trim(request("ImageFileNameD"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameD"))
				ElseIf fileTemp2="" Then
					fileTemp2=trim(request("ImageFileNameD"))
				ElseIf fileTemp3="" Then
					fileTemp3=trim(request("ImageFileNameD"))
				Else 
					fileTemp4=trim(request("ImageFileNameD"))
				End If 				
			End If 
		End If
		If sys_City<>"台中市" Or Trim(request("Speed"))="1" Then
			strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
			"ImageFileNameD,IISImagePath) " &_
			"values("&sMaxSN&",'"&UCase(trim(rsB("Billno")))&"','"&fileTemp1&"'" &_
			",'"&fileTemp2&"','"&fileTemp3&"'" &_
			",'"&fileTemp4&"','"&trim(rsB("IISImagePath"))&"')"

			conn.execute strBillImage  
		End If 
		'將審核無效照片設為-1
'		strfileFlag=""
'		FileNameArray=Array(trim(rsB("ImageFileNameA")),trim(rsB("ImageFileNameB")),trim(rsB("ImageFileNameC")),trim(rsB("ImageFileNameD")))
'		ColArray=Array("ImageFlagA","ImageFlagB","ImageFlagC","ImageFlagD")
'		If Trim(request("chkImgNoUseA"))="-1" Then
'			For i=0 To UBound(FileNameArray)
'				If Trim(FileNameArray(i))=Trim(request("ImageFileNameA")) Then
'					strfileFlag=Trim(ColArray(i))&"='-1'"
'					Exit for
'				End If 
'			Next
'		End If 
'		If Trim(request("chkImgNoUseB"))="-1" Then
'			For i=0 To UBound(FileNameArray)
'				If Trim(FileNameArray(i))=Trim(request("ImageFileNameB")) Then
'					If strfileFlag="" Then
'						strfileFlag=Trim(ColArray(i))&"='-1'"
'					Else
'						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
'					End If 
'					
'					Exit for
'				End If 
'			Next
'		End If 
'		If Trim(request("chkImgNoUseC"))="-1" Then
'			For i=0 To UBound(FileNameArray)
'				If Trim(FileNameArray(i))=Trim(request("ImageFileNameC")) Then
'					If strfileFlag="" Then
'						strfileFlag=Trim(ColArray(i))&"='-1'"
'					Else
'						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
'					End If 
'					Exit for
'				End If 
'			Next
'		End If 
'		If Trim(request("chkImgNoUseD"))="-1" Then
'			For i=0 To UBound(FileNameArray)
'				If Trim(FileNameArray(i))=Trim(request("ImageFileNameD")) Then
'					If strfileFlag="" Then
'						strfileFlag=Trim(ColArray(i))&"='-1'"
'					Else
'						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
'					End If 
'					Exit for
'				End If 
'			Next
'		End If 
	End If
	rsB.close
	Set rsB=Nothing 
'	If sys_City<>"台中市" Then
'		If strfileFlag<>"" Then
'			strImgUpd="Update BILLILLEGALIMAGETmp set "&strfileFlag&" where BillSn=" & Trim(request("CheckSn"))
'			conn.execute strImgUpd
'		End If 
'	End If 
	'將舉發BILL SN寫回檢舉資料billbaseTmp
	If sys_City<>"台中市" Or Trim(request("Speed"))="1" Then
		strUpd1="Update "&BillBaseName&" set BillStatus='5',CheckFlag='1',BillSn="&sMaxSN  &_
			" where Sn=" & ReportSn
		conn.execute strUpd1
	Else
		strUpd1="Update "&BillBaseName&" set BillStatus='5',CheckFlag='1'"  &_
			" where Sn=" & ReportSn
		conn.execute strUpd1
	End If 
	'將BillBaseTemp2改為已審核
'	strUpd2="Update BillBaseTmp set CheckFlag='1'"  &_
'		" where Sn=" & Trim(request("CheckSn"))
'	conn.execute strUpd2
	'寫入審核紀錄
	strIns2="Insert into ReportCaseCheckRecord(Sn,ReportSN,BillTempSN,CheckFlag,RecordMemberID" &_
		",RecordDate,Note)" &_
		" values((select nvl(max(Sn),0)+1 from ReportCaseCheckRecord),"&ReportSn &_
		",'"&Trim(request("CheckSn"))&"','1',"&Trim(session("User_ID"))&"" &_
		",sysdate,''" &_
		")"
	conn.execute strIns2
%>

<script language="JavaScript">
<%
	if sys_City="台中市" then
		if trim(request("IllegalSpeed"))<>"" and trim(request("RuleSpeed"))<>"" then
			if cdbl(request("IllegalSpeed"))-cdbl(request("RuleSpeed"))>40 then
				response.write "alert('超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)');"
			end if 
		end if 
	end if 
%>

<%if trim(request("DownSn_Temp"))<>"" then%>
	location.href="BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=trim(request("DownSn_Temp"))%>&Speed=<%=trim(request("Speed"))%>";
<%end if %>
	//alert("儲存完成!");
	//opener.myForm.submit();
	//window.close();
</script>
<%
Elseif chkIsSpeedTooOver=1 then
	%>
	<script language="JavaScript">
		alert("限速或實速超過300Km，請確認是否正確！！");
	</script>
	<%
Elseif chkIsExistBillNumFlag=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
	</script>
	<%
ElseIf chkReKeyInImgBill=1 Then 
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間已有相同舉發紀錄 ,請先確認是否重複舉發！！");
	</script>
<%
elseif chkIsSpeedRuleFlag_TC=1 then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%	
ElseIf chkIsRule5620002Flag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
	</script>
<%
ElseIf chkIsDoubleFlag_TC=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間兩小時內已有舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規時間兩小時內已有舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規時間兩小時內已有舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%	
ElseIf chkIsIllegalTimeNoRuleFlag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請去舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請去舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請去舉發單資料維護系統確認！！");
	</script>
<%
ElseIf chkIillegalDataDouble=1 Then 
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間已有相同舉發紀錄 ,請先確認是否重複舉發！！");
	</script>
<%
ElseIf chkReKeyInBill=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間、違規地點已有相同舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%
End If
	If chkIllegalDateAndCar_KS=1 Then
%>
	<script language="JavaScript">
		alert("<%=chkAlertString%>");
	</script>
<%
	End If 
end if
'無效
if trim(request("kinds"))="VerifyResultNull" then
	strUpd="Update billbaseTmp set BillStatus='6'" &_
		" where Sn=" & Trim(request("CheckSn"))
	conn.execute strUpd

	ConnExecute "民眾檢舉無效案件:"&strUpd,372
%>
<script language="JavaScript">
	
	alert("儲存完成!");
	opener.myForm.submit();
	window.close();
</script>
<%
end if

'無效
if trim(request("kinds"))="DeleteCase" then
	strUpd="Update billbaseTmp set recordstateid=-1" &_
		" where Sn=" & Trim(request("CheckSn"))
	conn.execute strUpd

	ConnExecute "刪除案件:"&strUpd,372
%>
<script language="JavaScript">
	alert("刪除完成!");
<%if trim(request("DownSn_Temp"))<>"" then%>
	location.href="BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=trim(request("DownSn_Temp"))%>";
<%end if %>
</script>
<%
end if
'response.write Session("ReportCaseCheckSn")
FirstSn=""
UpSn=""
DownSn=""
LastSn=""
AllSn=0
If Trim(Session("ReportCaseCheckSn"))<>"" Then
	ThisSn=-1
	ArrayReportCaseCheckSn=Split(Trim(Session("ReportCaseCheckSn")),",")
	For i=0 To UBound(ArrayReportCaseCheckSn)
		If Trim(ArrayReportCaseCheckSn(i))=Trim(request("CheckSn")) Then
			ThisSn=i
			Exit for
		End If 
	Next 
	FirstSn=Trim(ArrayReportCaseCheckSn(0))
	If ThisSn>0 Then
		UpSn=Trim(ArrayReportCaseCheckSn(ThisSn-1))
	End If 
	If ThisSn<UBound(ArrayReportCaseCheckSn) Then
		DownSn=Trim(ArrayReportCaseCheckSn(ThisSn+1))
	End If 
	LastSn=Trim(ArrayReportCaseCheckSn(UBound(ArrayReportCaseCheckSn)))
	AllSn=UBound(ArrayReportCaseCheckSn)+1
End If 
'response.write "<br>" & UpSn & "/" & DownSn & " " & FirstSn & "/" &LastSn
PicturePath="/ReportCaseImage"

	strSql1="select * from "&BillBaseName&" where Sn=" & Trim(request("CheckSn"))

'response.write strSql1
set rs1=conn.execute(strSql1)

chkIllegaldate30day=0
chkWarningNumber=0
If sys_City="台中市" And Trim(request("Speed"))<>"1" Then
	if datediff("d",rs1("illegaldate"),rs1("recorddate"))>=30 and isnull(rs1("JurgeDay")) then
		chkIllegaldate30day=1
	end if 
	
	if (trim(rs1("JurgeDay"))="" or isnull(rs1("JurgeDay"))) and (left(trim(rs1("rule1")),2)="56") then
		strchkWNo="select count(*) as cnt from warninggetbilldetail where billno=(select ReportNo from BillReportNoTemp where BillSN="&trim(rs1("Sn"))&")"
		set rsChkWno=conn.execute(strchkWNo)

		'response.write strchkWNo
		if not rsChkWno.eof then
			if cdbl(rsChkWno("cnt"))=0 then
				chkWarningNumber=1
			end if 
		end if 
		rsChkWno.close
		set rsChkWno=nothing 
	end if 
end if 
%>
<title>數位固定桿違規影像審核</title>
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {font-size: 12px}
.style3 {
font-size: 12px ;
color: #FF0000}
.style4 {
font-size: 12px ;
}
.style5 {
color: #0000FF;
font-size: 13px ;
}
.style6 {
color: #FF0000;
font-size: 13px ;
}
.style66 {
color: #FF0000;
font-size: 12px ;
}
.style67 {
color: #0033CC;
font-size: 11px ;
}
.btn2 {font-size: 13px}
.Text1{
font-weight:bold;
}
.Text2{
line-height:23px ;
font-size: 20px ;
font-weight:bold;
}
.styleA2 {font-size: 28px; line-height:100%;}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="myForm" method="post"> 
<table width='1200' border='1' align="left" cellpadding="0">
	<tr>
		<td rowspan="3" valign="top" width="65%">
		<!-- 影像大圖 -->
	<%if not rs1.eof Then
		file1=""
		file2=""
		file3=""
		file4=""
		If Trim(request("Speed"))="1" Then
			strImgFile="select * from BILLILLEGALIMAGETemp3 where billSn="&Trim(rs1("SN"))
		Else
			strImgFile="select * from BILLILLEGALIMAGETemp2 where billSn="&Trim(rs1("SN"))
		End If 		
		Set rsImgFile=conn.execute(strImgFile)
		If Not rsImgFile.eof Then
			If Trim(rsImgFile("IMAGEFILENAMEA"))<>"" Then
				file1= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEA"))
			End If 
			If Trim(rsImgFile("IMAGEFILENAMEB"))<>"" Then
				file2= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEB"))
			End If
			If Trim(rsImgFile("IMAGEFILENAMEC"))<>"" Then
				file3= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEC"))
			End If
			If Trim(rsImgFile("IMAGEFILENAMED"))<>"" Then
				file4= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMED"))
			End If
		End If 
		rsImgFile.close
		Set rsImgFile=Nothing 

	%>
		<input type="hidden" name="ImageFileNameA" value="<%
		if file1<>"" Then
			ImageFileNameAArray=Split(file1,"/")
			response.write ImageFileNameAArray(UBound(ImageFileNameAArray))
			ImageFileNameATemp=ImageFileNameAArray(UBound(ImageFileNameAArray))
			ImageFileNameTemp="/ReportCaseImage" & Replace(file1,ImageFileNameAArray(UBound(ImageFileNameAArray)),"")
		End if
		%>">
		<input type="hidden" name="ImagePathName" value="<%=ImageFileNameTemp%>">

		<%if file1<>"" then%>
		<%
		If UCase(Right(file1,3))="BMP" Or UCase(Right(file1,3))="PNG" Or UCase(Right(file1,3))="JPG" Or UCase(Right(file1,4))="JPEG" Or UCase(Right(file1,3))="GIF" Then
			IsPicture="1"
		Else
			IsPicture="0"
		End If 
		
		bPicWebPath=file1&"?nowTime="&now
		If IsPicture="1" then
			%>
			<img src="<%=bPicWebPath%>" border=1 height="<%
			response.write "460"
			%>" <%
			'放大鏡功能
			if isBig="Y"  then
			%>onmousemove="show(this, '<%=bPicWebPath%>')" onmousedown="show(this, '<%=bPicWebPath%>')"<%
			end if
			%> id="imgSource"> 
			<!-- <input type="button" name="btnImgNoUseA" value="相片無效" onclick="setImageNotUse('A');"> -->
			<input type="hidden" name="chkImgNoUseA" value="1">
			
		<%else%>
			<a href="<%=bPicWebPath%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameATemp,14)
			
			%></a>
		<%End If %>
			<div id="div1" style="position:absolute; overflow:hidden; width:<%
			'If sys_City=ApconfigureCityName Then
				If sys_City="雲林縣" Then
					response.write "750"
				Else
					response.write "230"
				End If 
				
			'Else
			'	response.write "210"
			'End If 
			%>px; height:<%
			'If sys_City=ApconfigureCityName Then
				If sys_City="雲林縣" Then
					response.write "140"
				Else
					response.write "110"
				End If 
				
			'Else
			'	response.write "90"
			'End If 
			%>px; left:<%
			if trim(request("divX"))="" Then
				If sys_City="台中市" Then
					response.write "780"
				elseIf sys_City="雲林縣" Then
					response.write "400"
				Else
					response.write "800"
				End If 
			else
				response.write trim(request("divX"))
			end if
			%>px; top:<%
			if trim(request("divY"))="" Then
				If sys_City="台中市" Then
					response.write "360"
				elseIf sys_City="雲林縣" Then
					response.write "320"
				Else
					response.write "360"
				End If 
				
			else
				response.write trim(request("divY"))
			end if
			%>px; z-index:1;border-right: white thin ridge; border-top: white thin ridge; border-left: white thin ridge; border-bottom: white thin ridge <%
		'放大鏡功能
		if isBig="N" Or IsPicture="0" then
		%> ;visibility: hidden;<%
		end if
		%>" onMousedown="initializedragie( )">
				<img id="BigImg" style='position:relative' src="<%=bPicWebPath%>">
			
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
		<td height="100" width="23%" align="center">
	<%if not rs1.eof Then
		if file2<>"" Then
	%>

		<input type="hidden" name="ImageFileNameB" value="<%
			ImageFileNameBarray=Split(file2,"/")
			response.write ImageFileNameBarray(UBound(ImageFileNameBarray))
			ImageFileNameBTemp=ImageFileNameBarray(UBound(ImageFileNameBarray))
		%>">
		<!-- 影像小圖 B-->
		<%
			If UCase(Right(file2,3))="BMP" Or UCase(Right(file2,3))="PNG" Or UCase(Right(file2,3))="JPG" Or UCase(Right(file2,4))="JPEG" Then
				IsPictureB="1"
			Else
				IsPictureB="0"
			End If 
			sPicWebPath2=file2&"?nowTime="&now

			If IsPictureB="1" then
		%>
		<img src="<%=sPicWebPath2%>" border=1 id="SmallImgB" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>" <%
			response.write "ondblclick=""ChangeImgB()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath2&"')"""
		%>>
			<!-- <input type="button" name="btnImgNoUseB" value="相片無效" onclick="setImageNotUse('B');"> -->
			<input type="hidden" name="chkImgNoUseB" value="1">
			<%else%>
			<a href="<%=sPicWebPath2%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameBTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%end if%>
		</td>
		<td rowspan="3" valign="top" style="width:23%;" >
			<table border='1' style="width:100%">
				<tr><td bgcolor="#FFFFCC"><%if sys_City="台中市" or sys_City="屏東縣" then%>當<%else%>七<%end if%>日內舉發案件</td></tr>
			</table>
			<div id="Layer1f1" style="overflow:auto; width:140px; height:300px; ">
	<%if not rs1.eof Then
		if sys_City="台中市" or sys_City="屏東縣" then
			RecDate1=Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate"))
			RecDate2=Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate"))
		else
			RecDate1=DateAdd("d",-7,Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate")))
			RecDate2=DateAdd("d",7,Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate")))
		end if 
		SqlRule2Plus=""
		RepeatBill=0
		If Trim(rs1("Rule2"))<>"" Then
			If Left(Trim(rs1("Rule1")),2)<>Left(Trim(rs1("Rule2")),2) Then
				SqlRule2Plus=" or Rule1 like '%"&Left(Trim(rs1("Rule2")),2)&"%' or Rule2 like '%"&Left(Trim(rs1("Rule2")),2)&"%'"
			End If 
		End If 
		strRB="select * from billbase where IllegalDate between to_date('"&RecDate1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
			" and to_date('"&RecDate2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
			" and CarNo='"&Trim(rs1("CarNo"))&"'" &_
			" and recordstateID=0"
		set rsRB=conn.execute(strRB)
		while Not rsRB.eof
			RepeatBill=1
			response.write "<a href='../Query/BillBaseData_Detail.asp?BillSN="&Trim(rsRB("Sn"))&"&BillType=0' target='_blank' >"&Trim(rsRB("BillNO"))
			response.write ginitdt(Trim(rsRB("IllegalDate")))&" "&Right("00"&hour(rsRB("IllegalDate")),2)&Right("00"&minute(rsRB("IllegalDate")),2)&"<br>"
			response.write Trim(rsRB("IllegalAddress"))&"</a><br><br>"
			rsRB.movenext
		wend
		rsRB.close
		set rsRB=nothing 

		strRB="select * from "&BillBaseName&" where IllegalDate between to_date('"&RecDate1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
			" and to_date('"&RecDate2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
			" and CarNo='"&Trim(rs1("CarNo"))&"'" &_
			" and recordstateID=0 and BillSn is null and sn<>"&Trim(rs1("sn"))
		set rsRB=conn.execute(strRB)
		while Not rsRB.eof
			RepeatBill=1
			If Trim(request("Speed"))="1" then
				response.write "<a href='../ReportCase/ReportCase_Detail_Speed.asp?ReportCaseSn="&Trim(rsRB("Sn"))&"' target='_blank' >"&Trim(rsRB("BillNO"))
			Else
				response.write "<a href='../ReportCase/ReportCase_Detail_TC.asp?ReportCaseSn="&Trim(rsRB("Sn"))&"' target='_blank' >"&Trim(rsRB("BillNO"))
			End if
			response.write ginitdt(Trim(rsRB("IllegalDate")))&" "&Right("00"&hour(rsRB("IllegalDate")),2)&Right("00"&minute(rsRB("IllegalDate")),2)&"<br>"
			response.write Trim(rsRB("IllegalAddress"))&"</a><br><br>"
			rsRB.movenext
		wend
		rsRB.close
		set rsRB=nothing 

		'response.write strRB
	End if
	%>		</div>
	<%if sys_City="台中市" And Trim(request("Speed"))<>"1" then%>
			<table border='1' style="width:100%">
				<tr><td bgcolor="#FFFFCC">相同標示單號案件</td></tr>
			</table>
			<div id="Layer1f1" style="overflow:auto; width:140px; height:125px; ">
	<%
		RepeatReportNo=0
		strRNo="select * from BillReportNoTemp where billsn="&trim(rs1("SN"))
		Set rsRNO=conn.execute(strRNo)
		If Not rsRNO.eof Then
			OldReportNo=Trim(rsRNO("ReportNo"))
		End If
		rsRNO.close
		Set rsRNO=Nothing
		
		If OldReportNo<>"" Then
			strChkRN="select * from billbase where recordstateid=0 and Exists (" &_
			"select billsn from BillReportNo where ReportNo='"&OldReportNo&"' and Billsn=billbase.SN" &_
			") and sn<>"&Trim(rs1("sn"))
			'response.write strChkRN
			Set rsChkRN=conn.execute(strChkRN)
			while Not rsChkRN.eof
				RepeatReportNo=1
				response.write "<a href='../Query/BillBaseData_Detail.asp?BillSN="&Trim(rsChkRN("Sn"))&"&BillType=0' target='_blank' >"&Trim(rsChkRN("BillNO"))&"<br>"
				response.write ginitdt(Trim(rsChkRN("IllegalDate")))&" "&Right("00"&hour(rsChkRN("IllegalDate")),2)&Right("00"&minute(rsChkRN("IllegalDate")),2)&"<br>"
				response.write Trim(rsChkRN("IllegalAddress"))&"</a><br><br>"
				rsChkRN.movenext
			wend
			rsChkRN.close
			Set rsChkRN=Nothing 

			strChkRN="select * from billbaseTmp where recordstateid=0 and BillSn is null and Exists (" &_
			"select billsn from BillReportNoTemp where ReportNo='"&OldReportNo&"' and Billsn=billbaseTmp.SN" &_
			") and sn<>"&Trim(rs1("sn"))
			'response.write strChkRN
			Set rsChkRN=conn.execute(strChkRN)
			while Not rsChkRN.eof
				RepeatReportNo=1
				response.write "<a href='../ReportCase/ReportCase_Detail_TC.asp?ReportCaseSn="&Trim(rsChkRN("Sn"))&"' target='_blank' >"&Trim(rsChkRN("BillNO"))&"<br>"
				response.write ginitdt(Trim(rsChkRN("IllegalDate")))&" "&Right("00"&hour(rsChkRN("IllegalDate")),2)&Right("00"&minute(rsChkRN("IllegalDate")),2)&"<br>"
				response.write Trim(rsChkRN("IllegalAddress"))&"</a><br><br>"
				rsChkRN.movenext
			wend
			rsChkRN.close
			Set rsChkRN=Nothing 
		End If 
	%>
			</div>
	<%End if%>
		</td>
	</tr>
	<tr>
		<td height="100" align="center">
	<%if not rs1.eof Then
		if file3<>"" Then
	%>
		<input type="hidden" name="ImageFileNameC" value="<%
			ImageFileNameCarray=Split(file3,"/")
			response.write ImageFileNameCarray(UBound(ImageFileNameCarray))
			ImageFileNameCTemp=ImageFileNameCarray(UBound(ImageFileNameCarray))
		%>">
		<!-- 影像小圖 C-->
		<%
			If UCase(Right(file3,3))="BMP" Or UCase(Right(file3,3))="PNG" Or UCase(Right(file3,3))="JPG" Or UCase(Right(file3,4))="JPEG" Then
				IsPictureC="1"
			Else
				IsPictureC="0"
			End If 

			sPicWebPath=file3&"?nowTime="&now
			If IsPictureC="1" then
		%>
		<img src="<%=sPicWebPath%>" border=1 id="SmallImgC" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>"  <%
			response.write "ondblclick=""ChangeImgC()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath&"')"""
		%>>
			<!-- <input type="button" name="btnImgNoUseC" value="相片無效" onclick="setImageNotUse('C');"> -->
			<input type="hidden" name="chkImgNoUseC" value="1">
			<%else%>
			<a href="<%=sPicWebPath%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameCTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="100" align="center">
	<%if not rs1.eof Then
		if file4<>"" Then
	%>
		<input type="hidden" name="ImageFileNameD" value="<%
			ImageFileNameDarray=Split(file4,"/")
			response.write ImageFileNameDarray(UBound(ImageFileNameDarray))
			ImageFileNameDTemp=ImageFileNameDarray(UBound(ImageFileNameDarray))
		%>">
		<!-- 影像小圖 D-->
		<%
			If UCase(Right(file4,3))="BMP" Or UCase(Right(file4,3))="PNG" Or UCase(Right(file4,3))="JPG" Or UCase(Right(file4,4))="JPEG" Then
				IsPictureD="1"
			Else
				IsPictureD="0"
			End If 

			sPicWebPath3=file4&"?nowTime="&now

			If IsPictureD="1" then
		%>
		<img src="<%=sPicWebPath3%>" border=1 id="SmallImgD" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>" <%
'		If (sys_City="宜蘭縣" And Trim(Session("Unit_ID"))="TQ00") Or sys_City="高雄市" Then
'			response.write "ondblclick=""ChangeImg()"""
'		Else
			response.write "ondblclick=""ChangeImgD()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath3&"')"""
'		End If 
		%>>
			<!-- <input type="button" name="btnImgNoUseD" value="相片無效" onclick="setImageNotUse('D');"> -->
			<input type="hidden" name="chkImgNoUseD" value="1">
			<%else%>
			<a href="<%=sPicWebPath3%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameDTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="100" colspan="3" valign="top">
		<%if not rs1.eof then%>
		<table width='100%' border='1' align="left" cellpadding="0">
			<tr>
				<td bgcolor="#FFFFCC" width="6%"><div align="right"> <span class="style3">＊</span>車號&nbsp;</div></td>
				<td >
				<input type="text" size="9" name="CarNo" onBlur="getVIPCar();" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
				end if
				%>" style=ime-mode:disabled maxlength="8" class="Text2" onkeydown="funTextControl(this);">
				<span class="style6">
			    <div id="Layer7" style="position:absolute; width:70px; height:24px; z-index:0;  border: 1px none #000000; color: #FF0000; font-weight: bold;"><%
			if trim(Session("SpecUser"))="1" then
				strSC="select count(*) as cnt from SpecCar where CarNo='"&trim(rs1("CarNo"))&"' and RecordStateID<>-1"
				set rsSC=conn.execute(strSC)
				if not rsSC.eof then
					if trim(rsSC("cnt"))<>"0" then
						response.write "＊業管車輛"
					end if
				end if
				rsSC.close
				set rsSC=nothing
			end if
				%></div>
				</span>
				
				</td>
				<td bgcolor="#FFFFCC" width="8%"><div align="right"><span class="style3">＊</span>車種&nbsp;</div>
				</td>
				<td colspan="3" >
                    <!-- 簡式車種 -->
                    <input type="text" maxlength="1" size="2" value="<%
                    if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
                    	response.write trim(rs1("CarSimpleID"))
                    end if
                    %>" name="CarSimpleID" onBlur="getRuleAll();" style=ime-mode:disabled onkeydown="funTextControl(this);">
                    <div id="Layer012" style="display: inline; width:300px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1汽車 / 2拖車 / 3重機/ 4輕機/5動力機械/6臨時車牌</font></div>
				</td>
			<%If Trim(request("Speed"))="1" then%>
				<td bgcolor="#FFFFCC" width="7%"><div align="right"><span class="style3">＊</span>違規時間(起)</div></td>
				<td width="13%">
					<!-- 違規日期(起) -->
					<input type="text" size="7" maxlength="7" name="StartIllegalDate" class='Text1' value="<%
					if trim(rs1("StartIllegalDate"))<>"" and not isnull(rs1("StartIllegalDate")) then 
						response.write gInitDT(rs1("StartIllegalDate"))
					end If
					%>" style=ime-mode:disabled onkeydown="funTextControl(this);" onkeyup="IllegalDateKeyUP2()" >&nbsp;
					<!-- 違規時間(起)區間 -->
					<input type="text" size="6" maxlength="6" name="StartIllegalTime" class='Text1' value="<%
					if trim(rs1("StartIllegalDate"))<>"" and not isnull(rs1("StartIllegalDate")) then 
						response.write Right("00"&hour(rs1("StartIllegalDate")),2)&Right("00"&minute(rs1("StartIllegalDate")),2)&Right("00"&Second(rs1("StartIllegalDate")),2)
					end if
					%>" onBlur="this.value=this.value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" onKeyUP="IllegalTimeKeyUP2('1')">
				</td>
			<%End if%>
				<td bgcolor="#FFFFCC" width="7%"><div align="right"><span class="style3">＊</span>違規時間<%
				if Trim(request("Speed"))="1" then
					response.write "(迄)"
				End If 
				%></div></td>
				<td width="13%" <%if sys_City<>"台中市"  and Trim(request("Speed"))<>"1" then %>colspan="3"<%End if%>>
					<!-- 違規日期 -->
					<input type="text" size="7" maxlength="7" name="IllegalDate" class='Text1' value="<%
					if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
						response.write gInitDT(rs1("IllegalDate"))
					end If
					%>" onBlur="getBillFillDate()" style=ime-mode:disabled onkeydown="funTextControl(this);" onkeyup="IllegalDateKeyUP()" >&nbsp;
				<%If Trim(request("Speed"))="1" then%>
					<!-- 違規時間區間 -->
					<input type="text" size="6" maxlength="6" name="IllegalTime" class='Text1' value="<%
					if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
						response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)&Right("00"&Second(rs1("IllegalDate")),2)
					end if
					%>" onBlur="this.value=this.value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" onKeyUP="IllegalTimeKeyUP2('2')">
				<%else%>
					<!-- 違規時間 -->
					<input type="text" size="3" maxlength="4" name="IllegalTime" class='Text1' value="<%
					if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
						response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
					end if
					%>" onBlur="this.value=this.value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" onKeyUP="IllegalTimeKeyUP()">
				<%End if%>
				</td>
			<%if sys_City="台中市" and Trim(request("Speed"))<>"1" then %>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">標示單號</div></td>
					<td >
						<input type="text" size="12" name="ReportNo" onkeydown="funTextControl(this);" value="<%
				
					response.write OldReportNo
					%>" style=ime-mode:disabled maxlength="11">
						<input type="hidden" name="OldReportNo" onkeydown="funTextControl(this);" value="<%
				
					response.write OldReportNo
					%>">
					</td>
			<%End If %>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>地點&nbsp;</div></td>
				<td colspan="3">
					<input type="text" size="4" value="<%
					response.write Trim(rs1("IllegalAddressID"))
					%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<%if sys_City="台中市" Or sys_City="高雄市" then %>
						<!-- 區號 -->
						<input type="hidden" class="btn5" size="3" value="<%=Trim(rs1("IllegalZip"))%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						
						<!-- <img src="../Image/BillkeyInButtonsmall.jpg" onclick="QryIllegalZip();"> -->
						<div id="LayerIllZip" style="display: inline; width:160px; height:30; z-index:0;  border: 1px none #000000;"><%
					if Trim(rs1("IllegalZip"))<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&Trim(rs1("IllegalZip"))&"'"
						set rsZip1=conn.execute(strZip1)
						if not rsZip1.eof then
							response.write trim(rsZip1("ZipName"))
						end if
						rsZip1.close
						set rsZip1=nothing
					end if
					%></div>
					<%end if%>
					<input type="text" size="40" value="<%
					if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
						response.write trim(rs1("IllegalAddress"))
					end If
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%
					if trim(request("chkHighRoad"))="1" then 
						response.write "checked"
					ElseIf trim(rs1("IllegalAddress"))="台61苗栗段南下-122.2K至130.2K" then
						response.write "checked"
					ElseIf trim(rs1("IllegalAddress"))="台61苗栗段北上-122.6K至115.6K" then
						response.write "checked"
					ElseIf trim(rs1("IllegalAddress"))="台61台中段南下-148.1K至156.4K" then
						response.write "checked"
					ElseIf trim(rs1("IllegalAddress"))="台西鄉台61線快速公路(北往南)218K+250m起點至226K+250m終點全長8公里" then
						response.write "checked"
					ElseIf trim(rs1("IllegalAddress"))="台61彰化段北上-177K至168.1K" then
						response.write "checked"
					elseif left(trim(rs1("Rule1")),5)="33101" then
						response.write "checked"
					End If 
					%> onclick="setIllegalRule()" <%if sys_City="南投縣" then response.write "disabled"%>>
					<div id="Layerert45" style="display: inline; width:30px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><span class="style1">快速道路</span></div>
                </td>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style3">＊</span>法條一</div></td>
				<td colspan="5">
					<input type="text" maxlength="9" size="7" value="<%
					if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
						response.write trim(rs1("Rule1"))
					end If
					%>" name="Rule1" onKeyUp="getRuleData1();" style=ime-mode:disabled onkeydown="funTextControl(this);" >
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<img src="../Image/BillLawPlusButton_Small.JPG" onclick="Add_LawPlus()" alt="附加說明">
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
				if sys_City="基隆市" Then
					If Left(trim(rs1("Rule1")),2)="40" Or Left(trim(rs1("Rule1")),5)="43102" Or Left(trim(rs1("Rule1")),5)="33101" then
						if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) And trim(rs1("IllegalSpeed"))<>"0" then
							response.write trim(rs1("IllegalSpeed"))
						end If
					End if
				Else
					if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) And trim(rs1("IllegalSpeed"))<>"0" then
						response.write trim(rs1("IllegalSpeed"))
					end If
				End If 
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
				if sys_City<>"基隆市" And sys_City<>"屏東縣" Then
					If Left(trim(rs1("Rule1")),2)="40" Or Left(trim(rs1("Rule1")),5)="43102" Or Left(trim(rs1("Rule1")),5)="33101" then
						if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) And trim(rs1("RuleSpeed"))<>"0" then
							response.write trim(rs1("RuleSpeed"))
						end If
					End if
				Else
					if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) And trim(rs1("RuleSpeed"))<>"0" then
						response.write trim(rs1("RuleSpeed"))
					end If
				End if
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					&nbsp;
					<span class="style5">
					<div id="Layer1" style="display: inline;position:absolute ; width:230px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><%
					strR1="select * from Law where itemid='"&trim(rs1("Rule1"))&"' and Version=2"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof Then
						response.write rsR1("IllegalRule")
					End If 
					rsR1.close
					Set rsR1=Nothing 
					if trim(rs1("Rule4"))<>"" then
						response.write "("&trim(rs1("Rule4"))&")"
					ElseIf Trim(Session("SpeedKeyIn_Rule4"))<>"" Then
						response.write "("&trim(Session("SpeedKeyIn_Rule4"))&")"
					end if 
					%></div></span>
					<input type="hidden" name="ForFeit1" value="<%
					response.write trim(rs1("ForFeit1"))
					%>">
					<input type="hidden" value="<%
					if trim(rs1("Rule4"))<>"" then
						response.write trim(rs1("Rule4"))
					ElseIf Trim(Session("SpeedKeyIn_Rule4"))<>"" Then
						response.write trim(Session("SpeedKeyIn_Rule4"))
					end if 
					%>" name="Rule4">
				</td>
		    </tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right">法條二</div></td>
				<td colspan="3">
					<input type="text" maxlength="9" size="7" value="<%
					if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						response.write trim(rs1("Rule2"))
					end If
					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer2" style="display: inline;position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%
					strR1="select * from Law where itemid='"&trim(rs1("Rule2"))&"' and Version=2"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof Then
						response.write rsR1("IllegalRule")
					End If 
					rsR1.close
					Set rsR1=Nothing 
					%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%
					response.write trim(rs1("ForFeit2"))
					%>">

				</td>
				<%If Trim(request("Speed"))="1" then%>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><%
				if sys_City<>"基隆市" then
				%><span class="style3">＊</span><%
				End if
				%>距離(公尺)&nbsp;</div></td>
		  		<td>
					<input type="text" name="Distance" value="<%
						if trim(rs1("Distance"))<>"" and not isnull(rs1("Distance")) then 
							response.write trim(rs1("Distance"))
						end If
						%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">

					<input type="hidden" name="JurgeDay" value="" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">
				</td>
				<%End if%>
				<%if sys_City<>"台中市" then %>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><span class="style3">＊</span>舉發人&nbsp;</div></td>
		  		<td colspan="<%
				If Trim(request("Speed"))<>"1" then
					response.write "3"
				End if
				%>">
					<input type="text" size="9" name="BillMem1" value="<%

				If Trim(Session("SpeedKeyIn_BillMem1")) <>"" Then

					Response.Write Trim(Session("SpeedKeyIn_BillMem1"))

				ElseIf Trim(rs1("BillMemID1"))<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(rs1("BillMemID1"))
					Set rsMem=conn.execute(strMem)
					If Not rsMem.eof Then
						response.write Trim(rsMem("LoginID"))
					End If
					rsMem.close
					Set rsMem=nothing 
				End If 
				%>" onKeyUp="getBillMemID1();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer12" style="display: inline; width:60px; height:30;  z-index:0;  border: 1px none #00000; "><%
				If Trim(Request("BillMemName1")) <>"" And sys_City="台東縣" Then

					Response.Write Trim(Request("BillMemName1"))

				ElseIf Trim(rs1("BillMem1"))<>"" Then
					response.write Trim(rs1("BillMem1"))
				End If 
					%></div>
					</span>
					<input type="hidden" value="<%
				If Trim(Request("BillMemID1")) <>"" And sys_City="台東縣" Then

					Response.Write Trim(Request("BillMemID1"))

				ElseIf Trim(rs1("BillMemID1"))<>"" Then
					response.write Trim(rs1("BillMemID1"))
				End If 
					%>" name="BillMemID1">
					<input type="hidden" value="<%
				If Trim(Request("BillMemName1")) <>"" And sys_City="台東縣" Then

					Response.Write Trim(Request("BillMemName1"))

				ElseIf Trim(rs1("BillMem1"))<>"" Then
					response.write Trim(rs1("BillMem1"))
				End If 
					%>" name="BillMemName1">
				<%End if%>
				<%if sys_City="台中市" And Trim(request("Speed"))<>"1" then %>
					
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">民眾檢舉日期</div></td>
					<td >
						<input type="text" name="JurgeDay" value="<%
						if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then 
							response.write gInitDT(rs1("JurgeDay"))
						end If
						%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">
					</td>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">民眾檢舉案號</div></td>
					<td >
						<input type="text" name="ReportCaseNo" value="<%
						if trim(rs1("ReportCaseNo"))<>"" and not isnull(rs1("ReportCaseNo")) then
							response.write trim(rs1("ReportCaseNo"))
						end if
						%>" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:100px;" onblur="this.value=this.value.toUpperCase()">

						<input type="hidden" size="7" name="BillMem2" value="<%%>" style=ime-mode:disabled onkeydown="funTextControl(this);">
						<input type="hidden" value="<%%>" name="BillMemID2">
						<input type="hidden" value="<%%>" name="BillMemName2">
					</td>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">檢舉人證號</div></td>
					<td >
						<input type="text" name="ReportCreditID" value="<%
						if trim(rs1("ReportCreditID"))<>"" and not isnull(rs1("ReportCreditID")) then
							response.write trim(rs1("ReportCreditID"))
						end if
						%>" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:100px;" onblur="this.value=this.value.toUpperCase()">
					</td>
				<%ElseIf sys_City<>"台中市" then%>

					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">舉發人二</div></td>
					<td >
						
						<input type="text" size="7" name="BillMem2" value="<%
				If Trim(rs1("BillMemID2"))<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(rs1("BillMemID2"))
					Set rsMem=conn.execute(strMem)
					If Not rsMem.eof Then
						response.write Trim(rsMem("LoginID"))
					End If
					rsMem.close
					Set rsMem=nothing 
				End If 
					%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
						<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
						<span class="style5">
						<div id="Layer13" style="display: inline; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%
				If Trim(rs1("BillMem2"))<>"" Then
					response.write Trim(rs1("BillMem2"))
				End If 
						%></div>
						</span>
						<input type="hidden" value="<%
				If Trim(rs1("BillMemID2"))<>"" Then
					response.write Trim(rs1("BillMemID2"))
				End If 
						%>" name="BillMemID2">
						<input type="hidden" value="<%
				If Trim(rs1("BillMem2"))<>"" Then
					response.write Trim(rs1("BillMem2"))
				End If 
						%>" name="BillMemName2">
					</td>
				<%End if%>
			</tr>
			<tr>
				<%if sys_City="台中市" then %>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><span class="style3">＊</span>舉發人&nbsp;</div></td>
		  		<td >
					<input type="text" size="9" name="BillMem1" value="<%
				If Trim(rs1("BillMemID1"))<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(rs1("BillMemID1"))
					Set rsMem=conn.execute(strMem)
					If Not rsMem.eof Then
						response.write Trim(rsMem("LoginID"))
					End If
					rsMem.close
					Set rsMem=nothing 
				End If 
				%>" onKeyUp="getBillMemID1();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer12" style="display: inline; width:60px; height:30;  z-index:0;  border: 1px none #00000; "><%
				If Trim(rs1("BillMem1"))<>"" Then
					response.write Trim(rs1("BillMem1"))
				End If 
					%></div>
					</span>
					<input type="hidden" value="<%
				If Trim(rs1("BillMemID1"))<>"" Then
					response.write Trim(rs1("BillMemID1"))
				End If 
					%>" name="BillMemID1">
					<input type="hidden" value="<%
				If Trim(rs1("BillMem1"))<>"" Then
					response.write Trim(rs1("BillMem1"))
				End If 
					%>" name="BillMemName1">
				<%End if%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span><span class="style4">舉發單位</span></div></td>
				<td colspan="<%
				if sys_City<>"台中市" Then
					response.write "3"
				End If 
				%>">
					<input type="text" size="4" name="BillUnitID" value="<%=Trim(rs1("BillUnitID"))%>" onKeyUp="getUnit();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="display: inline; width:200px; height:30px; z-index:0;  border: 1px none #000000; "><%
					if Trim(rs1("BillUnitID"))<>"" then
						strUnitName="select UnitName from UnitInfo where UnitID='"&Trim(rs1("BillUnitID"))&"'"
						set rsUnitName=conn.execute(strUnitName)
						if not rsUnitName.eof then
							response.write trim(rsUnitName("UnitName"))
						end if
						rsUnitName.close
						set rsUnitName=nothing
					end if
					%></div>
					</span>
				<%if sys_City<>"台中市" And Trim(request("Speed"))<>"1" then %>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span class="style4">民眾檢舉時間</span>
					<input type="text" name="JurgeDay" value="<%
					if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then 
						response.write gInitDT(rs1("JurgeDay"))
					end If
					%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">
				<%End If %>
				</td>
				<td bgcolor="#FFFFCC" width="8%">

				<div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>

				<div align="right"><span class="style3">＊</span>填單日期</div></td>
				<td width="9%">
				
				&nbsp;<input type="text" size="6" value="<%=ginitdt(date)%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">

				<input type="hidden" name="SelSN" value="<%=trim(rs1("SN"))%>">

				</td>

				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種&nbsp;</td>
				<td width="6%">
                &nbsp;<input type="text" maxlength="2" size="4" value="<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then 
					response.write rs1("CarAddID")
				end If
				%>" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
                
				</td>

				<td bgcolor="#FFFFCC" width="8%">
		
				<div align="right">專案代碼&nbsp;</div></td>
				<td width="12%">
					&nbsp;<input type="text" size="5" value="<%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then 
					response.write rs1("ProjectID")
				end If
				%>" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>

					<!-- <div id="Layer5012" style="position:absolute; width:125px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1檢舉達人<br>&nbsp;9拖吊</font></div> -->

					<!-- 採証工具 -->
					<input maxlength="1" size="4" value="0" name="UseTool"  onkeyup="getFixID();" type='hidden' style=ime-mode:disabled> 
			        <div id="Layer11" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%
					'if trim(rs1("SiteCode"))<>"" and not isnull(rs1("SiteCode")) then
					'	response.write trim(rs1("SiteCode"))
					'end if
					%>' onBlur="setFixEquip();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					</div>
					<!-- <font color="#ff000" size="2"> 1固定桿/ 2雷達三腳架/ 3相機</font> -->
			    
				</td>

			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right" width="8%">備註&nbsp;</td>
				<td colspan="9">
				<input type="Text" size="40" value="<%
					if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
						response.write trim(rs1("Note"))
					end if
					%>" name="Note" style=ime-mode:active>
				</td>
			</tr>
		</table>
		<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
		<%end if%>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td height="28" colspan="3" align="center">
		<%if Trim(rs1("BillStatus"))="8" then%>
			<input type="button" name="Submit2932" onClick="funDeleteCase();" value="確認刪除案件" <%
		if rs1.eof then
			response.write "disabled"
		Elseif Trim(rs1("Recordstateid"))="-1" Then
			response.write "disabled"
		end If
		
			%> style="font-size: 10pt; width: 100px; height: 27px">
		<%end if%>
			<input type="button" value="審核通過 F2" onclick="InsertBillVase();"  <%
		RecordStateIDTemp=""
		If Trim(rs1("BillSn"))<>"" Then
			strBill="select * from billbase where sn="& Trim(rs1("BillSn"))
			Set rsBill=conn.execute(strBill)
			If Not rsBill.eof Then
				RecordStateIDTemp=Trim(rsBill("RecordStateID"))
			End If 
			rsBill.close
			Set rsBill=nothing 
		End If 
		
		If Trim(request("OtherCase"))="1" And Trim(request("Speed"))="1" And RecordStateIDTemp="-1" Then

		else
			if rs1.eof then
				response.write "disabled"
				checkF2Flag=1
			ElseIf Trim(rs1("CheckFlag"))<>"0" Then
				response.write "disabled"
				checkF2Flag=1
			elseif chkIllegaldate30day=1 then
				response.write "disabled"
				checkF2Flag=1
			elseif chkWarningNumber=1 then
				response.write "disabled"
				checkF2Flag=1
			elseif Trim(rs1("BillStatus"))="8" then
				response.write "disabled"
				checkF2Flag=1
			Elseif Trim(rs1("Recordstateid"))="-1" Then
				response.write "disabled"
				checkF2Flag=1
			end If
		End if
			%> style="font-size: 10pt; width: 100px; height: 27px">
			<%if chkIllegaldate30day=1 then%>
			<font color="red">(違規日超過上傳日30天)</font>
			<%end if%>
			<%if chkWarningNumber=1 then%>
			<font color="red">(於標示單管理中，查無此標示單號)</font>
			<%end if%>
			<input type="button" name="Submit2932" onClick="funVerifyResult();" value="審核無效 F9" <%
		if rs1.eof then
			response.write "disabled"
		ElseIf Trim(rs1("CheckFlag"))<>"0" Then
			response.write "disabled"
		Elseif Trim(rs1("Recordstateid"))="-1" Then
			response.write "disabled"
		end If
			%> style="font-size: 10pt; width: 100px; height: 27px">
		<%if Trim(rs1("BillStatus"))<>"8" And Trim(request("Speed"))<>"1" And sys_City<>"彰化縣" And sys_City<>"金門縣" then%>
			<input type="button" name="Submit2932" onClick="funDelResult();" value="直接刪除" <%
			if rs1.eof then
				response.write "disabled"
			ElseIf Trim(rs1("CheckFlag"))<>"0" Then
				response.write "disabled"
			end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
		<%end if%>
			<img src="/image/space.gif" width="29" height="8">
			<input type="hidden" name="kinds" value="">
		<%If Trim(request("OtherCase"))="1" And Trim(request("Speed"))="1" then%>

		<%else%>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=FirstSn%>&Speed=<%=Trim(request("Speed"))%>'" value="<< 第一筆 Home" style="font-size: 9pt; width: 90px; height: 27px" <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=UpSn%>&Speed=<%=Trim(request("Speed"))%>'" value="< 上一筆 PgUp" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<%=ThisSn+1 & " / " & AllSN%>
			<input type="button" name="SubmitNext" onClick="location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=DownSn%>&Speed=<%=Trim(request("Speed"))%>'" value="下一筆 PgDn >" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=LastSn%>&Speed=<%=Trim(request("Speed"))%>'" value="最後一筆 End >>" style="font-size: 9pt; width: 90px; height: 27px" <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>

		<%End if%>
			<input type="hidden" name="DownSn_Temp" value="<%=DownSn%>">
			&nbsp; &nbsp; 
		<%If (sys_City="台中市" Or sys_City="彰化縣" Or sys_City="金門縣") And Trim(request("Speed"))<>"1" Then%>
			<input type="button" name="Submitpic2932" onClick="funPictureList('<%=trim(request("CheckSn"))%>');" value="相片列表" <%
			if rs1.eof then
				response.write "disabled"
			end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
		<%End if%>
			<img src="/image/space.gif" width="29" height="8">
			<input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" style="font-size: 10pt; width: 70px; height: 27px">
			
			<img src="/image/space.gif" width="29" height="8">
			
			<%If sys_City="彰化縣" then%>
				<input type="checkbox" name="CaseInByMem" ><font style="font-size: 10pt">違規日逾期強制建檔</font>
			<%End If %>


             <input type="hidden" name="Tmp_Order" value="<%=Session("BillCnt_Image")%>">
				<input type="hidden" name="CheckSn" value="<%=Trim(request("CheckSn"))%>">				
				<input type="hidden" value="<%=Trim(request("Speed"))%>" name="Speed">
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案日期 -->
				<input type="hidden" size="12" maxlength="7" name="DealLineDate">
				<!-- 應到案處所 -->
				<input type="hidden" size="10" value="" name="MemberStation">
				<input type="hidden" value="" name="Rule3">
				<input type="hidden" name="ForFeit3" value="">
				
				<input type="hidden" name="ForFeit4" value="">
				<input type="hidden" name="Billno1" value="<%=Trim(rs1("Billno"))%>">
				<input type="hidden" value="" name="Insurance">
				<input type="hidden" value="" name="BillMemID3">
				<input type="hidden" value="" name="BillMemID4">
				<input type="hidden" value="" name="BillMemName3">
				<input type="hidden" value="" name="BillMemName4">
				<!-- <input type="button" value="？" name="StationSelect" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=660,height=375,resizable=yes,scrollbars=yes")'> -->
				<div id="Layer5" style="position:absolute ; width:221px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000; visibility :hidden;"></div>
				<input type="hidden" name="SessionFlag" value="1">
				<!--浮動視窗座標-->
				<input type="hidden" name="divX" value="<%
				if trim(request("divX"))="" then
					If sys_City=ApconfigureCityName Then
						response.write "650"
					elseIf sys_City="苗栗縣" Then
						response.write "1210"
					elseIf sys_City="台中市" Then
						response.write "780"
					Else
						response.write "800"
					End If 
				else
					response.write trim(request("divX"))
				end if
				%>">
				<input type="hidden" name="divY" value="<%
				if trim(request("divY"))="" Then
					If sys_City=ApconfigureCityName Then
						response.write "490"
					elseIf sys_City="苗栗縣" Then
						response.write "40"
					elseIf sys_City="台中市" Then
						response.write "360"
					Else
						response.write "360"
					End If 
				else
					response.write trim(request("divY"))
				end if
				%>">
				
		</td>
	</tr>
<%
if sys_City="台中市" then
	If Trim(request("CheckSn"))<>"" Then
		StrRea="select * from (" &_
			"select * from ReportCaseCheckRecord where CheckFlag in ('2','6') and BillTempSN="&Trim(request("CheckSn"))&" order by Recorddate desc" &_
			") where rownum<=1"
		set rsRea=conn.execute(StrRea)
		while Not rsRea.eof
	%>
	<tr>
		<td colspan="2">
			<br>
			<%
			If Trim(rsRea("CheckFlag"))="2" Then
				response.write "審核未通過原因"
			ElseIf Trim(rsRea("CheckFlag"))="6" Then
				response.write "委外退件原因"
			End If 
			%>
			: <strong>
			<%response.write trim(rsRea("Note"))%>
			</strong>
			審核時間: <strong>
			<%response.write year(rsRea("RecordDate"))-1911&right("00"&month(rsRea("RecordDate")),2)&right("00"&day(rsRea("RecordDate")),2) & " &nbsp; " &right("00"&hour(rsRea("RecordDate")),2)&right("00"&minute(rsRea("RecordDate")),2)%>
			</strong>
		</td>
	</tr>
	<%
			
			rsRea.movenext
		wend
		rsRea.close
		set rsRea=nothing 
	End If 
End If 
%>
<%
if sys_City="彰化縣" then
%>
	<tr>
		<td colspan="2">
		<br>
		<span style="color: #FF0000;font-size: 18px;"><strong>( 注意事項 )</strong></span>
		<br>
		<span style="color: #FF0000;font-size: 18px;"><strong>1.舉發單只能印第一、二張相片</strong></span>
		<br>
		<span style="color: #FF0000;font-size: 18px;"><strong>2.如果要更換相片順序，請使用『相片列表』功能</strong></span>
		</td>
	</tr>
<%
End if
%>
</table>

</form>

</body>
<script language="JavaScript">
var TDLawNum=0;
var TDLawErrorLog1=0;
var TDLawErrorLog2=0;
var TDLawErrorLog3=0;
var TDLawErrorLog4=0;
var TDStationErrorLog=0;
var TDUnitErrorLog=0;
var TDFastenerErrorLog1=0;
var TDFastenerErrorLog2=0;
var TDFastenerErrorLog3=0;
var TDMemErrorLog1=0;
var TDMemErrorLog2=0;
var TDMemErrorLog3=0;
var TDMemErrorLog4=0;
var TDProjectIDErrorLog=0;
var TDIllZipErrorLog=0;
var TDVipCarErrorLog=0;
var SpeedError=0;
var TodayDate=<%=ginitdt(date)%>;

var InsertFlag=0;
<%if sys_City="宜蘭縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="屏東縣" and Trim(request("Speed"))="1" then%>
MoveTextVar("CarNo,CarSimpleID,StartIllegalDate,StartIllegalTime,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,RuleSpeed,IllegalSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="南投縣" Or sys_City="屏東縣" or sys_City="花蓮縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,RuleSpeed,IllegalSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="苗栗縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalSpeed,RuleSpeed,Rule1,Rule2||IllegalAddressID,IllegalAddress,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,ProjectID,CarAddID");
<%elseif sys_City="台中市" and Trim(request("Speed"))<>"1" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime,ReportNo||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,JurgeDay,ReportCaseNo,ReportCreditID||BillMem1,BillUnitID,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="台中市" and Trim(request("Speed"))="1" then%>
MoveTextVar("CarNo,CarSimpleID,StartIllegalDate,StartIllegalTime,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,Distance||BillMem1,BillUnitID,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="高雄市" and Trim(request("Speed"))="1" then%>
MoveTextVar("CarNo,CarSimpleID,StartIllegalDate,StartIllegalTime,IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,Distance,BillMem1||BillUnitID,BillFillDate,CarAddID,ProjectID");
<%elseif Trim(request("Speed"))="1" then%>
MoveTextVar("CarNo,CarSimpleID,StartIllegalDate,StartIllegalTime,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,Distance,BillMem1||BillUnitID,BillFillDate,CarAddID,ProjectID");
<%else%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%end if%>


//新增告發單
function InsertBillVase(){

	var error=0;
	var errorString="";
	var TodayDate=<%=ginitdt(date)%>;
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}else if (myForm.BillType.value=="2"){
		
		/*smith remark 逕舉不一定要輸入固定桿編號. 可能是員警拍照
		if (myForm.FixID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入固定桿編號。";
		}
		*/
	}
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}

	if (myForm.CarSimpleID.value==""){
		//error=error+1;
		//errorString=errorString+"\n"+error+"：請輸入簡式車種。";
	}//else if(myForm.CarNo.value != "" && chkCarNoFormat(myForm.CarNo.value)!= 0) {
	//	if (chkCarNoFormat(myForm.CarNo.value) != myForm.CarSimpleID.value){
	//		error=error+1;
	//		errorString=errorString+"\n"+error+"：車號格式與簡式車種不符。";
	//	}
	//}
	var IllDateFlag=0;
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
		IllDateFlag=1;
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="9" && myForm.IllegalDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="1" && myForm.IllegalDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if (!ChkIllegalDate60_109(myForm.IllegalDate.value)){
	<%If sys_City="彰化縣" then%>
		if (myForm.CaseInByMem.checked==false || myForm.Note.value=="")
		{
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期已超過二個月期限，如確定可舉發請勾選逾期強制建檔，並且在備註輸入原因。";
		}	
	<%else%>
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	<%end if %>
		
	}
<%if Trim(request("Speed"))="1" then%>
	if (myForm.Rule1.value.substr(0,2)=="40" || myForm.Rule1.value.substr(0,5)=="43102" || myForm.Rule1.value.substr(0,5)=="33101" || myForm.Rule2.value.substr(0,2)=="40" || myForm.Rule2.value.substr(0,5)=="43102" || myForm.Rule2.value.substr(0,5)=="33101")
	{
		if (myForm.StartIllegalDate.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入違規日期(起)。";
		}else if(!dateCheck( myForm.StartIllegalDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)輸入錯誤。";
		}else if( myForm.StartIllegalDate.value.substr(0,1)=="9" && myForm.StartIllegalDate.value.length==7 ){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)輸入錯誤。";
		}else if( myForm.StartIllegalDate.value.substr(0,1)=="1" && myForm.StartIllegalDate.value.length==6 ){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)輸入錯誤。";
	<%If sys_City="高雄市" then%>
		}else if (!ChkIllegalDate2M_KS(myForm.StartIllegalDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)已超過二個月期限。";
		}
	<%else%>
		}else if (!ChkIllegalDate60_109(myForm.StartIllegalDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)已超過二個月期限。";
		}
	<%end if%>

		if (myForm.StartIllegalTime.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入違規時間(起)。";
		}else if(myForm.StartIllegalTime.value.length < 4){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規時間(起)輸入錯誤。";
		}else if(myForm.StartIllegalTime.value.substr(0,2) > 23 || myForm.IllegalTime.value.substr(0,2) < 0){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規時間(起)輸入錯誤。";
		}else if(myForm.StartIllegalTime.value.substr(2,2) > 59 || myForm.IllegalTime.value.substr(2,2) < 0){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規時間(起)輸入錯誤。";
		}
		<%if sys_City<>"基隆市" then%>
		if (myForm.Distance.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入距離。";
		}
		<%end if%>
	}
	
<%else%>
	if (myForm.IllegalTime.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規時間。";
	}else if(myForm.IllegalTime.value.length < 4){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}else if(myForm.IllegalTime.value.substr(0,2) > 23 || myForm.IllegalTime.value.substr(0,2) < 0){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}else if(myForm.IllegalTime.value.substr(2,2) > 59 || myForm.IllegalTime.value.substr(2,2) < 0){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}
<%end if%>
	if (myForm.IllegalAddress.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點。";
	}
	if (myForm.Rule1.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規法條一。";
	}else if (TDLawErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
	}else if (myForm.Rule1.value.substr(0,2)>68){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
	}
	if (myForm.Rule1.value==myForm.Rule2.value && myForm.Rule1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一與違規法條二重複。";
	}
	if (myForm.Rule2.value!=""){
		if (TDLawErrorLog2==1){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}else if (myForm.Rule2.value.substr(0,2)>68){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}
	}
	if (myForm.BillFillDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入填單日期。";
	}else if (!dateCheck( myForm.BillFillDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="9" && myForm.BillFillDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="1" && myForm.BillFillDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
	}else if (!ChkIllegalDate60_109(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過二個月。";
	}
	if (myForm.MemberStation.value==""){
		if(myForm.BillType.value=="1"){
			//攔停才嗆破輸入
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案處所。";
		}
	}else if (TDStationErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案處所輸入錯誤。";
	}
	if (myForm.DealLineDate.value==""){
		if(myForm.BillType.value=="1"){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案日期。";
		}
	}else if (!dateCheck( myForm.DealLineDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate60_109(myForm.DealLineDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過二個月。";
	}
	if (myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發單位代號。";
		TDUnitErrorLog==0
	}else if (TDUnitErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發單位代號輸入錯誤。";
	}
	if (myForm.BillMem1.value==""){
		//固定桿不需要輸入舉發人
		//if (myForm.UseTool.value!="1"){
		    error=error+1;
			errorString=errorString+"\n"+error+"：請輸入舉發人1 代碼。";
		//}
	}else if (TDMemErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 輸入錯誤。";
	}
	if (TDMemErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 輸入錯誤。";
	}
	if (TDMemErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼3 輸入錯誤。";
	}
	if (TDMemErrorLog4==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼4 輸入錯誤。";
	}
	if (eval(myForm.BillFillDate.value) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
	}
	if (myForm.JurgeDay.value!=""){
		if(!dateCheck( myForm.JurgeDay.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：民眾檢舉時間輸入錯誤。";	
		}else if (IllDateFlag==0){
			Iyear=parseInt(myForm.IllegalDate.value.substr(0,myForm.IllegalDate.value.length-4))+1911;
			Imonth=myForm.IllegalDate.value.substr(myForm.IllegalDate.value.length-4,2);
			Iday=myForm.IllegalDate.value.substr(myForm.IllegalDate.value.length-2,2);
			var IllDate=new Date(Iyear,Imonth-1,Iday);

			Jyear=parseInt(myForm.JurgeDay.value.substr(0,myForm.JurgeDay.value.length-4))+1911;
			Jmonth=myForm.JurgeDay.value.substr(myForm.JurgeDay.value.length-4,2);
			Jday=myForm.JurgeDay.value.substr(myForm.JurgeDay.value.length-2,2);
			var JDay=new Date(Jyear,Jmonth-1,Jday);

			var OverDate=new Date();
			OverDate=DateAdd("d",6,IllDate);
			if (JDay > OverDate){
				error=error+1;
				errorString=errorString+"\n"+error+"：民眾檢舉時間已超過七天，民眾檢舉發生超過七日之交通違規，依法不得舉發。";	
			}
			if (JDay < IllDate){
				error=error+1;
				errorString=errorString+"\n"+error+"：民眾檢舉時間不可小於違規日。";
			}
		}
	}
	if (TDProjectIDErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：專案代碼輸入錯誤。";
	}
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if(parseInt(myForm.RuleSpeed.value) > parseInt(myForm.IllegalSpeed.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：限速、限重大於實際車速、車重。";
		}
		if ((myForm.Rule1.value.substr(0,3))!="293" && (myForm.Rule2.value.substr(0,3))!="293")	{
			if(parseInt(myForm.RuleSpeed.value) < 30){
				error=error+1;
				errorString=errorString+"\n"+error+"：限速、限重小於 30Km/h。";
			}
		}		
		if(parseInt(myForm.RuleSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：限速、限重大於 300Km/h。";
		}
		if(parseInt(myForm.IllegalSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：實際車速、車重大於 300Km/h。";
		}
		if((parseInt(myForm.IllegalSpeed.value)-parseInt(myForm.RuleSpeed.value) ) > 150){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速大於 150Km/h。";
		}
	}
<%if sys_City="高雄市" then%>
	if (SpeedError==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：超速 100~150Km/h ，請輸入密碼後才可建檔。";
	}
	if (myForm.IllegalSpeed.value!="" || myForm.RuleSpeed.value!=""){
		if ((myForm.Rule1.value.substr(0,5))!="33101" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule1.value.substr(0,5))!="43102" && (myForm.Rule2.value.substr(0,5))!="33101" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,5))!="43102" && (myForm.Rule2.value.substr(0,2))!="29"){
			error=error+1;
			errorString=errorString+"\n"+error+"：非超速、重法條，請勿輸入車速。";
		}
	}
	if (TDIllZipErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規地點區號 輸入錯誤。";
	}
	else if(myForm.IllegalZip.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點區號。";
	}
<%end if%>
	if ((myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"1");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}else if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"1");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}
	if ((myForm.Rule1.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}
<%if sys_City="雲林縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then %>
	if (TDVipCarErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：車號 "+myForm.CarNo.value+" 為業管車輛。";
	}
<%end if%>
<%if sys_City="台中市" and Trim(request("Speed"))<>"1" then %>
//	if ((myForm.Rule1.value.substr(0,2))=="55"){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：第55條不可逕行舉發。";
//	}
	if (myForm.ReportNo.value!=""){
		if (myForm.ReportNo.value.length<11){
			error=error+1;
			errorString=errorString+"\n"+error+"：告示單號不可少於11碼。";
		}
	}
<%end if%>
	if ((myForm.Rule1.value.substr(0,3))=="293" && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 40){
				if ((myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,5))=="33101"){
				<%If Trim(request("Speed"))="1" Then%>
					if (myForm.Rule1.value=="4340033" || myForm.Rule2.value=="4340033" || myForm.Rule1.value=="4340045" || myForm.Rule2.value=="4340045" || myForm.Rule1.value=="4340069" || myForm.Rule2.value=="4340069"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340033、4340045、4340069需另單舉發。";
					}
				<%else%>
					if (myForm.Rule1.value=="4340003" || myForm.Rule2.value=="4340003" || myForm.Rule1.value=="4340044" || myForm.Rule2.value=="4340044" || myForm.Rule1.value=="4340068" || myForm.Rule2.value=="4340068"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340003、4340044、4340068需另單舉發。";
					}					
				<%end if %>
				}
			}
		}
	}
<%if RepeatReportNo=1 then%>
	if (myForm.ReportNo.value==myForm.OldReportNo.value)
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：此標示單號已有其他案件使用，請先確認是否正確。";
	}
<%end if %>
<%if sys_City="雲林縣" then %>
	if (myForm.chkHighRoad.checked==true && myForm.IllegalAddress.value.indexOf('快速')==-1)
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：違規地點如勾選快速道路，違規地點名稱必須包含『快速』兩字。";
	}
<%end if%>
	if (error==0){
		if (InsertFlag==0){
			InsertFlag=1;
			getChkCarIllegalDate();
		}
	}else{
		alert(errorString);
	}
}

//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function getChkCarIllegalDate(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2="";
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	NewBillUnitID=myForm.BillUnitID.value;
	NewIllegalAddress=myForm.IllegalAddress.value;
	runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress+"&nowTime=<%=now%>");

	//window.open("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function setChkCarIllegalDate(CarCnt,Illdate,RuleDetail)
{
	var ErrorStringChkCarIllegal="";
	if (CarCnt=="1"){
		ChkCarIlldateFlag="1";
	}else{
		ChkCarIlldateFlag="0";
	}
	if (RuleDetail==2){
		alert("舉發單位代號輸入錯誤。");
		InsertFlag=0;
<%if sys_City="高雄市" then%>
	}else if (RuleDetail==3 || RuleDetail==4){
		alert("此車號為業管車輛。");
		InsertFlag=0;
<%end if%>
<%if sys_City="南投縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%elseif sys_City="宜蘭縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。");
		InsertFlag=0;
<%end if%>
<%if sys_City="台中市" then%>
	//}else if (RuleDetail==6){
	//	alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
	//	InsertFlag=0;
<%elseif sys_City<>"台東縣" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間、違規地點，有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%end if%>
	}else{
		if (RuleDetail==1 || RuleDetail==3){
			ErrorStringChkCarIllegal='違規事實與簡式車種不符，請確認是否正確。\n';
		}
		if (ChkCarIlldateFlag=="1"){
		<%if sys_City="宜蘭縣" Or sys_City="基隆市" Or sys_City="台南市" then%>
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號於'+Illdate+'，有違規舉發記錄，請確認有無連續開單。\n';
		<%else%>
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號於'+Illdate+'，有相同違規舉發，請確認有無連續開單。\n';
		<%end if%>
		}
		<%if sys_City="高雄市" then%>
		if ((myForm.IllegalAddressID.value=='00212' || myForm.IllegalAddressID.value=='00213') && myForm.chkHighRoad.checked==false){
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此違規地點為快速道路，請確認是否勾選快速道路。\n';
		}
		<%end if%>
		<%if sys_City="台中市" then%>
			if (!ChkIllegalDateTC(myForm.IllegalDate.value) && myForm.JurgeDay.value==""){
				ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+"違規日期已超過30天。\n";
			}
		<%end if%>	
		<%if RepeatBill=1 then%>
			<%if sys_City="台中市" or sys_City="屏東縣" then%>
				ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號在當日內有其他違規案件。\n';
			<%else%>
				ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號在七日內有其他違規案件。\n';
			<%end if%>
		<%end if %>
		if (ErrorStringChkCarIllegal != ""){
			if(confirm(ErrorStringChkCarIllegal + '\n是否確定要存檔？')){
				myForm.kinds.value="DB_insert";
				myForm.submit();
			}else{
				InsertFlag=0;
			}
		}else{
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}
}
//是否為特殊用車
function getVIPCar(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");

}
//檢查違規日期是否超過30天(台中市)
function ChkIllegalDateTC(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-29,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//檢查違規日期是否超過89天(台中市)
function ChkIllegalDateTC89(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-89,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//檢查輔助車種
function getAddID(){
	//myForm.CarAddID.value=myForm.CarAddID.value.replace(/[^\d]/g,'');
	Layer110.style.visibility='hidden';
	if (myForm.CarAddID.value.length>0){
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7" && myForm.CarAddID.value != "8" && myForm.CarAddID.value != "9" && myForm.CarAddID.value != "10" && myForm.CarAddID.value != "11"){
			alert("輔助車種填寫錯誤!");
			//myForm.CarAddID.value = "";
			myForm.CarAddID.focus();
		}
	}
}
//檢查簡式車種
function getRuleAll(){
	//myForm.CarSimpleID.value=myForm.CarSimpleID.value.replace(/[^\d]/g,'');
	//Layer012.style.visibility='hidden';
	if (myForm.CarSimpleID.value.length>0){
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "5" && myForm.CarSimpleID.value != "6" && myForm.CarSimpleID.value != "7"){
			alert("簡式車種填寫錯誤!");
			myForm.CarSimpleID.focus();
			myForm.CarSimpleID.value = "";
		}
	}
}
//違規事實1(ajax)
function getRuleData1(flag){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail_forLawPlus.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
	<%if not rs1.eof then%>
		<%'if trim(rs1("ProsecutionTypeID"))<>"R" then%>
		CallChkLaw1();
		<%'end if%>
	<%end if%>
		if (event){
			if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
				if (myForm.Rule1.value.length=="7"){
					if ((myForm.Rule1.value.substr(0,2))!="29" && ((myForm.Rule1.value.substr(0,5))!="33101" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,5))!="43102")){
						myForm.Rule2.select();
						myForm.IllegalSpeed.value="";
						myForm.RuleSpeed.value="";
					}else{
						if (flag!="NoSelect"){
						<%if sys_City="屏東縣" then%>
							if (myForm.RuleSpeed.value==""){
								myForm.RuleSpeed.select();
							}else{
								myForm.IllegalSpeed.select();
							}
						<%else%>
							myForm.IllegalSpeed.select();
						<%end if %>
						}
					}
				}
			}
		}

	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
	//AutoGetRuleID(1);
}
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
	<%if not rs1.eof then%>
		CallChkLaw2();
	<%end if%>
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
			if (myForm.Rule2.value.length=="7"){
			<%if sys_City="苗栗縣" then%>
				myForm.IllegalAddressID.select();
			<%else%>
				myForm.BillMem1.select();
			<%end if %>
			}
		}
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
	//AutoGetRuleID(1);
}
//function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
//	Rule1tmp=myForm.Rule1.value;
//		if ((Rule1tmp.substr(0,2))!="33" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,2))!="43" && (Rule1tmp.substr(0,2))!="29"){
			//myForm.BillMem1.focus();
//		}
//}
//到案處所(ajax)
function getStation(){
	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation.asp?StationID="+StationNum+"&nowTime=<%=now%>");
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
		response.write "116"
else
		response.write "117"
end if
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_Unit.asp?SType=U","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum+"&nowTime=<%=now%>");
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}

function UserInputBillType(){

}
//逕舉不一定要輸入固定桿編號. 除了是下方選擇使用固定桿
function getFixID(){
	if (myForm.UseTool.value.length == "1"){
		if (myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3"){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.focus();
			myForm.UseTool.value = "";
		}else if (myForm.UseTool.value == "1"){
			//Layer11.style.visibility = "visible"; 
		}else{
			//Layer11.style.visibility = "hidden"; 
		}
	}
}
//違規地點代碼(ajax)
function getillStreet(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91) || event.keyCode==<%
	if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
	else
		response.write "117"
	end if 
		%>){
		myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
		if (event.keyCode==<%
	if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
	else
		response.write "117"
	end if 
		%>){	
			event.keyCode=0;
			event.returnValue=false;
			OstreetID=myForm.IllegalAddressID.value;
			window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
		if (myForm.IllegalAddressID.value.length > 2){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum+"&nowTime=<%=now%>");
		}
	
		if (myForm.IllegalAddressID.value.length == 6){
		<%if sys_City="苗栗縣" then %>
			myForm.IllegalAddress.select();
		<%else%>
			myForm.Rule1.select();
		<%end if%>
		}
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event){
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
			myForm.BillMem1.value=myForm.BillMem1.value.toUpperCase();
		}
		if (event.keyCode==<%
	if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
			response.write "116"
	else
			response.write "117"
	end if
		%>){	
			event.keyCode=0;
			event.returnValue=false;
			window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
	}
	if (myForm.BillMem1.value.length > 2){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum+"&nowTime=<%=now%>");
	}else if (myForm.BillMem1.value.length <= 2 && myForm.BillMem1.value.length > 0){
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=1;
	}else{
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=0;
	}
}
//舉發人2(ajax)
function getBillMemID2(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem2.value=myForm.BillMem2.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 2){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum+"&nowTime=<%=now%>");
	}else if (myForm.BillMem2.value.length <= 2 && myForm.BillMem2.value.length > 0){
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=1;
	}else{
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=0;
	}
}
//舉發人3(ajax)
function getBillMemID3(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem3.value=myForm.BillMem3.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 2){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum+"&nowTime=<%=now%>");
	}else if (myForm.BillMem3.value.length <= 2 && myForm.BillMem3.value.length > 0){
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=1;
	}else{
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=0;
	}
}
//舉發人4(ajax)
function getBillMemID4(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem4.value=myForm.BillMem4.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 2){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum+"&nowTime=<%=now%>");
	}else if (myForm.BillMem4.value.length <= 2 && myForm.BillMem4.value.length > 0){
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=1;
	}else{
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=0;
	}
}
function getBillFillDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}

//	if (myForm.IllegalDate.value.length >= 6 ){
//		myForm.BillFillDate.value=myForm.IllegalDate.value;
//		getDealLineDate();
//	}
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.BillFillDate.value;
	if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday);
		var DLineDate=new Date()
		DLineDate=DateAdd("d",getDealDateValue,BFillDate);
		Dyear=parseInt(DLineDate.getFullYear())-1911;
		Dmonth=parseInt(DLineDate.getMonth())+1;
		Dday=DLineDate.getDate();
		Dyear=Dyear.toString();
		if (Dmonth < 10){
			Dmonth="0"+Dmonth;
		}
		if (Dday < 10){
			Dday="0"+Dday;
		}
		myForm.DealLineDate.value=Dyear+Dmonth+Dday;
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
//用固定桿編號抓出違規地點
function setFixEquip(){
	if (myForm.FixID.value.length > 2){
		var FixNum=myForm.FixID.value;
		runServerScript("getFixIDAddress.asp?FixNum="+FixNum+"&nowTime=<%=now%>");
	}
}
function RuleSpeedforLaw(){
	//myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	else
		response.write "41"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "40"
	else
		response.write "40"
	end if
	%>公里以上。";
				IntError=IntError+1;
				<%If Trim(request("Speed"))="1" Then%>
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340069(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340045)";
				<%else%>
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
				<%end if%>
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
<%if sys_City="高雄市" then%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 100 && (myForm.IllegalSpeed.value - myForm.RuleSpeed.value) < 150)
		{
			SpeedError=1;
			window.open("../BillKeyIn/ChkSpeedPW.asp","ChkSpeedPW","left=300,top=100,width=350,height=200,resizable=yes,scrollbars=no");
		}else{
			SpeedError=0;
		}
	}
<%end if%>
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function funGetSpeedRule(){
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
	<%if not rs1.eof then%>
		<%'if trim(rs1("ProsecutionTypeID"))<>"R" then%>
		CallChkLaw1();
		<%'end if%>
		CallChkLaw2();
	<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	else
		response.write "41"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "40"
	else
		response.write "40"
	end if
	%>公里以上。";
				IntError=IntError+1;
				<%If Trim(request("Speed"))="1" Then%>
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340069(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340045)";
				<%else%>
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
				<%end if%>
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
<%if sys_City="高雄市" then%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 100 && (myForm.IllegalSpeed.value - myForm.RuleSpeed.value) < 150)
		{
			SpeedError=1;
			window.open("../BillKeyIn/ChkSpeedPW.asp","ChkSpeedPW","left=300,top=100,width=350,height=200,resizable=yes,scrollbars=no");
		}else{
			SpeedError=0;
		}
	}
<%end if%>
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule("NoSelect");
	<%end if%>
}
function CallChkLaw1(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule1.value)){
			alert("請確認法條一是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule1.value) && myForm.Rule2.value==""){
		alert("請確認法條一是否填寫正確");
	}
}
function CallChkLaw2(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule2.value)){
			alert("請確認法條二是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value==""){
		alert("請確認法條二是否填寫正確");
	}
}
//法律條文建檔檢查
function funcChkLaw(thisLaw){
	if (thisLaw.length>=2){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			//當有打速限及車速時 法條一定落在33XXXX,40XXXX,43XXXX
			if ((thisLaw.substr(0,2))!="33" && (thisLaw.substr(0,2))!="40" && (thisLaw.substr(0,2))!="43" && (thisLaw.substr(0,2))!="29"){
				return false;
			}else{
				//違規地點含有"快速道路"判斷法條是否選33XXX而非選40XXX
				if ((myForm.IllegalAddress.value.indexOf("快速道路",0)) != -1){
					if ((thisLaw.substr(0,2))=="40"){
						return false;
					}else{
						return true;
					}
				}else{
					return true;
				}
			}
		}else{
			return true;
		}
	}else{
		return true;
	}
}


//審核無效
function funVerifyResult(){
//	if(confirm('確定要將此筆檢舉案件設為無效？')){
//		myForm.kinds.value="VerifyResultNull";
//		myForm.submit();
//	}
	<%
	if trim(request("Speed"))="1" then
		if sys_City="雲林縣" then
%>
	UrlStr="../ReportCase/ReportCase_Verify_YL.asp?CheckType=5&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("Sn"))%>&Speed=1";
<%	
		else
%>
	UrlStr="../ReportCase/ReportCase_Verify.asp?CheckType=5&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("Sn"))%>&Speed=1";
<%		
		end if 
	else
%>
	UrlStr="../ReportCase/ReportCase_Verify.asp?CheckType=0&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("Sn"))%>";
<%
	end if
	%>
	newWin(UrlStr,"ReportCase_Verify",800,450,0,0,"yes","yes","yes","no");
}

//刪除
function funDeleteCase(){
	if(confirm('確定要刪除此筆案件？')){
		myForm.kinds.value="DeleteCase";
		myForm.submit();
	}
}

function funDelResult(){
	UrlStr="../ReportCase/ReportCase_Verify.asp?CheckType=3&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("Sn"))%>";
	newWin(UrlStr,"ReportCase_Verify",800,450,0,0,"yes","yes","yes","no");
}

function KeyDown(){ 
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "117"
else
		response.write "116"
end if 
	%>){	//F5查詢
		event.keyCode=0;   
		event.returnValue=false;   
<%if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then %>
	}else if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%end if %>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
<%
	if not rs1.eof then
		if trim(rs1("CheckFlag"))="0" then
			if checkF2Flag=0 then
%>
		InsertBillVase();
<%			end if
		end if 
	end if
%>
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		event.returnValue=false;  
	//}else if (event.keyCode==117){ //F6查詢
	//	event.keyCode=0;   
	//	event.returnValue=false;  
	//	funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		event.returnValue=false;  
		window.close();
	}else if (event.keyCode==120){ //F9審核無效
		event.keyCode=0;   
		event.returnValue=false;  
<%
	if not rs1.eof then
		if trim(rs1("CheckFlag"))="0" then
%>
		funVerifyResult();
<%		end if 
	end if
%>
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		event.returnValue=false;  

	}else if (event.keyCode==122){ //F11略過
		event.keyCode=0;   
		event.returnValue=false;  

	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		event.returnValue=false; 
	<%if UpSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=UpSn%>&Speed=<%=Trim(request("Speed"))%>'
	<%end if %>
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		event.returnValue=false; 
	<%if UpSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=FirstSn%>&Speed=<%=Trim(request("Speed"))%>'
	<%end if %>
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=DownSn%>&Speed=<%=Trim(request("Speed"))%>'
	<%end if %>
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_CH.asp?CheckSn=<%=LastSn%>&Speed=<%=Trim(request("Speed"))%>'
	<%end if %>
	}
}

function AutoGetIllStreet(){	//按F6可以直接顯示相關路段
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
else
		response.write "117"
end if 
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F6可以直接顯示相關法條
	//if (event.keyCode==117){	
//		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
		window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	//}
}
function ProjectF5(){
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
		response.write "116"
else
		response.write "117"
end if
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}
	if (myForm.ProjectID.value.length > 0){
		var BillProjectID=myForm.ProjectID.value;
		runServerScript("getProjectID.asp?BillProjectID="+BillProjectID+"&nowTime=<%=now%>");
<%if sys_City="苗栗縣" then%>
		if (myForm.ProjectID.value=="9"){
			myForm.CarAddID.value="8";
		}
<%end if%>
	}else{
		Layer001.innerHTML="";
		TDProjectIDErrorLog=0;
	}
}
//用地點、車速抓違規法條
function setIllegalRule(flag){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
	<%if not rs1.eof then%>
		if ((myForm.Rule1.value.substr(0,2))!="29"){
			if (myForm.IllegalDate.value>="1120630"){
				IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
			}else{
				IllegalRule=getIllegalRule_Old1120630(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
			}
		<%if trim(request("Speed"))="1" then%>
			if (IllegalRule=="4000005")
			{
				IllegalRule="4000011";
			}else if (IllegalRule=="4000006")
			{
				IllegalRule="4000012";
			}else if (IllegalRule=="4000007")
			{
				IllegalRule="4000013";
			}else if (IllegalRule=="4310241")
			{
				IllegalRule="4310256";
			}else if (IllegalRule=="4310242")
			{
				IllegalRule="4310257";
			}else if (IllegalRule=="4310212")
			{
				IllegalRule="4310227";
			}else if (IllegalRule=="3310134")
			{
				IllegalRule="3310146";
			}else if (IllegalRule=="3310136")
			{
				IllegalRule="3310148";
			}else if (IllegalRule=="4310240")
			{
				IllegalRule="4310255";
			}else if (IllegalRule=="4310210")
			{
				IllegalRule="4310225";
			}else if (IllegalRule=="4310211")
			{
				IllegalRule="4310226";
//			}else if (IllegalRule=="4310212")
//			{
//				IllegalRule="4310227";
			}
		<%end if %>
			if (IllegalRule!="Null"){
				myForm.Rule1.value=IllegalRule;
				getRuleData1(flag);
			}
		}
		if ((myForm.Rule2.value.substr(0,2))!="29" && ((myForm.Rule1.value.substr(0,5))!="33101" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,5))!="43102")){
			IllegalRule2=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
			if (IllegalRule2!="Null"){
				myForm.Rule2.value=IllegalRule2;
				getRuleData2();
			}
		}
	<%end if%>
	}else{
//		if ((myForm.Rule1.value.substr(0,2))!="29" && ProsecutionTypeID=="R"){
//			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,"0","0",ProsecutionTypeID,myForm.chkHighRoad.checked);
//			if (IllegalRule!="Null"){
//				myForm.Rule1.value=IllegalRule;
//				getRuleData1();
//			}
//		}
	}
}
//附加說明
function Add_LawPlus(){
	if (myForm.Rule1.value==""){
		alert("請先輸入違規法條一!!");
	}else{
	RuleID=myForm.Rule1.value;
	window.open("Query_LawPlus.asp?RuleID="+RuleID+"&theRuleVer=<%=theRuleVer%>","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
	}
}
function changeStreet(){
	//if (myForm.getStreetName.value!=""){
		myForm.kinds.value="getStreet";
		myForm.submit();
	//}
}
<%'if sys_City="高雄市" then%>
var sys_City="<%=sys_City%>";
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes,status=yes");

}
function getIllZip(){
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value+"&nowTime=<%=now%>");
}
<%'end if %>
function funcUpdSaveLocation(){
		myForm.kinds.value="";
		myForm.submit();
}
function NewWindow(Width, Height, URL, WinName){
	var nWidth = Width;
	var nHeight = Height;
	var sURL = URL;
	var nTop = 0;
	var nLeft = 0;
	var sWinSize = "left=" + nLeft + ",top=" + nTop + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	var sWinName = WinName;
	OldObj = window.open(sURL,sWinName,sWinSize + "," + sWinStatus);
}

	//-----------上下左右-------------
	function funTextControl(obj){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			
			//if (obj==myForm.CarNo && myForm.CarNo.value!=""){
				//myForm.IllegalDate.select();
			//}else{
				CodeEnter(obj.name);
			//}
		}else if (event.keyCode==38){ //上換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}else if (event.keyCode==40){ //下換欄
			event.keyCode=0;
			event.returnValue=false;
			
			//if (obj==myForm.CarNo && myForm.CarNo.value!=""){
			//	myForm.IllegalDate.select();
			//}else{
				CodeMoveRight(obj.name);
			//}
		}else if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
else
		response.write "117"
end if 
	%>){ 
			event.keyCode=0;
			event.returnValue=false;
			if (obj==myForm.Rule1){
				AutoGetRuleID(1);
			}else if (obj==myForm.Rule2){
				AutoGetRuleID(2);
			}
		}else if (event.keyCode==9){ //tab
			event.keyCode=0;
			event.returnValue=false;
			
			if (obj==myForm.CarNo && myForm.CarNo.value!=""){
				myForm.IllegalDate.select();
			}else{
				CodeEnter(obj.name);
			}
		}
	}
	//------------------------------

function IllegalDateKeyUP(){
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
		if (myForm.IllegalDate.value.substr(0,1)=="1"){
			if (myForm.IllegalDate.value.length=="7"){
				myForm.IllegalTime.select();
			}
		}else{
			if (myForm.IllegalDate.value.length=="6"){
				myForm.IllegalTime.select();
			}
		}
	}
}

function IllegalTimeKeyUP(){
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
<%if sys_City="苗栗縣" then%>
		if (myForm.IllegalTime.value.length=="4"){
			myForm.IllegalSpeed.select();
		}
<%else%>
		if (myForm.IllegalTime.value.length=="4"){
			if (myForm.IllegalAddressID.value==""){
				myForm.IllegalAddressID.select();
			}else if (myForm.IllegalAddress.value==""){
				myForm.IllegalAddress.select();
			}else{
				myForm.Rule1.select();
			}
		}
<%end if %>
	}
}
<%If Trim(request("Speed"))="1" then%>
function IllegalDateKeyUP2(){
	myForm.StartIllegalDate.value=myForm.StartIllegalDate.value.replace(/[^\d]/g,'');
	if(eval(TodayDate) < eval(myForm.StartIllegalDate.value)){
		alert("違規日期不得大於今天!!");
		myForm.StartIllegalDate.select();
	}
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
		if (myForm.StartIllegalDate.value.substr(0,1)=="1"){
			if (myForm.StartIllegalDate.value.length=="7"){
				myForm.StartIllegalTime.select();
			}
		}else{
			if (myForm.StartIllegalDate.value.length=="6"){
				myForm.StartIllegalTime.select();
			}
		}
	}
}

function IllegalTimeKeyUP2(Sn){
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
		if (Sn=="1")
		{
			if (myForm.StartIllegalTime.value.length=="6"){
				myForm.IllegalDate.select();
			}
		}else{
<%if sys_City="苗栗縣" then%>
			if (myForm.IllegalTime.value.length=="6"){
				myForm.IllegalSpeed.select();
			}
<%else%>
			if (myForm.IllegalTime.value.length=="6"){
				if (myForm.IllegalAddressID.value==""){
					myForm.IllegalAddressID.select();
				}else if (myForm.IllegalAddress.value==""){
					myForm.IllegalAddress.select();
				}else{
					myForm.Rule1.select();
				}
			}	
<%end if %>
		}
	}
}
<%end if%>


//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	window.open("../Query/ShowIllegalImage.asp?FileName="+FileName,"UploadFile","left=0,top=0,location=0,width=910,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes");
}
<%If sys_City="台中市" or sys_City="彰化縣" or sys_City="金門縣" Then%>
function funPictureList(CaseSn){
	window.open("ReportCaseImageList_TC.asp?CaseSn="+CaseSn,"ReportCaseImageList_TC","left=0,top=0,location=0,width=1000,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes");
}
<%end if%>
//開啟詳細資料
function OpenDetail(FileName, SN){
	//+ URL 傳送時會不見,所以置換,到Server Side 再換回來
	NewWindow(1000, 600, '../ProsecutionImage/ProsecutionImageDetail.asp?FileName=' + FileName.replace(/\+/g, '@2@') + '&SN='+SN, 'MyDetail');
}
//開啟檢視圖
function OpenPic2(FileName){
	NewWindow(1000, 700, FileName, 'MyPic');
}


//=====放大鏡=======================================
var iDivHeight = <%
	'If sys_City=ApconfigureCityName Then
	If sys_City="雲林縣" Then
		response.write "140"
	else
		response.write "110"
	end if 
		
	'Else
	'	response.write "90"
	'End If 
			%>; //放大?示?域?度
var iDivWidth = <%
	'If sys_City=ApconfigureCityName Then
	If sys_City="雲林縣" Then
		response.write "750"
	else
		response.write "230"
	end if 
	'Else
	'	response.write "210"
	'End If 
			%>;//放大?示?域高度
var iMultiple = 4; //放大倍?

//?示放大?，鼠?移?事件和鼠???事件都??用本事件
//??：src代表?略?，sFileName放大?片名?
//原理：依据鼠????略?左上角（0，0）上的位置控制放大?左上角???示?域左上角（0，0）的位置
function show(src, sFileName)
{
//判?鼠?事件?生?是否同?按下了
if ((event.button == 1) && (event.ctrlKey == true)){
  iMultiple -= 1;
  //myForm.CarNo.focus();
}else
  if (event.button == 1){
  iMultiple += 1;
   //myForm.CarNo.focus();
  }
if (iMultiple < 3) iMultiple = 3;

if (iMultiple > 14) iMultiple = 14;
  
var iPosX, iPosY; //放大????示?域左上角的坐?
var iMouseX = event.offsetX; //鼠????略?左上角的?坐?
var iMouseY = event.offsetY; //鼠????略?左上角的?坐?
var iBigImgWidth = src.clientWidth * iMultiple;  //放大??度，是?略?的?度乘以放大倍?
var iBigImgHeight = src.clientHeight * iMultiple; //放大?高度，是?略?的高度乘以放大倍?

if (iBigImgWidth <= iDivWidth)
{
  iPosX = (iDivWidth - iBigImgWidth) / 2;
  
}
else
{
  if ((iMouseX * iMultiple) <= (iDivWidth / 2))
  {
  iPosX = 0;
  }
  else
  {
  if (((src.clientWidth - iMouseX) * iMultiple) <= (iDivWidth / 2))
  {
    iPosX = -(iBigImgWidth - iDivWidth);
  }
  else
  {
    iPosX = -(iMouseX * iMultiple - iDivWidth / 2);
  }
  }
}

if (iBigImgHeight <= iDivHeight)
{
  iPosY = (iDivHeight - iBigImgHeight) / 2;
}
else
{
  if ((iMouseY * iMultiple) <= (iDivHeight / 2))
  {
	iPosY = 0;
  }
  else
  {
	  if (((src.clientHeight - iMouseY) * iMultiple) <= (iDivHeight / 2))
	  {
		iPosY = -(iBigImgHeight - iDivHeight);
	  }
	  else
	  {
		iPosY = -(iMouseY * iMultiple - iDivHeight / 2);
	  }
  }
}
div1.style.height = iDivHeight;
div1.style.width = iDivWidth;

myForm.BigImg.width = iBigImgWidth;
myForm.BigImg.height = iBigImgHeight;
myForm.BigImg.style.top = iPosY;
myForm.BigImg.style.left = iPosX;
}
//============================================================
var ImageFileNameTemp;
function ChangeImgB(){
<%if sPicWebPath2<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgB.src;

	myForm.SmallImgB.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameB.value;
	myForm.ImageFileNameB.value=ImageFileNameTemp;
<%end if%>
}

function ChangeImgC(){
<%if sPicWebPath<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgC.src;

	myForm.SmallImgC.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameC.value;
	myForm.ImageFileNameC.value=ImageFileNameTemp;
<%end if%>
}

function ChangeImgD(){
<%if sPicWebPath3<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgD.src;

	myForm.SmallImgD.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameD.value;
	myForm.ImageFileNameD.value=ImageFileNameTemp;
<%end if%>
}

function setImageNotUse(ImgID){
<%if bPicWebPath<>"" then%>
//	if (ImgID=="A")
//	{
//		if (myForm.chkImgNoUseA.value=="-1")
//		{
//			myForm.chkImgNoUseA.value="1";
//			myForm.btnImgNoUseA.style.backgroundColor='';
//			
//		}else{
//			myForm.chkImgNoUseA.value="-1";
//			myForm.btnImgNoUseA.style.backgroundColor='red';
//		}		
//	}
<%end if %>
<%if sPicWebPath2<>"" then%>
//	if (ImgID=="B")
//	{
//		if (myForm.chkImgNoUseB.value=="-1")
//		{
//			myForm.chkImgNoUseB.value="1";
//			myForm.btnImgNoUseB.style.backgroundColor='';
//			
//		}else{
//			myForm.chkImgNoUseB.value="-1";
//			myForm.btnImgNoUseB.style.backgroundColor='red';
//		}		
//	}
<%end if %>
<%if sPicWebPath<>"" then%>
//	if (ImgID=="C")
//	{
//		if (myForm.chkImgNoUseC.value=="-1")
//		{
//			myForm.chkImgNoUseC.value="1";
//			myForm.btnImgNoUseC.style.backgroundColor='';
//			
//		}else{
//			myForm.chkImgNoUseC.value="-1";
//			myForm.btnImgNoUseC.style.backgroundColor='red';
//		}		
//	}
<%end if %>
<%if sPicWebPath3<>"" then%>
//	if (ImgID=="D")
//	{
//		if (myForm.chkImgNoUseD.value=="-1")
//		{
//			myForm.chkImgNoUseD.value="1";
//			myForm.btnImgNoUseD.style.backgroundColor='';
//			
//		}else{
//			myForm.chkImgNoUseD.value="-1";
//			myForm.btnImgNoUseD.style.backgroundColor='red';
//		}		
//	}
<%end if %>
}
//-------------浮動視窗------------------
var dragswitch=0 ;
var nsx ;

function drag_dropns(name){ 
temp=eval(name) 
temp.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP) 
temp.onmousedown=gons 
temp.onmousemove=dragns 
temp.onmouseup=stopns 
} 

function gons(e){ 
temp.captureEvents(Event.MOUSEMOVE) 
nsx=e.x 
nsy=e.y 
} 
function dragns(e){ 
if (dragswitch==1){ 
temp.moveBy(e.x-nsx,e.y-nsy) 
return false 
} 
} 

function stopns(){ 
temp.releaseEvents(Event.MOUSEMOVE) 
}

var dragapproved=false ;

function drag_dropie(){ 
if (dragapproved==true){ 
myForm.divX.value=tempx+event.clientX-iex
myForm.divY.value=tempy+event.clientY-iey 
document.all.div1.style.pixelLeft=tempx+event.clientX-iex 
document.all.div1.style.pixelTop=tempy+event.clientY-iey 
return false 
} 
} 

function initializedragie(){ 
iex=event.clientX 
iey=event.clientY 
tempx=div1.style.pixelLeft 
tempy=div1.style.pixelTop 
dragapproved=true 
document.onmousemove=drag_dropie 
} 

if (document.all){ 
document.onmouseup=new Function("dragapproved=false") 
} 
//------------------------------------------------
<%
if not rs1.eof then
%>
//myForm.CarNo.select();
getBillFillDate();
getDealLineDate();

//setIllegalRule();
<%
	if trim(rs1("CarSimpleID"))="" or isnull(rs1("CarSimpleID")) or trim(rs1("CarSimpleID"))="0" then
		if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
%>
<%if sys_City<>"高雄市" then%>
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType!=0){
			myForm.CarSimpleID.value=CarType;
		}
<%end if%>
		
<%
		end if
	end if

	If Trim(Session("SpeedKeyIn_BillMem1")) <>"" Then
%>
		getBillMemID1(1);
<%
	end if 
	If (sys_City="雲林縣" or sys_City="屏東縣")  And Trim(request("Speed"))="1" Then 
%>
		setIllegalRule();
<%
	end if 
end if
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</script>

</html>
