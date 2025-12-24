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
	
	chkIsSpeedRuleFlag_TC=0
	chkIsDoubleFlag_TC=0
	chkIsRule5620002Flag_TC=0

	chkIsIllegalTimeNoRuleFlag_TC=0
	If sys_City="台中市" Then
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
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
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
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
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
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
			illegalDate1=DateAdd("h",-2,illegalDateTmp)
			illegalDate2=DateAdd("h",2,illegalDateTmp)
			strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
			strChk="select count(*) as cnt " &_
				" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
				" and Rule1=to_char('"&trim(request("Rule1"))&"')" &_
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

		illegalDateTmpTC=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
		strIllDateTC=" and IllegalDate=TO_DATE('"&year(illegalDateTmpTC)&"/"&month(illegalDateTmpTC)&"/"&day(illegalDateTmpTC)&" "&Hour(illegalDateTmpTC)&":"&minute(illegalDateTmpTC)&":00','YYYY/MM/DD/HH24/MI/SS')"
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
	End If 
	
If chkIsSpeedTooOver=0 And chkIsExistBillNumFlag=0 And chkIsSpeedRuleFlag_TC=0 And chkIsDoubleFlag_TC=0  And chkIsRule5620002Flag_TC=0 And chkIsIllegalTimeNoRuleFlag_TC=0 Then
	BIllSnChk_Flag=0
	strBIllSnChk="select count(*) as cnt from billbase where sn=" & Trim(request("CheckSn"))
	Set rsBIllSnChk=conn.execute(strBIllSnChk)
	If Not rsBIllSnChk.eof Then
		If CDbl(rsBIllSnChk("cnt"))>0 Then
			BIllSnChk_Flag=1
		End If 
	End If 
	rsBIllSnChk.close
	Set rsBIllSnChk=Nothing 

	strSqlA="select * from BillBaseTmp where Sn=" & Trim(request("CheckSn"))
	set rsA=conn.execute(strSqlA)
	If Not rsA.eof then
	ReportSn=trim(rsA("Sn"))
	
	'SN抓最大值
	If BIllSnChk_Flag=1 then
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
	
	theCarSimpleID="null"
	If trim(request("CarSimpleID"))<>"" Then
		theCarSimpleID=trim(request("CarSimpleID"))
	End If 
	theCarAddID="null"
	If trim(request("CarAddID"))<>"" Then
		theCarAddID=trim(request("CarAddID"))
	End If 
	theIllegalDate=""
	if trim(request("IllegalDate"))<>"" then
		theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
	else
		theIllegalDate = "null"
	end if	

	theIllegalSpeed="null"
	If Trim(request("IllegalSpeed"))<>"" Then
		theIllegalSpeed=trim(request("IllegalSpeed"))
	End If 
	theRuleSpeed="null"
	If Trim(request("RuleSpeed"))<>"" Then
		theRuleSpeed=trim(request("RuleSpeed"))
	End If 
	theForFeit1="null"
	If Trim(request("ForFeit1"))<>"" Then
		theForFeit1=trim(request("ForFeit1"))
	End If 
	theForFeit2="null"
	If Trim(request("ForFeit2"))<>"" Then
		theForFeit2=trim(request("ForFeit2"))
	End If 
	theForFeit3="null"
	If Trim(request("ForFeit3"))<>"" Then
		theForFeit3=trim(request("ForFeit3"))
	End If 
	theForFeit4="null"
	If Trim(request("ForFeit4"))<>"" Then
		theForFeit4=trim(request("ForFeit4"))
	End If 
	if trim(request("Insurance"))="" then
		theInsurance=0
	else
		theInsurance=cint(trim(request("Insurance")))
	end if
	if trim(request("UseTool"))="" then
		theUseTool=0
	else
		theUseTool=trim(request("UseTool"))
	end If
	
	theBillFillDate=""
	if trim(request("BillFillDate"))<>"" then
		theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
	else
		theBillFillDate = "null"
	end if
	theDealLineDate=""
	if trim(request("DealLineDate"))<>"" then
		theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
	else
		theDealLineDate="null"
	end If

	theJurgeDay=""
	if trim(request("JurgeDay"))<>"" then
		theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
	else
		theJurgeDay="null"
	end if
	'BillBase
	'If sys_City="高雄市" Then
		ColAdd=",IllegalZip"
		valueAdd=",'"&trim(request("IllegalZip"))&"'"
	'End if	
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
			" values("&sMaxSN&",'2',''" &_
			",'"&UCase(trim(request("CarNo")))&"',"&theCarSimpleID &_						          
			","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
			",'"&theIllegalAddress&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
			","&theRuleSpeed&","&theForFeit1&",'"&trim(request("Rule2"))&"'" &_
			","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
			","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
			",'',null,''" &_
			",'','','"&trim(rsA("MemberStation"))&"'" &_
			",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
			",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
			",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
			",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			","&theBillFillDate&","&theDealLineDate&",'0',0,sysdate,'" & trim(Session("User_ID")) &"'" &_
			",'"&trim(request("Note"))&"','1','"&trim(rsA("RuleVer"))&"'" &_
			",'"&trim(rsA("DriverSex"))&"','"&trim(rsA("ImageFileName"))&"'" &_
			""&valueAdd&"," & theJurgeDay &"" &_
			")"
			'response.write strInsert
			'response.end
			conn.execute strInsert  

		ConnExecute "影像建檔:"&strInsert,371

		strUpdbilltmp="Update billbasetmp set ReportCaseNo='"&Trim(request("ReportCaseNo"))&"',ReportCreditID='"&Trim(request("ReportCreditID"))&"',CloseDate=sysdate where Sn=" & Trim(request("CheckSn"))
		conn.execute strUpdbilltmp  
	End If
	rsA.close
	Set rsA=Nothing 
	'寫入BILLILLEGALIMAGE
	strSqlB="select * from BILLILLEGALIMAGETemp2 where BillSn=" & Trim(request("CheckSn"))
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

		strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"ImageFileNameD,IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(rsB("Billno")))&"','"&fileTemp1&"'" &_
		",'"&fileTemp2&"','"&fileTemp3&"'" &_
		",'"&fileTemp4&"','"&trim(rsB("IISImagePath"))&"')"

		conn.execute strBillImage  

	End If
	rsB.close
	Set rsB=Nothing 

		'台中市要填告發單號
	if sys_City="台中市" Or sys_City="連江縣" Then
		If Trim(request("ReportNo"))<>"" Then
			strReportNo="insert into BillReportNo(BillSN,ReportNo)" &_
				" values("&sMaxSN&",'"&trim(request("ReportNo"))&"')"
			conn.execute strReportNo
		End If 
	End If


	'將舉發BILL SN寫回檢舉資料billbaseTmp
	strUpd1="Update billbaseTmp set BillStatus='9',CheckFlag='1',BillSn="&sMaxSN  &_
		" where Sn=" & ReportSn
	conn.execute strUpd1

	
	'檢舉案件檢查一周內是否有違規
		if trim(request("JurgeDay"))<>"" Then
			IllegalZipName=""

			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
			If sys_City="台中市" Then
				illegalDate1=illegalDateTmp
				illegalDate2=illegalDateTmp
			Else
				illegalDate1=DateAdd("d",-7,illegalDateTmp)
				illegalDate2=DateAdd("d",7,illegalDateTmp)
			End If 			
				
			strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" 0:0:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

			If sys_City="台中市" Then
				If (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule2")),2)="55" And Len(trim(request("Rule2")))=7) Or (left(trim(request("Rule2")),2)="56" And Len(trim(request("Rule2")))=7) Then
					strIllDate=strIllDate & " and (Rule1 like '55%' or Rule1 like '56%' or Rule2 like '55%' or Rule2 like '56%')"
				else
					strIllDateAdd=""
					If trim(request("Rule2"))<>"" Then
						strIllDateAdd=" or Rule1 like '"&left(trim(request("Rule2")),2)&"%' or Rule2 like '"&left(trim(request("Rule2")),2)&"%'"
					End If 
					strIllDate=strIllDate & " and (Rule1 like '"&left(trim(request("Rule1")),2)&"%' or Rule2 like '"&left(trim(request("Rule1")),2)&"%' "&strIllDateAdd&")"
				End if
				
			End If 

			strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate,JurgeDay" &_
					" from Billbase where sn<>"&sMaxSN &_
					" and carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Recordstateid=0 " & strIllDate & " and JurgeDay is not null "
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
	%>
		<script language="JavaScript">
			window.open("JurgeCaseAlert.asp?BillSn=<%=sMaxSN%>&IllegalZipName=<%=IllegalZipName%>","JurgeCaseAlert","left=100,top=20,location=0,width=700,height=555,resizable=yes,scrollbars=yes")
		</script>
	<%		
			End If 
			rsChk.close
			Set rsChk=Nothing 
		End If

	'台中市6個月內同一員警同一違規車號，要跳提示
	If sys_City="台中市" Then
		strDbl="select count(*) as cnt from billbase where BillMemID1='"&trim(request("BillMemID1"))&"' " &_
			" and CarNo='"&UCase(trim(request("CarNo")))&"' " &_
			" and Rule1=to_char('"&trim(request("Rule1"))&"') " &_
			" and Recorddate between to_date('"&Year(DateAdd("m",-6,now))&"/"&Month(DateAdd("m",-6,now))&"/"&Day(DateAdd("m",-6,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
			" and to_date('"&Year(now)&"/"&Month(now)&"/"&Day(now)&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
			" and recordstateid=0"
		Set rsDbl=conn.execute(strDbl)
		If Not rsDbl.eof Then
			If CDbl(rsDbl("cnt"))>1 then
%>
	<script language="JavaScript">
		alert("此舉發員警於六個月內已對同一違規車號舉發<%=CDbl(rsDbl("cnt"))%>次！！");
		alert("此舉發員警於六個月內已對同一違規車號舉發<%=CDbl(rsDbl("cnt"))%>次！！");
		alert("此舉發員警於六個月內已對同一違規車號舉發<%=CDbl(rsDbl("cnt"))%>次！！");
	</script>
<%		
			End If 
		End If 
		rsDbl.close
		Set rsDbl=Nothing 
	End If 
%>
<script language="JavaScript">
<%if trim(request("DownSn_Temp"))<>"" then%>
	location.href="BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=trim(request("DownSn_Temp"))%>";
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
		alert("限速或實速超過300Km，請確認是否正確！！");
		alert("限速或實速超過300Km，請確認是否正確！！");
	</script>
	<%
Elseif chkIsExistBillNumFlag=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
		alert("儲存失敗，此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
		alert("儲存失敗，此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
	</script>
<%	
ElseIf chkIsSpeedRuleFlag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
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
ElseIf chkIsRule5620002Flag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
	</script>
<%
ElseIf chkIsIllegalTimeNoRuleFlag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請自己去舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請自己去舉發單資料維護系統確認！！");
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請自己去舉發單資料維護系統確認！！");
	</script>
<%
End If
	If chkIllegalDateAndCar_KS=1 Then
%>
	<script language="JavaScript">
		alert("<%=chkAlertString%>");
		alert("<%=chkAlertString%>");
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

	ConnExecute "影像建檔無效案件:"&strUpd,372
%>
<script language="JavaScript">
	
	alert("儲存完成!");
	opener.myForm.submit();
	window.close();
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

strSql1="select * from BillBaseTmp where Sn=" & Trim(request("CheckSn"))
'response.write strSql1
set rs1=conn.execute(strSql1)

'已建檔案件要抓建檔資料
sysCarNo=""
sysOldCarNo=""
sysCarSimpleID=""
sysIllegalDate=""
sysIllegalTime=""
sysReportNo=""
OldReportNo=""
sysIllegalAddressID=""
sysIllegalAddress=""
sysIllegalZip=""
sysOldCarSimpleID=""
sysRule1=""
sysOldRule1=""
sysRuleSpeed=""
sysIllegalSpeed=""
sysRule4=""
sysRule2=""
sysOldRule2=""
sysForFeit1=""
sysForFeit2=""
sysJurgeDay=""
sysBillMemID1=""
sysBillMem1=""
sysBillUnitID=""
sysBillFillDate=""
sysCarAddID=""
sysProjectID=""
sysNote=""
if trim(rs1("BillStatus"))="9" then
	strBill="select * from billbase where sn="&trim(rs1("BillSN"))
	set rsBill=conn.execute(strBill)
	if not rsBill.eof then
		'response.write rsBill("sn")
		sysCarNo=trim(rsBill("CarNo"))
		sysCarSimpleID=trim(rsBill("CarSimpleID"))
		if trim(rsBill("IllegalDate"))<>"" and not isnull(rsBill("IllegalDate")) then 
			sysIllegalDate=gInitDT(rsBill("IllegalDate"))
			sysIllegalTime=Right("00"&hour(rsBill("IllegalDate")),2)&Right("00"&minute(rsBill("IllegalDate")),2)
		end if
		sysIllegalAddressID=Trim(rsBill("IllegalAddressID"))
		sysIllegalAddress=Trim(rsBill("IllegalAddress"))
		if trim(rsBill("IllegalZip"))<>"" and not isnull(rsBill("IllegalZip")) then
			sysIllegalZip=trim(rsBill("IllegalZip"))
		end if
		sysRule1=trim(rsBill("Rule1"))
		if trim(rsBill("RuleSpeed"))<>"" and not isnull(rsBill("RuleSpeed")) then
			sysRuleSpeed=trim(rsBill("RuleSpeed"))
		end If
		if trim(rsBill("IllegalSpeed"))<>"" and not isnull(rsBill("IllegalSpeed")) then
			sysIllegalSpeed=trim(rsBill("IllegalSpeed"))
		end If
		if trim(rsBill("Rule4"))<>"" then
			sysRule4=trim(rsBill("Rule4"))
		end if
		if trim(rsBill("Rule2"))<>"" and not isnull(rsBill("Rule2")) then
			sysRule2=trim(rsBill("Rule2"))
		end If
		if trim(rsBill("ForFeit1"))<>"" and not isnull(rsBill("ForFeit1")) then
			sysForFeit1=trim(rsBill("ForFeit1"))
		end If
		if trim(rsBill("ForFeit2"))<>"" and not isnull(rsBill("ForFeit2")) then
			sysForFeit2=trim(rsBill("ForFeit2"))
		end If
		if trim(rsBill("JurgeDay"))<>"" and not isnull(rsBill("JurgeDay")) then 
			sysJurgeDay=gInitDT(rsBill("JurgeDay"))
		end If
		If Trim(rsBill("BillMemID1"))<>"" Then
			sysBillMemID1=Trim(rsBill("BillMemID1"))
		End If 
		If Trim(rsBill("BillMem1"))<>"" Then
			sysBillMem1=Trim(rsBill("BillMem1"))
		End If 
		if Trim(rsBill("BillUnitID"))<>"" then
			sysBillUnitID=Trim(rsBill("BillUnitID"))
		end if 
		if trim(rsBill("BillFillDate"))<>"" and not isnull(rsBill("BillFillDate")) then 
			sysBillFillDate=gInitDT(rsBill("BillFillDate"))
		end If
		if trim(rsBill("CarAddID"))<>"" and not isnull(rsBill("CarAddID")) then 
			sysCarAddID=trim(rsBill("CarAddID"))
		end If
		if trim(rsBill("ProjectID"))<>"" and not isnull(rsBill("ProjectID")) then 
			sysProjectID=trim(rsBill("ProjectID"))
		end If
		if trim(rsBill("Note"))<>"" and not isnull(rsBill("Note")) then
			sysNote=trim(rsBill("Note"))
		end if
	end if
	rsBill.close
	set rsBill=nothing 
	strRNo="select * from BillReportNo where billsn="&trim(rs1("SN"))
	Set rsRNO=conn.execute(strRNo)
	If Not rsRNO.eof Then
		sysReportNo=Trim(rsRNO("ReportNo"))
		OldReportNo=Trim(rsRNO("ReportNo"))
	End If
	rsRNO.close
	Set rsRNO=nothing

else
	
	if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
		sysIllegalDate=gInitDT(rs1("IllegalDate"))
	end If
	if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
		sysIllegalTime=Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
	end if
	strRNo="select * from BillReportNoTemp where billsn="&trim(rs1("SN"))
	Set rsRNO=conn.execute(strRNo)
	If Not rsRNO.eof Then
		sysReportNo=Trim(rsRNO("ReportNo"))
		OldReportNo=Trim(rsRNO("ReportNo"))
	End If
	rsRNO.close
	Set rsRNO=nothing
	sysIllegalAddressID=Trim(rs1("IllegalAddressID"))
	if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
		sysIllegalAddress=trim(rs1("IllegalAddress"))
	end If
	if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
		sysIllegalZip=trim(rs1("IllegalZip"))
	end if
	
	if trim(rs1("Rule4"))<>"" then
		sysRule4=trim(rs1("Rule4"))
	end if
	if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then 
		sysJurgeDay=gInitDT(rs1("JurgeDay"))
	end If
	If Trim(rs1("BillMemID1"))<>"" Then
		sysBillMemID1=Trim(rs1("BillMemID1"))
	End If 
	If Trim(rs1("BillMem1"))<>"" Then
		sysBillMem1=Trim(rs1("BillMem1"))
	End If 
	if Trim(rs1("BillUnitID"))<>"" then
		sysBillUnitID=Trim(rs1("BillUnitID"))
	end if 
	sysBillFillDate=ginitdt(date)
	if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then 
		sysCarAddID=trim(rs1("CarAddID"))
	end If
	if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then 
		sysProjectID=trim(rs1("ProjectID"))
	end If
	if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
		sysNote=trim(rs1("Note"))
	end if
end if 
	if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
		sysOldCarNo=trim(rs1("CarNo"))
	end if
	if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
		sysOldCarSimpleID=trim(rs1("CarSimpleID"))
    end if
	if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
		sysOldRule1=trim(rs1("Rule1"))
	end If
	if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
		sysOldRule2=trim(rs1("Rule2"))
	end If
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
		
		strImgFile="select * from BILLILLEGALIMAGETemp2 where billSn="&Trim(rs1("SN"))
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
			<input type="hidden" name="btnImgNoUseA" value="相片無效" onclick="setImageNotUse('A');">
			<input type="hidden" name="chkImgNoUseA" value="1">
			
		<%else%>
			<a href="<%=bPicWebPath%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameATemp,14)
			
			%></a>
		<%End If %>
			<div id="div1" style="position:absolute; overflow:hidden; width:<%
			'If sys_City=ApconfigureCityName Then
				response.write "230"
			'Else
			'	response.write "210"
			'End If 
			%>px; height:<%
			'If sys_City=ApconfigureCityName Then
				response.write "110"
			'Else
			'	response.write "90"
			'End If 
			%>px; left:<%
			if trim(request("divX"))="" then
				response.write "780"
			else
				response.write trim(request("divX"))
			end if
			%>px; top:<%
			if trim(request("divY"))="" Then
				response.write "360"
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
			</div>
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
			<input type="hidden" name="btnImgNoUseB" value="相片無效" onclick="setImageNotUse('B');">
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
				<tr><td>當日內舉發案件</td></tr>
			</table>
			<br>
	<%if not rs1.eof Then
		'RecDate1=DateAdd("d",-7,Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate")))
		'RecDate2=DateAdd("d",7,Year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate")) & "/" & Day(rs1("IllegalDate")))
		RecDate1=Year(sysIllegalDate) & "/" & Month(sysIllegalDate) & "/" & Day(sysIllegalDate)
		RecDate2=Year(sysIllegalDate) & "/" & Month(sysIllegalDate) & "/" & Day(sysIllegalDate)

		SqlRule2Plus=""
		RepeatBill=0
		If Trim(rs1("Rule2"))<>"" Then
			If Left(Trim(rs1("Rule1")),2)<>Left(Trim(rs1("Rule2")),2) Then
				SqlRule2Plus=" or Rule1 like '%"&Left(Trim(rs1("Rule2")),2)&"%' or Rule2 like '%"&Left(Trim(rs1("Rule2")),2)&"%'"
			End If 
		End If 
		strRB="select * from billbase where IllegalDate between to_date('"&RecDate1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
			" and to_date('"&RecDate2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
			" and CarNo='"&sysCarNo&"'" &_
			" and recordstateID=0"
		set rsRB=conn.execute(strRB)
		while Not rsRB.eof
			RepeatBill=1
			response.write "<a href='../Query/BillBaseData_Detail.asp?BillSN="&Trim(rsRB("Sn"))&"&BillType=0' target='_blank' >"&Trim(rsRB("BillNO"))&""
			response.write ginitdt(Trim(rsRB("IllegalDate")))&" "&Right("00"&hour(rsRB("IllegalDate")),2)&Right("00"&minute(rsRB("IllegalDate")),2)&"<br>"
			response.write Trim(rsRB("IllegalAddress"))&"</a><br><br>"
			rsRB.movenext
		wend
		rsRB.close
		'response.write strRB
	End if
	%>
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
			<input type="hidden" name="btnImgNoUseC" value="相片無效" onclick="setImageNotUse('C');">
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
			<input type="hidden" name="btnImgNoUseD" value="相片無效" onclick="setImageNotUse('D');">
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
					<table >
					<tr>
					<td >
				<input type="text" size="9" name="CarNo" onBlur="getVIPCar();" value="<%
				response.write sysCarNo
				%>" style=ime-mode:disabled maxlength="8" class="Text2" onkeydown="funTextControl(this);">
					</td>
					<td  style="vertical-align:text-top;">
				<span class="style6">
			    <div id="Layer7" style="position:absolute; width:70px; height:24px; z-index:0;  border: 1px none #000000; color: #FF0000; font-weight: bold;"><%
					'response.write sysOldCarNo
				%></div>
				</span>

				<input type="hidden" size="9" name="MemCarNo" value="<%
					response.write sysOldCarNo
				%>" >
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC" width="8%"><div align="right"><span class="style3">＊</span>車種&nbsp;</div>
				</td>
				<td colspan="3" >
                    <!-- 簡式車種 -->
					<table >
					<tr>
					<td >
                    <input type="text" maxlength="1" size="2" value="<%
                    	response.write sysCarSimpleID
                    %>" name="CarSimpleID" onBlur="getRuleAll();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					</td>
					<td >
					<div id="Layers7" style="position:absolute; width:70px; height:24px; z-index:0;  border: 1px none #000000; color: #FF0000; font-weight: bold;"><%
						response.write sysOldCarSimpleID
					%></div>
					</td>
					<td  style="vertical-align:text-top;">
                    <div id="Layer012" style="display: inline; width:300px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1汽車/2拖車/3重機/4輕機/5動力機械/6臨時車牌/7試車牌</font></div>
					
					<input type="hidden" maxlength="1" size="2" value="<%
						response.write sysOldCarSimpleID
                    %>" name="MemCarSimpleID" >
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC" width="7%"><div align="right"><span class="style3">＊</span>違規時間</div></td>
				<td width="13%">
							<!-- 違規日期 -->
					<input type="text" size="6" maxlength="7" name="IllegalDate" class='Text1' value="<%
						response.write sysIllegalDate
					%>" onBlur="getBillFillDate()" style=ime-mode:disabled onkeydown="funTextControl(this);"  >&nbsp;
							<!-- 違規時間 -->
					<input type="text" size="3" maxlength="4" name="IllegalTime" class='Text1' value="<%
						response.write sysIllegalTime
					%>" onBlur="this.value=this.value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" >
					<input type="hidden" size="3" maxlength="4" name="MemIllegalTime" class='Text1' value="<%
						response.write sysIllegalTime
					%>" style=ime-mode:disabled onkeydown="funTextControl(this);" >
				</td>
				<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">標示單號</div></td>
				<td >
					<input type="text" size="12" name="ReportNo" onkeydown="funTextControl(this);" value="<%
					response.write sysReportNo
				%>" style=ime-mode:disabled maxlength="11">

					
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>地點&nbsp;</div></td>
				<td colspan="3">
					<input type="text" size="4" value="<%
					response.write sysIllegalAddressID
					%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					
					<input type="text" size="30" value="<%
						response.write sysIllegalAddress
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()" <%if sys_City="南投縣" then response.write "disabled"%>>
					<div id="Layerert45" style="display: inline; width:30px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><span class="style1">快速道路</span></div>
					<%if sys_City="台中市" then %>
						<br>
						<table >
						<tr>
						<td >
						區號
						<input type="text" class="btn5" size="3" value="<%
					response.write sysIllegalZip
						%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(rs1("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
						</td>
						<td >
						<div id="LayerIllZip" style="position:absolute ; width:160px; height:30; z-index:0;  border: 1px none #000000;""><%
					if trim(sysIllegalZip)<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&trim(sysIllegalZip)&"'"
						set rsZip1=conn.execute(strZip1)
						if not rsZip1.eof then
							response.write trim(rsZip1("ZipName"))
						end if
						rsZip1.close
						set rsZip1=nothing
					end if
					%></div>
						</td>
						</tr>
						</table>
					<%end if%>
                </td>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style3">＊</span>法條一</div></td>
				<td colspan="5">
					<table >
					<tr>
					<td >
					<input type="text" maxlength="8" size="7" value="<%
						response.write sysRule1
					%>" name="Rule1" onKeyUp="getRuleData1();" style=ime-mode:disabled onkeydown="funTextControl(this);" >

					<input type="hidden" maxlength="9" size="7" value="<%
						response.write sysOldRule1
					%>" name="MemRule1" >

					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<img src="../Image/BillLawPlusButton_Small.JPG" onclick="Add_LawPlus()" alt="附加說明">
					
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
						response.write sysRuleSpeed
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
						response.write sysIllegalSpeed
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					&nbsp;
					</td>
					<td  style="vertical-align:text-top;">
					<span class="style5">
					<div id="Layer1" style="position:absolute ; width:230px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><%
					if trim(sysRule1)<>"" then
						strR1="select * from Law where itemid='"&trim(sysRule1)&"' and Version=2"
						Set rsR1=conn.execute(strR1)
						If Not rsR1.eof Then
							response.write rsR1("IllegalRule")
						End If 
						rsR1.close
						Set rsR1=Nothing 
					end if 

					if trim(rs1("Rule4"))<>"" then
						response.write "("&sysRule4&")"
					end if 
					%></div></span>
					<input type="hidden" name="ForFeit1" value="<%
						response.write sysForFeit1
					%>">
					<input type="hidden" value="<%
						response.write sysRule4
					%>" name="Rule4">
					<input type="hidden" maxlength="8" size="7" value="<%
						response.write sysOldRule2
					%>" name="MemRule2" >
					</td>
					</tr>
					</table>
					
					<%
					response.write "<font color='red'>"&sysOldRule1&"</font>"
					
					%>
					
				</td>
		    </tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right">法條二</div></td>
				<td colspan="3">
					<table >
					<tr>
					<td >
					<input type="text" maxlength="9" size="7" value="<%
						response.write sysRule2
					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<%
					response.write "<br /><font color='red'>"&sysOldRule2&"</font>"
					%>
					</td>
					<td  style="vertical-align:text-top;">
					<span class="style5">
					<div id="Layer2" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%
				if sysRule2<>"" then
					strR1="select * from Law where itemid='"&trim(sysRule2)&"' and Version=2"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof Then
						response.write rsR1("IllegalRule")
					End If 
					rsR1.close
					Set rsR1=Nothing 
				end if
					%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%
					response.write sysForFeit2
					%>">
					</td>
					</tr>
					</table>
					

				</td>
				
				<%if sys_City="台中市" then %>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">民眾檢舉日期</div></td>
					<td >
						<input type="text" name="JurgeDay" value="<%
						response.write sysJurgeDay
						%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">

						<input type="hidden" size="7" name="BillMem2" value="<%%>" style=ime-mode:disabled onkeydown="funTextControl(this);">
						<input type="hidden" value="<%%>" name="BillMemID2">
						<input type="hidden" value="<%%>" name="BillMemName2">
					</td>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">民眾檢舉案號</div></td>
					<td >
						<input type="text" name="ReportCaseNo" value="<%
						if trim(rs1("ReportCaseNo"))<>"" and not isnull(rs1("ReportCaseNo")) then
							response.write trim(rs1("ReportCaseNo"))
						end if
						%>" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:100px;" onblur="this.value=this.value.toUpperCase()">

						
					</td>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">檢舉人證號</div></td>
					<td >
						<input type="text" name="ReportCreditID" value="<%
						if trim(rs1("ReportCreditID"))<>"" and not isnull(rs1("ReportCreditID")) then
							response.write trim(rs1("ReportCreditID"))
						end if
						%>" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:100px;" onblur="this.value=this.value.toUpperCase()">
					</td>			

				<%else%>
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
						<input type="hidden" value="<%=BillRecordID2%>" name="BillMemID2">
						<input type="hidden" value="<%
			
						%>" name="BillMemName2">
					</td>
				<%End if%>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><span class="style3">＊</span>舉發人&nbsp;</div></td>
		  		<td >
					<input type="text" size="9" name="BillMem1" value="<%
				If Trim(sysBillMemID1)<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(sysBillMemID1)
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
					response.write sysBillMem1
					%></div>
					</span>
					<input type="hidden" value="<%
					response.write sysBillMemID1
					%>" name="BillMemID1">
					<input type="hidden" value="<%
					response.write sysBillMem1
					%>" name="BillMemName1">
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span><span class="style4">舉發單位</span></div></td>
				<td >
					<input type="text" size="4" name="BillUnitID" value="<%=sysBillUnitID%>" onKeyUp="getUnit();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="display: inline; width:200px; height:30px; z-index:0;  border: 1px none #000000; "><%
					if Trim(rs1("BillUnitID"))<>"" then
						strUnitName="select UnitName from UnitInfo where UnitID='"&sysBillUnitID&"'"
						set rsUnitName=conn.execute(strUnitName)
						if not rsUnitName.eof then
							response.write trim(rsUnitName("UnitName"))
						end if
						rsUnitName.close
						set rsUnitName=nothing
					end if
					%></div>
					</span>
					
				</td>
				<td bgcolor="#FFFFCC" width="8%">

				<div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>

				<div align="right"><span class="style3">＊</span>填單日期</div></td>
				<td width="9%">
				
				&nbsp;<input type="text" size="6" value="<%=sysBillFillDate%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">

				<input type="hidden" name="SelSN" value="<%=trim(rs1("SN"))%>">

				</td>

				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種&nbsp;</td>
				<td width="6%">
                &nbsp;<input type="text" maxlength="2" size="4" value="<%
					response.write sysCarAddID
				%>" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
                
				</td>

				<td bgcolor="#FFFFCC" width="8%">
		
				<div align="right">專案代碼&nbsp;</div></td>
				<td width="12%">
					&nbsp;<input type="text" size="5" value="<%
					response.write sysProjectID
				%>" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>

					<!-- <div id="Layer5012" style="position:absolute; width:125px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1檢舉達人<br>&nbsp;9拖吊</font></div> -->

					<!-- 採証工具 -->
					<input maxlength="1" size="4" value="1" name="UseTool"  onkeyup="getFixID();" type='hidden' style=ime-mode:disabled> 
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
						response.write sysNote
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


			<input type="button" value="建檔 F2" onclick="InsertBillVase();"  <%
		if rs1.eof then
			response.write "disabled"
		ElseIf Trim(rs1("BillStatus"))<>"5" Then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
			
			<input type="button" name="Submit2932" onClick="funVerifyResult();" value="無效 F9" <%
		if rs1.eof then
			response.write "disabled"
		ElseIf Trim(rs1("BillStatus"))<>"5" Then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
			<img src="/image/space.gif" width="29" height="8">
			<input type="hidden" name="kinds" value="">
			
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=FirstSn%>'" value="<< 第一筆 Home" style="font-size: 9pt; width: 90px; height: 27px" <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=UpSn%>'" value="< 上一筆 PgUp" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<%=ThisSn+1 & " / " & AllSN%>
			<input type="button" name="SubmitNext" onClick="location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=DownSn%>'" value="下一筆 PgDn >" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=LastSn%>'" value="最後一筆 End >>" style="font-size: 9pt; width: 90px; height: 27px" <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>
			&nbsp; &nbsp; 
			<input type="button" name="Submit4232" onClick="funPrintCaseList_Report();" value="建檔清冊" style="font-size: 10pt; width: 100px; height: 27px">
			<img src="/image/space.gif" width="29" height="8">
			<input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" style="font-size: 10pt; width: 70px; height: 27px">
			
			
			
			<img src="/image/space.gif" width="29" height="8">

			<input type="hidden" name="DownSn_Temp" value="<%=DownSn%>">
             <input type="hidden" name="Tmp_Order" value="<%=Session("BillCnt_Image")%>">
				<input type="hidden" name="CheckSn" value="<%=Trim(request("CheckSn"))%>">
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
					response.write "780"
				else
					response.write trim(request("divX"))
				end if
				%>">
				<input type="hidden" name="divY" value="<%
				if trim(request("divY"))="" Then
					response.write "360"
				else
					response.write trim(request("divY"))
				end if
				%>">
				
		</td>
	</tr>
<%If sys_City="宜蘭縣" then%>
	<tr>
	<td colspan="2">
	<a href="逕舉相片建檔.doc" target="_blank"><font  class="styleA2">使用說明下載</font></a>
	</td>
	</tr>
<%End if%>
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
var TDIllZipErrorLog=0;
var TDProjectIDErrorLog=0;
var TDVipCarErrorLog=0;
var SpeedError=0;
var TodayDate=<%=ginitdt(date)%>;

var InsertFlag=0;
<%if sys_City="宜蘭縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="南投縣" Or sys_City="屏東縣" or sys_City="花蓮縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,RuleSpeed,IllegalSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="高雄市" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1,BillMem2||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="苗栗縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalSpeed,RuleSpeed,Rule1,Rule2||IllegalAddressID,IllegalAddress,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,ProjectID,CarAddID");
<%elseif sys_City="台中市" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime,ReportNo||IllegalAddressID,IllegalAddress,IllegalZip,Rule1,RuleSpeed,IllegalSpeed||Rule2,JurgeDay,ReportCaseNo,ReportCreditID||BillMem1,BillUnitID,BillFillDate,CarAddID,ProjectID");
<%else%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%end if%>

//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";

	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");

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
<%if sys_City="台中市" then %>
//	if (((myForm.Rule1.value.substr(0,2))=="35" || (myForm.Rule2.value.substr(0,2))=="35") && (myForm.IsVideo[0].checked==false && myForm.IsVideo[1].checked==false))
//	{
//		error=error+1;
//		errorString=errorString+"\n"+error+"：法條為35條時，請輸入有無全程錄影。";
//	}
<%end if%>
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}
	if (myForm.CarSimpleID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簡式車種。";
	}/*else if(myForm.CarNo.value != "" && chkCarNoFormat(myForm.CarNo.value)!= 0) {
		if (chkCarNoFormat(myForm.CarNo.value) != myForm.CarSimpleID.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：車號格式與簡式車種不符。";
		}
	}*/
	var IllDateFlag=0;
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
		IllDateFlag=1;
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="0"  ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤，請直接輸入年份，開頭不須補0。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="9" && myForm.IllegalDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="1" && myForm.IllegalDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
<%if sys_City="苗栗縣" then%>
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%elseif sys_City="台中市" then%>
	}else if (!ChkIllegalDateTC89(myForm.IllegalDate.value) && myForm.Note.value=="" ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限，請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%else%>
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限，請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%end if%>
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
	//if (myForm.ReportChk.checked==false){
<%if sys_City="基隆市" then %>
		if (myForm.IllegalAddressID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：交通隊規定，徑舉案件 違規地點需使用違規地點代碼設定。\n  請先至代碼維護系統-縣市路段代碼檔中進行設定 ";
		}
<%end if%>
	//}
<%if sys_City="花蓮縣" then %>
	if (myForm.IllegalAddress.value=="" && myForm.CityStreet.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點。";
	}

	if (myForm.CityStreet.value=="")
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點-鄉鎮路段。";
	}else if (TDCityStreetErrorLog==1)
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入完整的違規地點-鄉鎮路段。";
	}
<%else%>
	if (myForm.IllegalAddress.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點。";
	}
<%end if%>
	
<%if sys_City="台南市" then %>
	if (myForm.IllegalZip.value == ""){
		if ((myForm.IllegalAddress.value.indexOf("台86線") == -1) && (myForm.IllegalAddress.value.indexOf("台８６線") == -1) && (myForm.IllegalAddress.value.indexOf("台84線") == -1) && (myForm.IllegalAddress.value.indexOf("台８４線") == -1) && (myForm.IllegalAddress.value.indexOf("台61線") == -1) && (myForm.IllegalAddress.value.indexOf("台６１線") == -1)){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規地點非快速道路,請輸入區號。";
		}
	}
<%end if%>
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
	if (TDLawErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條三輸入錯誤。";
	}
	if (TDLawNum>=1 && myForm.Rule3.value!=""){
		if (myForm.Rule1.value==myForm.Rule3.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條一與違規法條三重複。";
		}
		if (myForm.Rule2.value==myForm.Rule3.value){	
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二與違規法條三重複。";
		}
	}
	if (myForm.BillFillDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入填單日期。";
	}else if (!dateCheck( myForm.BillFillDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="0"  ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤，請直接輸入年份，開頭不須補0。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="9" && myForm.BillFillDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="1" && myForm.BillFillDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
<%if sys_City<>"宜蘭縣" and sys_City<>"嘉義縣" and sys_City<>"嘉義市" then%>
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%else%>
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value) && myForm.ReportChk.checked==true){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%end if%>
<%if sys_City="苗栗縣" then%>
	}else if (!ChkIllegalDateML(myForm.BillFillDate.value) ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過二個月期限。";
	}
<%else%>
	}else if (!ChkIllegalDateML(myForm.BillFillDate.value) ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過二個月期限。";
	}
<%end if%>
<%If sys_City="宜蘭縣" or sys_City="嘉義市" or sys_City="花蓮縣" then%>
	if (myForm.MemberStation.value!="" || myForm.DriverPID.value!=""){
		if (myForm.MemberStation.value=="" || myForm.DriverPID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：轉歸責案件，身分證號與應到案處所都要輸入。";
		}
	}
<%end if %>
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
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入應到案日期。";
	}else if (!dateCheck( myForm.DealLineDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="0"  ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤，請直接輸入年份，開頭不須補0。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
<%if sys_City="苗栗縣" then%>
	}else if (!ChkIllegalDateML(myForm.DealLineDate.value) ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過二個月期限。";
	}
<%else%>
	}else if (!ChkIllegalDateML(myForm.DealLineDate.value) ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過二個月期限。";
	}
<%end if%>
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
			errorString=errorString+"\n"+error+"：請輸入舉發人代碼1。";
		//}
	}else if (TDMemErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 輸入錯誤。";
	}else if (myForm.BillMemID1.value==""){
	    error=error+1;
		errorString=errorString+"\n"+error+"：請重新再輸入一次舉發人代碼1。";
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
	/*
	if (myForm.BillMem1.value==myForm.BillMem2.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼2 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem3.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼3 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem4.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼4 重複。";
	}
	if (myForm.BillMem2.value==myForm.BillMem3.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 與 舉發人代碼3 重複。";
	}else if (myForm.BillMem2.value==myForm.BillMem4.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 與 舉發人代碼4 重複。";
	}
	if (myForm.BillMem3.value==myForm.BillMem4.value && myForm.BillMem3.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼3 與 舉發人代碼4 重複。";
	}
	*/
//	if (TDFastenerErrorLog1==1){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：代保管物1 輸入錯誤。";
//	}
//	if (TDFastenerErrorLog2==1){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：代保管物2 輸入錯誤。";
//	}
//	if (myForm.Fastener1.value==myForm.Fastener2.value && myForm.Fastener1.value!=""){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：代保管物1 與代保管物2 重複。";
//	}
	if (eval(myForm.BillFillDate.value) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
	}
	if(eval(myForm.DealLineDate.value) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期不得比填單日期早。";
	}
	if (myForm.JurgeDay.value!=""){
		if(!dateCheck( myForm.JurgeDay.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：民眾檢舉時間輸入錯誤。";	
		}else if (IllDateFlag==0){
		<%'包含違規日當天
			response.write "var CheckJurgeDay=6;"
		%>
			Iyear=parseInt(myForm.IllegalDate.value.substr(0,myForm.IllegalDate.value.length-4))+1911;
			Imonth=myForm.IllegalDate.value.substr(myForm.IllegalDate.value.length-4,2);
			Iday=myForm.IllegalDate.value.substr(myForm.IllegalDate.value.length-2,2);
			var IllDate=new Date(Iyear,Imonth-1,Iday);

			Jyear=parseInt(myForm.JurgeDay.value.substr(0,myForm.JurgeDay.value.length-4))+1911;
			Jmonth=myForm.JurgeDay.value.substr(myForm.JurgeDay.value.length-4,2);
			Jday=myForm.JurgeDay.value.substr(myForm.JurgeDay.value.length-2,2);
			var JDay=new Date(Jyear,Jmonth-1,Jday);

			var OverDate=new Date();
			OverDate=DateAdd("d",CheckJurgeDay,IllDate);
			if (JDay > OverDate){
				error=error+1;
				errorString=errorString+"\n"+error+"：民眾檢舉時間已超過違規日七天，民眾檢舉發生超過七日之交通違規，依法不得舉發。";	
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
		if ((myForm.Rule1.value)!="3310107" && (myForm.Rule1.value)!="3310108" && (myForm.Rule1.value)!="3310109" && (myForm.Rule1.value)!="3310110"){
			if(parseInt(myForm.RuleSpeed.value) > parseInt(myForm.IllegalSpeed.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：限速、限重大於實際車速、車重。";
			}
		}
		if ((myForm.Rule1.value.substr(0,3))!="293" && (myForm.Rule2.value.substr(0,3))!="293")	{
			if(parseInt(myForm.RuleSpeed.value) < 25){
				error=error+1;
				errorString=errorString+"\n"+error+"：限速、限重小於 25Km/h。";
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
<%if sys_City="台南市" then %>
	if (TDIllZipErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規地點區號 輸入錯誤。";
	}
<%end if%>
<%if sys_City="高雄市" then%>
	if (SpeedError==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：超速 100~150Km/h ，請輸入密碼後才可建檔。";
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
<%if sys_City="台中市" then%>
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
	<%if sys_City="台中市" then%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"2");
	<%elseif sys_City="台東縣" or sys_City="雲林縣" then%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"3");
	<%else%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"1");
	<%end if%>
		if ((myForm.Rule1.value)!="3310107" && (myForm.Rule1.value)!="3310108" && (myForm.Rule1.value)!="3310109" && (myForm.Rule1.value)!="3310110"){
			if (IllegalRule == false){
				error=error+1;
				errorString=errorString+"\n"+error+"：超速法條與車速不符。";
			}
		}
	}else if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
	<%if sys_City="台中市" then%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"2");
	<%elseif sys_City="台東縣" then%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"3");
	<%else%>
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"1");
	<%end if%>
		if ((myForm.Rule2.value)!="3310107" && (myForm.Rule2.value)!="3310108" && (myForm.Rule2.value)!="3310109" && (myForm.Rule2.value)!="3310110"){
			if (IllegalRule == false){
				error=error+1;
				errorString=errorString+"\n"+error+"：超速法條與車速不符。";
			}
		}
	}
	if ((myForm.Rule1.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}else if ((myForm.Rule2.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
<%if sys_City<>"花蓮縣" then%>
	}else if (((myForm.Rule1.value.substr(0,2))=="30" || (myForm.Rule2.value.substr(0,2))=="30") && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
<%end if %>
	}
	if (myForm.Rule1.value=="5610801" || myForm.Rule2.value=="5610801"){
		if (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4"){
			error=error+1;
			errorString=errorString+"\n"+error+"：機車不可開法條5610801。";
		}
	}
<%if sys_City="台中市" then%>
	if (myForm.ReportNo.value!=""){
		if (myForm.ReportNo.value.length<11){
			error=error+1;
			errorString=errorString+"\n"+error+"：標示單號不可少於11碼。";
		}
	}	
//	if (myForm.AcceptBatchNumberChk.checked==true && myForm.AcceptBatchNumber.value==""){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：有勾選批號檢查，但是未輸入批號，請輸入批號或取消勾選。";
//	}
<%end if%>
<%if sys_City="雲林縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then %>
	if (TDVipCarErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：車號 "+myForm.CarNo.value+" 為業管車輛。";
	}
<%end if%>
<%if sys_City="台中市" then %>
	//if (((myForm.Rule1.value.substr(0,2))=="55" || (myForm.Rule2.value.substr(0,2))=="55") && (myForm.ReportChk.checked==false)){
	//	error=error+1;
	//	errorString=errorString+"\n"+error+"：第55條不可逕行舉發。";
	//}
<%end if%>
<%if sys_City="苗栗縣" then%>
	if (myForm.Billno1.value!="")
	{
		if (myForm.Billno1.value.substr(0,1)!="F"){
			error=error+1;
			errorString=errorString+"\n"+error+"：請確認單號開頭碼是否正確。";
		}
	}
<%end if%>
	if ((((myForm.Rule1.value.substr(0,3))=="293" && myForm.Rule1.value.length==8) || ((myForm.Rule2.value.substr(0,3))=="293" && myForm.Rule2.value.length==8)) && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
<%if sys_City="台東縣" then %>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 40){
				if ((myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,5))=="33101"){
					if (myForm.Rule1.value=="4340003" || myForm.Rule2.value=="4340003" || myForm.Rule1.value=="4340044" || myForm.Rule2.value=="4340044" || myForm.Rule1.value=="4340068" || myForm.Rule2.value=="4340068"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340003、4340044、4340068需另單舉發。";
					}
				}
			}
		}
	}
<%else%>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 41){
				if ((myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,5))=="33101"){
					if (myForm.Rule1.value=="4340003" || myForm.Rule2.value=="4340003" || myForm.Rule1.value=="4340044" || myForm.Rule2.value=="4340044" || myForm.Rule1.value=="4340068" || myForm.Rule2.value=="4340068"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340003、4340044、4340068需另單舉發。";
					}
				}
			}
		}
	}
<%end if%>
	ErrorStringChkMemKey="";
	if (myForm.CarNo.value!=myForm.MemCarNo.value){
		ErrorStringChkMemKey=ErrorStringChkMemKey+"\n員警輸入車號為:"+myForm.MemCarNo.value;
	}
//	if (myForm.CarSimpleID.value!=""){
//		if (myForm.CarSimpleID.value!=myForm.MemCarSimpleID.value){
//			ErrorStringChkMemKey=ErrorStringChkMemKey+"\n員警輸入簡式車種為:"+myForm.MemCarSimpleID.value;
//		}
//	}
	if (myForm.Rule1.value!=myForm.MemRule1.value){
		ErrorStringChkMemKey=ErrorStringChkMemKey+"\n員警輸入法條一為:"+myForm.MemRule1.value;
	}
	if (myForm.Rule2.value!=myForm.MemRule2.value){
		ErrorStringChkMemKey=ErrorStringChkMemKey+"\n員警輸入法條二為:"+myForm.MemRule2.value;
	}
//	if (myForm.IllegalTime.value!=myForm.MemIllegalTime.value){
//		ErrorStringChkMemKey=ErrorStringChkMemKey+"\n員警輸入違規時間為:"+myForm.MemIllegalTime.value;
//	}

	if (error==0){
		if (ErrorStringChkMemKey!="")
		{
			if(confirm(ErrorStringChkMemKey + '\n請確認是否要繼續存檔？')){
				if(confirm(ErrorStringChkMemKey + '\n請確認是否要繼續存檔？')){
					if(confirm(ErrorStringChkMemKey + '\n請確認是否要繼續存檔？')){
						if (InsertFlag==0){
							InsertFlag=1;
							getChkCarIllegalDate();
						}
					}	
				}	
			}			
		}else{
			if (InsertFlag==0){
				InsertFlag=1;
				getChkCarIllegalDate();
			}
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
	runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress);

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
		alert("舉發單位代號輸入錯誤。");
		alert("舉發單位代號輸入錯誤。");
		InsertFlag=0;
<%if sys_City="高雄市" then%>
	}else if (RuleDetail==3 || RuleDetail==4){
		alert("此車號為業管車輛。");
		alert("此車號為業管車輛。");
		alert("此車號為業管車輛。");
		InsertFlag=0;
<%end if%>
<%if sys_City="南投縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。");
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。");
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%elseif sys_City="宜蘭縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。");
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。");
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。");
		InsertFlag=0;
<%end if%>
<%if sys_City="台中市" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%elseif sys_City<>"台東縣" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間、違規地點，有相同違規法條，請確認是否正確。");
		alert("此車號在同一違規時間、違規地點，有相同違規法條，請確認是否正確。");
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
		<%if RepeatBill=1 then%>
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號在當日內有其他違規案件。\n';
		<%end if %>
		<%if sys_City="台中市" then%>
			if (!ChkIllegalDateTC(myForm.IllegalDate.value)){
				ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+"違規日期已超過30天。\n";
			}
		<%end if%>	
		if (ErrorStringChkCarIllegal != ""){
			if(confirm(ErrorStringChkCarIllegal + '\n是否確定要存檔？')){
				if(confirm(ErrorStringChkCarIllegal + '\n是否確定要存檔？')){
					if(confirm(ErrorStringChkCarIllegal + '\n是否確定要存檔？')){
						myForm.kinds.value="DB_insert";
						myForm.submit();
					}else{
						InsertFlag=0;
					}
				}else{
					InsertFlag=0;
				}
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

//檢查違規日期是否超過80天(苗栗縣)	109/12/1逕舉改60天
function ChkIllegalDateML(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-60,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//檢查違規日期是否超過89天(台中市) 109/12/1 逕舉改成59天
function ChkIllegalDateTC89(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-59,thisDay);
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
function getRuleData1(AccKey){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail_forLawPlus.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
		CallChkLaw1();
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
			if ((myForm.Rule1.value.length=="7" && (myForm.Rule1.value.substr(0,3))!="293") || (myForm.Rule1.value.length>="8" && (myForm.Rule1.value.substr(0,3))=="293")){
				if ((myForm.Rule1.value.substr(0,2))!="33" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="43" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="33" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,2))!="43" && (myForm.Rule2.value.substr(0,2))!="29"){
					myForm.RuleSpeed.value="";
					myForm.Rule2.select();
				}else{
					if (myForm.IllegalSpeed.value==""){
						myForm.RuleSpeed.select();
					}
				}
			}
		}
<%else%>
		if ((myForm.Rule1.value.substr(0,2))!="33" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="43" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="33" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,2))!="43" && (myForm.Rule2.value.substr(0,2))!="29"){
			myForm.RuleSpeed.value="";
		}
<%end if%>
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		//Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
	if (AccKey!="1"){
		AutoGetRuleID(1);
	}	
}
//違規事實2(ajax)
function getRuleData2(AccKey){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
		CallChkLaw2();
		if ((myForm.Rule1.value.substr(0,2))!="33" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="43" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="33" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,2))!="43" && (myForm.Rule2.value.substr(0,2))!="29"){
			myForm.RuleSpeed.value="";
		}
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
			if ((myForm.Rule2.value.length=="7" && (myForm.Rule2.value.substr(0,3))!="293") || (myForm.Rule2.value.length>="8" && (myForm.Rule2.value.substr(0,3))=="293")){
				myForm.BillMem1.select();
			}
		}
<%end if%>
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
	if (AccKey!="1"){
		AutoGetRuleID(2);
	}
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
		runServerScript("getMemberStation.asp?StationID="+StationNum);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(AccKey){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (AccKey!="1"){
		if (event.keyCode==<%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then
				response.write "117"
			else
				response.write "116"
			end if
			%>){	
			event.keyCode=0;
			window.open("Query_Unit.asp?SType=U","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
	}
	
	if (myForm.BillUnitID.value.length > 0){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		FullToGoNextTag(4,'BillUnitID','BillFillDate');
	<%end if%>
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
<%if sys_City<>"基隆市" and sys_City<>"彰化縣" then%>
	myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
<%end if%>
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		OstreetID=myForm.IllegalAddressID.value;

		window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}else if (myForm.IllegalAddressID.value.length >= 1){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
	
}
//舉發人一(ajax)
function getBillMemID1(AccKey){
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem1.value=myForm.BillMem1.value.toUpperCase();
	}
<%end if%>
	if (AccKey!="1"){
		if (event.keyCode==<%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then
				response.write "117"
			else
				response.write "116"
			end if
			%>){	
			event.keyCode=0;
			window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
	}	
	if (myForm.BillMem1.value.length > 1){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		FullToGoNextTag(6,'BillMem1','BillMem2');
	<%end if%>
	}else if (myForm.BillMem1.value.length <= 1 && myForm.BillMem1.value.length > 0){
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
//舉發人二(ajax)
function getBillMemID2(){
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem2.value=myForm.BillMem2.value.toUpperCase();
	}
<%end if%>
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 1){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		FullToGoNextTag(6,'BillMem2','BillMem3');
	<%end if%>
	}else if (myForm.BillMem2.value.length <= 1 && myForm.BillMem2.value.length > 0){
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
//舉發人三(ajax)
function getBillMemID3(){
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem3.value=myForm.BillMem3.value.toUpperCase();
	}
<%end if%>
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 1){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		FullToGoNextTag(6,'BillMem3','BillMem4');
	<%end if%>
	}else if (myForm.BillMem3.value.length <= 1 && myForm.BillMem3.value.length > 0){
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
//舉發人四(ajax)
function getBillMemID4(){
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem4.value=myForm.BillMem4.value.toUpperCase();
	}
<%end if%>
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 1){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		FullToGoNextTag(6,'BillMem4','BillUnitID');
	<%end if%>
	}else if (myForm.BillMem4.value.length <= 1 && myForm.BillMem4.value.length > 0){
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
		runServerScript("getFixIDAddress.asp?FixNum="+FixNum);
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
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
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
		response.write "41"
	else
		response.write "41"
	end if
	%>公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
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
	}<%if sys_City="台中縣" then%>else if ((myForm.Rule1.value.substr(0,2)!="33" && myForm.Rule2.value.substr(0,2)!="33") && (myForm.chkHighRoad.checked==true)){
		alert("快速道路選項為勾選狀態!!");
	}else if ((myForm.Rule1.value.substr(0,2)=="33" || myForm.Rule2.value.substr(0,2)=="33") && (myForm.chkHighRoad.checked==false)){
		alert("快速道路選項未勾選!!");
	}<%end if%>
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
	UrlStr="../ReportCase/ReportCase_Verify.asp?CheckType=2&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("Sn"))%>";
	newWin(UrlStr,"ReportCase_Verify",800,450,0,0,"yes","yes","yes","no");
}


function KeyDown(){ 

	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		event.returnValue=false;  
		//funcOpenBillQry();

<%if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then %>
	}else if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%end if %>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
<%
	if not rs1.eof then
		if trim(rs1("billstatus"))="5" then
%>
		InsertBillVase();
<%
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
		location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=UpSn%>'
	<%end if %>
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		event.returnValue=false; 
	<%if UpSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=FirstSn%>'
	<%end if %>
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=DownSn%>'
	<%end if %>
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check_TC.asp?CheckSn=<%=LastSn%>'
	<%end if %>
	}
}

function AutoGetIllStreet(){	//按F5可以直接顯示相關路段
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F5可以直接顯示相關法條
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
	window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function ProjectF5(){
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}
	if (myForm.ProjectID.value.length > 0){
		var BillProjectID=myForm.ProjectID.value;
		runServerScript("getProjectID.asp?BillProjectID="+BillProjectID);
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
function setIllegalRule(){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
		if ((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		<%if sys_City="台中市" then%>
			IllegalRule=getIllegalRule2(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%elseif sys_City="台東縣" or sys_City="雲林縣" then%>
			IllegalRule=getIllegalRule3(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%else%>
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%end if%>
			if (IllegalRule!="Null"){
				if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
					myForm.Rule2.value=IllegalRule;
					getRuleData2();
				}else{
					myForm.Rule1.value=IllegalRule;
					getRuleData1();
				}
			}
		}
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

//逕舉建檔清冊
function funPrintCaseList_Report(){
	UrlStr="../Query/PrintCaseDataList_Report.asp?CallType=1";
	newWin(UrlStr,"CaseListWin2342",980,575,0,0,"yes","yes","yes","no");
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
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value);
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

//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	window.open("../Query/ShowIllegalImage.asp?FileName="+FileName,"UploadFile","left=0,top=0,location=0,width=910,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes");
}
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
	If sys_City=ApconfigureCityName Then
		response.write "110"
	Else
		response.write "90"
	End If 
			%>; //放大?示?域?度
var iDivWidth = <%
	If sys_City="高雄市" Then
		response.write "860"
	elseIf sys_City=ApconfigureCityName Then
		response.write "230"
	Else
		response.write "210"
	End If 
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
  myForm.CarNo.focus();
}else
  if (event.button == 1){
  iMultiple += 1;
   myForm.CarNo.focus();
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
	//myForm.BigImg.src=oSmallImg;
	document.getElementById("div1").style.backgroundImage = "url("+oSmallImg+")"; 

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
	//myForm.BigImg.src=oSmallImg;
	document.getElementById("div1").style.backgroundImage = "url("+oSmallImg+")"; 

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
	//myForm.BigImg.src=oSmallImg;
	document.getElementById("div1").style.backgroundImage = "url("+oSmallImg+")"; 

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameD.value;
	myForm.ImageFileNameD.value=ImageFileNameTemp;
<%end if%>
}

function setImageNotUse(ImgID){
<%if bPicWebPath<>"" then%>
	if (ImgID=="A")
	{
		if (myForm.chkImgNoUseA.value=="-1")
		{
			myForm.chkImgNoUseA.value="1";
			myForm.btnImgNoUseA.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseA.value="-1";
			myForm.btnImgNoUseA.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath2<>"" then%>
	if (ImgID=="B")
	{
		if (myForm.chkImgNoUseB.value=="-1")
		{
			myForm.chkImgNoUseB.value="1";
			myForm.btnImgNoUseB.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseB.value="-1";
			myForm.btnImgNoUseB.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath<>"" then%>
	if (ImgID=="C")
	{
		if (myForm.chkImgNoUseC.value=="-1")
		{
			myForm.chkImgNoUseC.value="1";
			myForm.btnImgNoUseC.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseC.value="-1";
			myForm.btnImgNoUseC.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath3<>"" then%>
	if (ImgID=="D")
	{
		if (myForm.chkImgNoUseD.value=="-1")
		{
			myForm.chkImgNoUseD.value="1";
			myForm.btnImgNoUseD.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseD.value="-1";
			myForm.btnImgNoUseD.style.backgroundColor='red';
		}		
	}
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
myForm.divX.value=tempx+event.clientX-iex;
myForm.divY.value=tempy+event.clientY-iey ;
document.getElementById("div1").style.left=(tempx+event.clientX-iex)+"px" ;
document.getElementById("div1").style.top=(tempy+event.clientY-iey)+"px" ;
return false ;
} 
} 

function initializedragie(){ 
iex=event.clientX ;
iey=event.clientY ;
tempx=document.getElementById("div1").offsetLeft ;
tempy=document.getElementById("div1").offsetTop ;
dragapproved=true ;
document.onmousemove=drag_dropie ;
} 

if (document.all){ 
document.onmouseup=new Function("dragapproved=false") 
} 
//------------------------------------------------
<%
if not rs1.eof then
%>
myForm.CarNo.select();
getBillFillDate();
getDealLineDate();
setIllegalRule();
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
end if
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</script>

</html>
