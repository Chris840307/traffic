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
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>逕舉資料建檔作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
AuthorityCheck(223)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
end if
'new代表新增案件 , update 代表資料庫已有該案件
if trim(request("filetype"))="" then
	thefiletype=""
else
	thefiletype=trim(request("filetype"))
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
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)
	'smith remark
	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
end function
'==========================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	
	theImageFileName=""
	sys_Other_Case=0
	if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then
		'檢查是否另案舉發過
		strImgName="select FileName from ProsecutionImageDetail where billsn="&trim(request("ReCoverSn"))
		set rsImgName=conn.execute(strImgName)
		if not rsImgName.eof then
			theImageFileName=trim(rsImgName("FileName"))
			sys_Other_Case=0
		else
			sys_Other_Case=1
	%>
	<script language="JavaScript">
		alert("此單號已做過另案舉發\n或此單號非影像建檔案件！！");
		window.close();
	</script>
	<%
		end if
		rsImgName.close
		set rsImgName=nothing
	end if
'新增告發單
if trim(request("kinds"))="DB_insert" Then
	
	'chkIsExistBillNumFlag=0
	if trim(request("Billno1"))<>"" then
		strchkno="select BillNo from BillBase where BillNo='"&trim(request("Billno1"))&"' and RecordStateID=0"
		set rschkno=conn.execute(strchkno)
		if not rschkno.eof then
			chkIsExistBillNumFlag=1
		else
			chkIsExistBillNumFlag=0
		end if
		rschkno.close
		set rschkno=nothing
	else
		chkIsExistBillNumFlag=0
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
	chkReportNoIsExist=0
	chkIsRule5620002Flag_TC=0
	chkIsSpeedRuleFlag_TC=0
	chkIllegalDateAndCar_KS=0
	chkAlertString=""
	chkIsDoubleFlag_TC=0
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

	chkIsIllegalTimeNoRuleFlag_TC=0
	If sys_City="台中市" Then
		If left(trim(request("Rule1")),2)="40" Or left(trim(request("Rule2")),2)="40" Or left(trim(request("Rule1")),5)="33101" Or left(trim(request("Rule2")),5)="33101" Or left(trim(request("Rule1")),5)="43102" Or left(trim(request("Rule2")),5)="43102" Then
			illegalDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
			illegalDate2=gOutDT(request("IllegalDate"))&" 23:59:59"
			strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
			strChk="select count(*) as cnt " &_
				" from Billbase where (Rule1 like '40%' or Rule2 like '40%' or Rule1 like '33101%' or Rule2 like '33101%' or Rule1 like '43102%' or Rule2 like '43102%') " &_
				" and carno='"&UCase(trim(request("CarNo")))&"'" &_
				" and Recordstateid=0 " & strIllDate & " and IllegalAddress='" & Trim(request("IllegalAddress")) & "'"
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
	
	'違規日期
	theIllegalDate=""
	if trim(request("IllegalDate"))<>"" then
		theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
	else
		theIllegalDate = "null"
	end if	

	chkIsRule56Flag=0
	chkIllegalAddress53Flag=0
	chkIllegalAddressID53Flag=0
	chkReKeyInBill=0
	chkIsDoubleFlag_KL=0
	If sys_City="基隆市" Then
		illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
		illegalDate1=DateAdd("h",-2,illegalDateTmp)
		illegalDate2=DateAdd("h",2,illegalDateTmp)
		strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"

		If (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) then
			strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate" &_
				" from Billbase where ((Rule1 like '56%' and length(Rule1)=7) or (Rule2 like '56%' and length(Rule2)=7)) " &_
				" and carno='"&UCase(trim(request("CarNo")))&"'" &_
				" and Recordstateid=0 " & strIllDate
			'response.write strChk
			Set rsChk=conn.execute(strChk)
			If Not rsChk.eof Then	
				chkIsRule56Flag=1
				chkIsRule56Unit=Trim(rsChk("UnitName"))
				chkIsRule56Rule=Trim(rsChk("Rule1"))
				chkIsRule56IllegalTime=Trim(rsChk("IllegalDate"))
			End If 
			rsChk.close
			Set rsChk=Nothing 
		end if

		strChk="select count(*) as cnt " &_
			" from Billbase where carno='"&UCase(trim(request("CarNo")))&"'" &_
			" and Recordstateid=0 " & strIllDate 
		'response.write strChk
		Set rsChk=conn.execute(strChk)
		If Not rsChk.eof Then	
			If CDbl(rsChk("cnt"))>0 Then
				chkIsDoubleFlag_KL=1
			End If 
		End If 
		rsChk.close
		Set rsChk=Nothing

		If left(trim(request("Rule1")),2)="53" Or left(trim(request("Rule1")),5)="48102" Or left(trim(request("Rule2")),2)="53" Or left(trim(request("Rule2")),5)="48102" Then
			strChk="select count(*) as cnt from Street where StreetID='"&Trim(request("IllegalAddressID"))&"'" &_
				" and Address='"&Trim(request("IllegalAddress"))&"'"

			'response.write strChk
			Set rsChk=conn.execute(strChk)
			If not rsChk.eof Then	
				If CInt(rsChk("cnt"))=0 then
					chkIllegalAddress53Flag=1
				End If 
			End If 
			rsChk.close
			Set rsChk=Nothing 
		End If 
		If left(trim(request("Rule1")),2)="53" Or left(trim(request("Rule2")),2)="53" Then
'			strChk="select RedLightCheck from Street where StreetID='"&Trim(request("IllegalAddressID"))&"'" 
'			Set rsChk=conn.execute(strChk)
'			If not rsChk.eof Then	
'				If trim(rsChk("RedLightCheck"))="1" Then
'				
'				else
'					chkIllegalAddressID53Flag=1
'				End If 
'			End If 
'			rsChk.close
'			Set rsChk=Nothing 
		End If 
		
	End If 
	
	chkReKeyInBill_CYC=0
	If sys_City="嘉義市" Then
		strChk="select count(*) as cnt from Billbase where CarNo='"&Trim(request("CarNo"))&"'" & _
		" and IllegalDate="& theIllegalDate & _
		" and Recordstateid=0 "

		'response.write strChk
		Set rsChk=conn.execute(strChk)
		If not rsChk.eof Then	
			If CInt(rsChk("cnt"))>0 then
				chkReKeyInBill_CYC=1
			End If 
		End If 
		rsChk.close
		Set rsChk=Nothing 
	End If 
	
	strChk="select count(*) as cnt from Billbase where CarNo='"&Trim(request("CarNo"))&"'" & _
	" and IllegalAddress='"&Trim(request("IllegalAddress"))&"'" & _
	" and IllegalDate="& theIllegalDate & _
	" and Rule1=to_char('"&Trim(request("Rule1"))&"') and Recordstateid=0 "

	'response.write strChk
	Set rsChk=conn.execute(strChk)
	If not rsChk.eof Then	
		If CInt(rsChk("cnt"))>0 then
			chkReKeyInBill=1
		End If 
	End If 
	rsChk.close
	Set rsChk=Nothing 

	if chkIsExistBillNumFlag=0 And chkIsSpeedTooOver=0 And chkReportNoIsExist=0 And chkIsRule56Flag=0 and chkIllegalAddress53Flag=0 and chkIllegalAddressID53Flag=0 And chkIsRule5620002Flag_TC=0 And chkIsSpeedRuleFlag_TC=0 And chkReKeyInBill=0 And chkIsDoubleFlag_TC=0 And chkIsIllegalTimeNoRuleFlag_TC=0 then
		

		
		'檢查是否有罰款金額
		if trim(request("ForFeit2"))="" then
			theForFeit2="null"
		else
			theForFeit2=trim(request("ForFeit2"))
		end if
		if trim(request("ForFeit3"))="" then
			theForFeit3="null"
		else
			theForFeit3=trim(request("ForFeit3"))
		end if
		if trim(request("ForFeit4"))="" then
			theForFeit4="null"
		else
			theForFeit4=trim(request("ForFeit4"))
		end if
		'第三責任險處理
		if trim(request("Insurance"))="" then
			theInsurance=0
		else
			theInsurance=cint(trim(request("Insurance")))
		end if
		'採証工具處理
		If Trim(request("ReportChk"))="1" Then
			theUseTool="8"
		else
			if trim(request("UseTool"))="" then
				theUseTool=0
			else
				theUseTool=trim(request("UseTool"))
			end If
		End If 
		'駕駛人生日
		theDriverBirth=""
		if trim(request("DriverBrith"))<>"" then
			theDriverBirth=DateFormatChange(trim(request("DriverBrith")))
		else 
			theDriverBirth = "null"
		end if
		'填單日期
		theBillFillDate=""
		if trim(request("BillFillDate"))<>"" then
			theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
		else
			theBillFillDate = "null"
		end if
		'應到案日期
		theDealLineDate=""
		if trim(request("DealLineDate"))<>"" then
			theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
		else
			theDealLineDate="null"
		end If

		theProjectID=trim(request("ProjectID"))
		'民眾檢舉時間
		theJurgeDay=""
		if trim(request("JurgeDay"))<>"" then
			theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
'			If sys_City="台中市" Then
'				theProjectID="119"
'			End If 
 		else
			theJurgeDay="null"
		end if
		'建檔日期
		'theRecordDate=year(now)&"/"&month(now)&"/"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)

		'時速處理
		if trim(request("IllegalSpeed"))="" then
			theIllegalSpeed="null"
		else
			theIllegalSpeed=trim(request("IllegalSpeed"))
		end if
		'限速處理
		if trim(request("RuleSpeed"))="" then
			theRuleSpeed="null"
		else
			theRuleSpeed=trim(request("RuleSpeed"))
		end if
		'輔助車種處理
		if trim(request("CarAddID"))="" then
			theCarAddID="0"
		else
			theCarAddID=trim(request("CarAddID"))
		end if
		'查流水號
		strSN="select BillBase_seq.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theSN=trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		'高雄市另案舉發
		if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then

			strImgUpd1="Update ProsecutionImageDetail set Billsn="&theSN&" where billsn="&trim(request("ReCoverSn"))
			conn.execute strImgUpd1

			strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
				"IISImagePath) " &_
				"values("&theSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(request("sys_ImageFileNameA"))&"'" &_
				",'"&trim(request("sys_ImageFileNameB"))&"','"&trim(request("sys_ImageFileNameC"))&"','"&trim(request("sys_IISImagePath"))&"')"
			'response.write strBillImage
			conn.execute strBillImage
		end if
		If sys_City="高雄市" Or sys_City="台中市" Then
			ColAdd=",IllegalZip"
			valueAdd=",'"&trim(request("IllegalZip"))&"'"
		End if		
		If sys_City="台南市" Then
			If Trim(request("IllegalZip"))<>"" then
				If Left(trim(request("IllegalAddress")),3)<>"台南市" Then
					theIllegalAddress="台南市"&trim(request("IllegalAddress"))
				Else
					theIllegalAddress=trim(request("IllegalAddress"))
				End If 
			Else
				theIllegalAddress=trim(request("IllegalAddress"))
			End If 
		ElseIf sys_City="花蓮縣x" then
			theIllegalAddress=trim(request("CityStreet"))&trim(request("IllegalAddress"))
		Else
			theIllegalAddress=trim(request("IllegalAddress"))
		End If 
		'BillBase
	if sys_City="苗栗縣" and trim(request("ReCoverSn"))<>"" Then
	'苗栗另案舉發的建檔日要保留
		if trim(request("otherRecordDate"))<>"" then
			theOtherRecordDate=DateFormatChange(trim(request("otherRecordDate")))
		else
			theOtherRecordDate="null"
		end If
		If trim(request("IsMail"))="" Or isnull(request("IsMail")) Then
			IsMail_Temp="1"
		Else
			IsMail_Temp=trim(request("IsMail"))
		End If 
		strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
			",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
			",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
			",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
			",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
			",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
			",BillFillerMemberID,BillFiller" &_
			",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
			",Note,FromNote,FromNoteId,EquipmentID,RuleVer,DriverSex,TrafficAccidentType,ImageFileName"&ColAdd&",JurgeDay)" &_
			" values("&theSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
			",'"&UCase(trim(request("CarNo")))&"',"&trim(request("CarSimpleID")) &_						          
			","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
			",'"&trim(theIllegalAddress)&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
			","&theRuleSpeed&","&trim(request("ForFeit1"))&",'"&trim(request("Rule2"))&"'" &_
			","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
			","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
			",'"&UCase(trim(request("DriverPID")))&"',"& theDriverBirth &",'"&trim(request("DriverName"))&"'" &_
			",'"&trim(request("DriverAddress"))&"','"&trim(request("DriverZip"))&"','"&trim(request("MemberStation"))&"'" &_
			",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
			",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
			",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
			",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			","&theBillFillDate&","&theDealLineDate&",'0',0," & theOtherRecordDate & ",'" & theRecordMemberID &"'" &_
			",'"&trim(request("Note"))&"','"&trim(request("FromNote"))&"','"&trim(request("FromNoteId"))&"','"&IsMail_Temp&"','"&theRuleVer&"'" &_
			",'"&trim(request("DriverSex"))&"','','"&trim(theImageFileName)&"'" &_
			""&valueAdd&","&theJurgeDay&")"
			'response.write strInsert
			conn.execute strInsert
			'theDriverBirth , theBillFillDate
	Else

if UCase(trim(request("Billno1"))) <> "" Then
Dim rs, newbillno, billmemid
Dim getbillsn, dispatchmemberid, getbilldate, getbillmemberid
Dim billstartnumber, billendnumber, counterfoireturn, recordstateid, recorddate
Dim maxGetBillSn, sql, maxSnRS, billmemidRS

newbillno = UCase(trim(request("Billno1")))

billmemid = trim(request("BillMem1"))

' 先判斷 getbilldetail 是否有資料
Set rs = Conn.Execute("SELECT GETBILLSN FROM getbilldetail WHERE BILLNO = '" & newbillno & "'")

If rs.EOF Then
     
    ' 無資料，開始新增流程

    ' 取得 GETBILLSN 最大值 + 1
    Set maxSnRS = Conn.Execute("SELECT NVL(MAX(GETBILLSN),0) AS MAXSN FROM GETBILLBASE")

    maxGetBillSn = CInt(maxSnRS("MAXSN")) + 1
    maxSnRS.close
    Set maxSnRS=Nothing

    getbillsn = maxGetBillSn


    dispatchmemberid = 1

    ' 取得當下時間，格式為 yyyy-mm-dd hh:mm:ss
    getbilldate = Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2) & " " & _
              Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
    recorddate = Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2) & " " & _
              Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)


    ' 取得 GETBILLMEMBERID，根據 billmemid 查詢 memberdata
    Set billmemidRS = Conn.Execute("SELECT MEMBERID FROM memberdata WHERE loginid = '" & billmemid & "'")
    If billmemidRS.EOF Then
        getbillmemberid = Null
    Else
        getbillmemberid = billmemidRS("MEMBERID")
    End If
    billmemidRS.Close


Dim prefix, numPart, numValue, startNum, endNum, iii , chch , isNumericc



' 取出字首(第一個字元)和後8位數字字串
prefix = Left(newbillno, 1) ' E
numPart = trim(Mid(newbillno, 2)) ' 22000151

' 將數字字串轉成整數
numValue = CLng(numPart)

' 計算起始號碼
Dim modValue
modValue = numValue Mod 50

If modValue = 0 Then
    ' 尾數是50的倍數，起始號碼為尾數 - 49
    startNum = numValue - 49
ElseIf modValue <= 50 Then
    ' 尾數不為50倍數，且在前50個號碼內，起始號碼為尾數 - (尾數 mod 50) + 1
    startNum = numValue - modValue + 1
Else
    ' 超過50，起始號碼為尾數
    startNum = numValue
End If

endNum = startNum + 49



' 補零成8位字串
Dim startNumStr, endNumStr
startNumStr = Right("00000000" & CStr(startNum), 8)
endNumStr = Right("00000000" & CStr(endNum), 8)

' 組合完整起訖號碼
billstartnumber = prefix & startNumStr
billendnumber = prefix & endNumStr

	RecordMemberID = 1
    counterfoireturn = 0
    recordstateid = 0
note = Null
isBillIn = Null
	dbi_time = Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2) & " " & Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)

    ' 建立 SQL Insert 語句
'sql = "Insert into GetBillBase (GETBILLSN,DISPATCHMEMBERID,GETBILLDATE,GETBILLMEMBERID,BILLSTARTNUMBER,BILLENDNUMBER,COUNTERFOIRETURN,RecordDate,RecordMemberID,DELETEMEMBERID,note,BillIn,BILLRETURNDATE) values ('" & GETBILLSN & "','" & theRecordMemberID & "','" & GetBillDate &"','" & GetBillMemberID & "','" & BILLSTARTNUMBER & "','" & BILLENDNUMBER & "','" & COUNTERFOIRETURN & "','"& GetBillDate &"' ,'" & RecordMemberID & "','" & Note & "','" & Note & "','" & isBillIn & "','" & Note & "')"  
'sql = "INSERT INTO GETBILLBASE (GETBILLSN, DISPATCHMEMBERID, GETBILLDATE, RECORDSTATEID, RECORDDATE, RECORDMEMBERID) VALUES ( '1', '1', '2025-12-17 13:13:13',  '0', '2025-12-17 14:14:14','1')"
			
sql = "INSERT INTO GETBILLBASE (GETBILLSN, DISPATCHMEMBERID, BILLSTARTNUMBER,BILLENDNUMBER, GETBILLDATE,  GETBILLMEMBERID,  RECORDSTATEID, COUNTERFOIRETURN, RECORDDATE, RECORDMEMBERID) VALUES ( " & GETBILLSN & ", " & theRecordMemberID & ", '" & billstartnumber &"', '" & billendnumber & "',TO_DATE('"&dbi_time&"','YYYY-MM-DD HH24:MI:SS'),  "& getbillmemberid &", 0,0, TO_DATE('"&dbi_time&"','YYYY-MM-DD HH24:MI:SS'),"& theRecordMemberID &" )"
        Conn.Execute(sql)


Dim i, billno
For i = startNum To endNum
    billno = prefix & Right("00000000" & CStr(i), 8)
    sql = "INSERT INTO GETBILLDETAIL (GETBILLSN, BILLNO, BILLSTATEID, RECORDDATE, RECORDMEMBERID, NOTECONTENT) " & _
          "VALUES (" & getbillsn & ", '" & billno & "', 463, NULL, NULL, NULL)"
    Conn.Execute sql
Next
sql = "UPDATE GETBILLDETAIL SET RECORDDATE=TO_DATE('"&dbi_time&"','YYYY-MM-DD HH24:MI:SS'),RECORDMEMBERID="&theRecordMemberID&" WHERE BILLNO='"&newbillno&"'"
Conn.Execute sql

' 起始號碼到 newbillno 前一號碼的範圍
Dim updateStartNum, updateEndNum
updateStartNum = startNum
updateEndNum = numValue - 1

noteContent = 555

For i = updateStartNum To updateEndNum
    billno = prefix & Right("00000000" & CStr(i), 8)
    sql = "UPDATE GETBILLDETAIL SET BILLSTATEID = " & noteContent & " WHERE BILLNO = '" & billno & "'"
    Conn.Execute sql
Next


End If
End If

		
		strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
			",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
			",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
			",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
			",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
			",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
			",BillFillerMemberID,BillFiller" &_
			",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
			",Note,FromNote,FromNoteId,EquipmentID,RuleVer,DriverSex,TrafficAccidentType,ImageFileName"&ColAdd&",JurgeDay,IsVideo)" &_
			" values("&theSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
			",'"&UCase(trim(request("CarNo")))&"',"&trim(request("CarSimpleID")) &_						          
			","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
			",'"&trim(theIllegalAddress)&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
			","&theRuleSpeed&","&trim(request("ForFeit1"))&",'"&trim(request("Rule2"))&"'" &_
			","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
			","&theForFeit4&","&theInsurance&","&theUseTool&",'"&theProjectID&"'" &_
			",'"&UCase(trim(request("DriverPID")))&"',"& theDriverBirth &",'"&trim(request("DriverName"))&"'" &_
			",'"&trim(request("DriverAddress"))&"','"&trim(request("DriverZip"))&"','"&trim(request("MemberStation"))&"'" &_
			",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
			",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
			",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
			",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
			","&theBillFillDate&","&theDealLineDate&",'0',0,SYSDate,'" & theRecordMemberID &"'" &_
			",'"&trim(request("Note"))&"','"&trim(request("FromNote"))&"','"&trim(request("FromNoteId"))&"','"&trim(request("IsMail"))&"','"&theRuleVer&"'" &_
			",'"&trim(request("DriverSex"))&"','','"&trim(theImageFileName)&"'" &_
			""&valueAdd&","&theJurgeDay&",'"&Trim(request("IsVideo"))&"')"
			'response.write strInsert
			conn.execute strInsert
			'theDriverBirth , theBillFillDate   
	End If 

		'舉發單扣件明細檔 BillFastenerDetail
		if trim(request("Fastener1"))<>"" then
			strInsFastene1="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener1"))&"','')"
			conn.execute strInsFastene1
		end if
		if trim(request("Fastener2"))<>"" then
			strInsFastene2="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener2"))&"','')"
			conn.execute strInsFastene2
		end if
		
		'台中市要填告發單號
		if sys_City="台中市" Or sys_City="連江縣" Then
			If Trim(request("ReportNo"))<>"" Then
				strReportNo="insert into BillReportNo(BillSN,ReportNo)" &_
					" values("&theSN&",'"&trim(request("ReportNo"))&"')"
				conn.execute strReportNo
			End If 
		End If
		
		if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then
			strOther="Insert into OtherBill values("&trim(request("ReCoverSn"))&","&theSN&",sysdate,"&theRecordMemberID&",'1')"
			conn.execute strOther
		end If
		
		'將舉發BILL SN寫回檢舉資料billbaseTmp
		If Trim(request("ReportCaseSn"))<>"" then
			strUpd="Update billbaseTmp set Carno='"&UCase(trim(request("CarNo")))&"',BillStatus='8',CloseDate=sysdate,RecordMemberID='" & theRecordMemberID &"',BillSn="&theSN  &_
				" where Sn=" & Trim(request("ReportCaseSn"))
			conn.execute strUpd
		End If 
	%>
	<script language="JavaScript">
		//alert("新增完成");
	<%if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then%>
		alert("另案舉發完成!!");
		<%if trim(request("LinkUr"))="S" then%>
		opener.myForm.submit();
		<%end if%>
		window.close();
	<%end if%>
	</script>
	<%
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
						strIllDateAdd=" or trim(Rule1) like '"&left(trim(request("Rule2")),2)&"%' or Rule2 like '"&left(trim(request("Rule2")),2)&"%'"
					End If 
					strIllDate=strIllDate & " and (trim(Rule1) like '"&left(trim(request("Rule1")),2)&"%' or Rule2 like '"&left(trim(request("Rule1")),2)&"%' "&strIllDateAdd&")"
				End if
			End If 

			strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate,JurgeDay" &_
					" from Billbase where sn<>"&theSN &_
					" and carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Recordstateid=0 " & strIllDate & " and JurgeDay is not null "
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
	%>
		<script language="JavaScript">
			window.open("JurgeCaseAlert.asp?BillSn=<%=theSN%>&IllegalZipName=<%=IllegalZipName%>","JurgeCaseAlert","left=100,top=20,location=0,width=700,height=555,resizable=yes,scrollbars=yes")
		</script>
	<%		
			End If 
			rsChk.close
			Set rsChk=Nothing 
		End If
	


	If sys_City="台南市" Then
		If left(trim(request("Rule1")),3)="431" or left(trim(request("Rule1")),3)="433" or left(trim(request("Rule2")),3)="431" or left(trim(request("Rule2")),3)="433" Then
%>
		<script language="JavaScript">
			alert("提醒：第43條第1項各款與第3項，應加開第4項吊扣牌照之舉發。\n( 本訊息僅為提醒，按確定後可繼續作業 )");

		</script>
<%	
		end if 

		If Trim(request("IllegalZip"))<>"" Then
			strIllZip="select ZipName from Zip where ZipNo='"&Trim(request("IllegalZip"))&"'"
			Set rsIllZip=conn.execute(strIllZip)
			If Not rsIllZip.eof Then
				IllegalZipName=Trim(rsIllZip("ZipName"))
			End If 
			rsIllZip.close
			Set rsIllZip=Nothing 

			strIllDate=" and IllegalAddress like '"&IllegalZipName&"%'"
		End If 
		
		If (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule2")),2)="55" And Len(trim(request("Rule2")))=7) Or (left(trim(request("Rule2")),2)="56" And Len(trim(request("Rule2")))=7) then
			strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate,JurgeDay" &_
				" from Billbase where sn<>"&theSN &_
				" and carno='"&UCase(trim(request("CarNo")))&"'" &_
				" and Recordstateid=0 " & strIllDate & " and (Rule1 like '55%' or Rule1 like '56%' or Rule2 like '55%' or Rule2 like '56%')"
			'response.write strChk
			Set rsChk=conn.execute(strChk)
			If Not rsChk.eof Then	
%>
	<script language="JavaScript">
		alert("此車號已有民眾檢舉舉發紀錄，請確認是否有重複檢舉！！");
		window.open("JurgeCaseAlert.asp?BillSn=<%=theSN%>&IllegalZipName=<%=IllegalZipName%>","JurgeCaseAlert","left=100,top=20,location=0,width=700,height=555,resizable=yes,scrollbars=yes")
	</script>
<%		
			End If 
			rsChk.close
			Set rsChk=Nothing 
		End If 
	End If 
	
	
	ElseIf chkIsExistBillNumFlag=1 then
	%>
	<script language="JavaScript">
		alert("此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
	</script>
	<%
	ElseIf chkIsSpeedTooOver=1 then
	%>
	<script language="JavaScript">
		alert("限速或實速超過300Km，請確認是否正確！！");
	</script>
	<%
	ElseIf chkReportNoIsExist=1 Then
	%>
	<script language="JavaScript">
		alert("此告示單號：<%=UCase(trim(request("ReportNo")))%>，已建檔！！");
	</script>
	<%
	ElseIf chkIsRule56Flag=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，違規時間前後2小時內已有違規停車舉發紀錄 ,舉發紀錄 <%=chkIsRule56Unit%> ,法條：<%=chkIsRule56Rule%> ,違規時間： <%=chkIsRule56IllegalTime%> ！！");
	</script>
<%	
	ElseIf chkIllegalAddress53Flag=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，法條53條、48條1項2款，違規地點只可用違規地點代碼代入，不可自行輸入或修改違規地點。！！");
	</script>
<%	
	ElseIf chkIllegalAddressID53Flag=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，法條53條，交通隊規定違規地點只可用闖紅燈路段代碼，請先至『代碼維護系統-縣市路段代碼檔』設定。！！");
	</script>
<%
	ElseIf chkIsRule5620002Flag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此違規日已有5620002舉發紀錄 ,舉發紀錄 <%=chkIsRule5620002Unit%> ,違規時間： <%=chkIsRule5620002IllegalTime%> ！！");
	</script>
<%	ElseIf chkIsSpeedRuleFlag_TC=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規日、違規地點已有超速舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%
	ElseIf chkReKeyInBill=1 Then
	%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間、違規地點已有相同舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%
	ElseIf chkIsDoubleFlag_TC=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在此違規時間兩小時內已有舉發紀錄 ,請先至舉發單資料維護系統確認！！");
	</script>
<%	
	ElseIf chkIsIllegalTimeNoRuleFlag_TC=1 Then
%>
	<script language="JavaScript">
		alert("儲存失敗，此車號在相同違規時間已有舉發紀錄 ,請自己去舉發單資料維護系統確認！！");
	</script>
<%	
	end If

	If chkIllegalDateAndCar_KS=1 Then
%>
	<script language="JavaScript">
		alert("<%=chkAlertString%>");
	</script>
<%
	End If

	if chkIsDoubleFlag_KL=1 then
%>
	<script language="JavaScript">
		alert("再次提醒，此車號 <%=UCase(trim(request("CarNo")))%>，在兩小時內有其他舉發紀錄。(此訊息僅為提示)");
	</script>
<%
	end if 

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
	</script>
<%		
			End If 
		End If 
		rsDbl.close
		Set rsDbl=Nothing 
		'檢查登記簿
		If Trim(request("AcceptBatchNumber"))<>"" And Trim(request("AcceptBatchNumberChk"))="1" Then 
			If Trim(request("ReportChk"))="1" Then
				strDbl="select * from BillStopCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
					" and BillNo='"&Trim(request("Billno1"))&"' and RecordStateID=0"
				Set rsDbl=conn.execute(strDbl)
				If rsDbl.eof Then
	%>
		<script language="JavaScript">
			alert("此單號，登記簿沒有登打資料！！");
		</script>
	<%		
				End If 
				rsDbl.close
				Set rsDbl=Nothing 
			else
				strBRC="select * from BillRunCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
					" and CarNo='"&trim(request("CarNo"))&"' and RecordStateID=0"
				set rsBRC=conn.execute(strBRC)
				if not rsBRC.eof then
					If Trim(rsBRC("CarSimpleID"))<>Trim(request("CarSimpleID")) Then
	%>
		<script language="JavaScript">
			alert("輸入的車種與登記簿車種不符！！");
		</script>
	<%
					End If 
				Else
	%>
		<script language="JavaScript">
			alert("此車號，登記簿沒有登打記錄！！");
		</script>
	<%
				end if
				rsBRC.close
				set rsBRC=Nothing
			End If 
		End If 
	End If 

	If chkReKeyInBill_CYC=1 Then
	%>
		<script language="JavaScript">
			alert("儲存成功!\n此車號在相同違規時間有其他舉發紀錄，\n請至舉發單資料維護確認是否重複舉發！！");
		</script>
	<%
	End If 

	If Trim(request("ReportCaseSn"))<>"" Then
%>
<script language="JavaScript">
	
	alert("儲存完成!");
	opener.myForm.submit();
	window.close();
</script>
<%
	End If 
end if

Session.Contents.Remove("BillTime_Report")
BillTime_ReportTmp=DateAdd("s" , 1, now)
Session("BillTime_Report")=date&" "&hour(BillTime_ReportTmp)&":"&minute(BillTime_ReportTmp)&":"&second(BillTime_ReportTmp)
'response.write Session("BillTime_Report")

'總共幾筆
if trim(request("Tmp_Order"))="" then
	'response.write Session("BillCnt_Report")&"99"
	Session.Contents.Remove("BillCnt_Report")
	Session.Contents.Remove("BillOrder_Report")
	strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID&" and ImageFileName is null"
	set rsCnt1=conn.execute(strSqlCnt)
		Session("BillCnt_Report")=trim(rsCnt1("cnt"))
		Session("BillOrder_Report")=trim(rsCnt1("cnt"))+1
	rsCnt1.close
	set rsCnt1=nothing
else
	Session("BillCnt_Report")=trim(request("Tmp_Order"))
	Session("BillOrder_Report")=trim(request("Tmp_Order"))+1
end if

bIllegalDate=""
bIllegalTime=""
bIllegalAddressID=""
bIllegalAddress=""
bRule1=""
bForFeit1=""
bRule2=""
bRule4=""
bForFeit2=""
bLoginID1=""
bBillMem1=""
bBillMemID1=""
bLoginID2=""
bBillMem2=""
bBillMemID2=""
bLoginID3=""
bBillMem3=""
bBillMemID3=""
bLoginID4=""
bBillMem4=""
bBillMemID4=""
bBillUnitID=""
bBillType=""
bDealLineDate=""
bMemberStation=""
bBillFillDate=""
bRuleSpeed=""
bEquipMent=""
bCarAddId=""
OtherBillNo=""
bRecordDate=""
bHighSpeedRoad=""
bIsVideo=""
bchkbDealLineDate=""
bReportCaseJurgeDay=""
'高雄市民眾檢舉系統不要抓上一筆資料
If Trim(request("ReportCaseSn"))="" Then
'抓上一筆的資料
	'另案舉發
	if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then
		strSql="select * from billbase where sn="&trim(request("ReCoverSn"))
	else
		strSql="select * from (select * from BillBase" &_
		" where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID &_
		" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
		" and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is null order by RecordDate desc)" &_
		" where rownum=1"
	end if
	set rs1=conn.execute(strSql)
	if not rs1.eof then
		'另案舉發
		if sys_City="苗栗縣" and trim(request("ReCoverSn"))<>"" then
			if trim(rs1("BillNo"))<>"" and not isnull(rs1("BillNo")) then
				OtherBillNo=trim(rs1("BillNo"))
			end if 
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				bRecordDate=ginitdt(trim(rs1("RecordDate")))
			end if
		end if
		if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then
			bBillType="2"
			otherbCarNo=trim(rs1("CarNo"))
			otherbCarSimpleID=trim(rs1("CarSimpleID"))
			otherbIllegalSpeed=trim(rs1("IllegalSpeed"))
			if trim(rs1("CarAddId"))<>"" and not isnull(rs1("CarAddId")) then
				bCarAddId=trim(rs1("CarAddId"))
			end if
		else
			if trim(rs1("BillNo"))<>"" and not isnull(rs1("BillNo")) then
				bBillType="1"
			else
				bBillType="2"
			end if
			otherbCarNo=""
			otherbCarSimpleID=""
			otherbIllegalSpeed=""
			bCarAddId=""
		end if
		if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
			bRuleSpeed=trim(rs1("RuleSpeed"))
		end	if
		if sys_City="高雄市" Or sys_City="台中市" then
			if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
				bIllZip=trim(rs1("IllegalZip"))
			end	if
		end if
		if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			bIllegalDate=ginitdt(trim(rs1("IllegalDate")))
		end if
		if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			if hour(rs1("IllegalDate"))>9 then
				theChangeTime=theChangeTime&hour(rs1("IllegalDate"))
			else
				theChangeTime=theChangeTime&"0"&hour(rs1("IllegalDate"))
			end if
			if minute(rs1("IllegalDate"))>9 then
				theChangeTime=theChangeTime&minute(rs1("IllegalDate"))
			else
				theChangeTime=theChangeTime&"0"&minute(rs1("IllegalDate"))
			end if
			bIllegalTime=theChangeTime
		end if
		if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
			bIllegalAddressID=trim(rs1("IllegalAddressID"))
		end	if
		if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
			bIllegalAddress=trim(rs1("IllegalAddress"))
		end	if
		if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
			bRule1=trim(rs1("Rule1"))
		end	If
		if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
			If Left(trim(rs1("Rule1")),5)="33101" Then
				bHighSpeedRoad="1"
			End If 
		end	If
		
		if trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1")) then
			bForFeit1=trim(rs1("ForFeit1"))
		end	if
		if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
			bRule2=trim(rs1("Rule2"))
		end	if
		if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
			bForFeit2=trim(rs1("ForFeit2"))
		end	if
		if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
			bRule4=trim(rs1("Rule4"))
		end	if
		if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
			strMem1="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID1"))
			set rsMem1=conn.execute(strMem1)
			if not rsMem1.eof then
				bLoginID1=trim(rsMem1("LoginID"))
			end if
			rsMem1.close
			set rsMem1=nothing
		end if
		if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
			bBillMem1=trim(rs1("BillMem1"))
		end if
		if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
			bBillMemID1=trim(rs1("BillMemID1"))
		end if
		if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
			strMem2="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID2"))
			set rsMem2=conn.execute(strMem2)
			if not rsMem2.eof then
				bLoginID2=trim(rsMem2("LoginID"))
			end if
			rsMem2.close
			set rsMem2=nothing
		end if
		if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
			bBillMem2=trim(rs1("BillMem2"))
		end if
		if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
			bBillMemID2=trim(rs1("BillMemID2"))
		end if
		if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
			strMem3="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID3"))
			set rsMem3=conn.execute(strMem3)
			if not rsMem3.eof then
				bLoginID3=trim(rsMem3("LoginID"))
			end if
			rsMem3.close
			set rsMem3=nothing
		end if
		if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
			bBillMem3=trim(rs1("BillMem3"))
		end if
		if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
			bBillMemID3=trim(rs1("BillMemID3"))
		end if
		if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
			strMem4="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID4"))
			set rsMem4=conn.execute(strMem4)
			if not rsMem4.eof then
				bLoginID4=trim(rsMem4("LoginID"))
			end if
			rsMem4.close
			set rsMem4=nothing
		end if
		if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
			bBillMem4=trim(rs1("BillMem4"))
		end if
		if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
			bBillMemID4=trim(rs1("BillMemID4"))
		end if
		if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
			bBillUnitID=trim(rs1("BillUnitID"))
		end if
		if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
			bDealLineDate=ginitdt(trim(rs1("DealLineDate")))
		end if
		if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
			bMemberStation=trim(rs1("MemberStation"))
		end if
		if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
			bBillFillDate=trim(ginitdt(rs1("BillFillDate")))
		end if
		if trim(rs1("UseTool"))<>"" and not isnull(rs1("UseTool")) then
			bUseTool=trim(rs1("UseTool"))
		end if
		if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
			bEquipMent=trim(rs1("EquipMentID"))
		end if
		if trim(rs1("IsVideo"))<>"" and not isnull(rs1("IsVideo")) then
			bIsVideo=trim(rs1("IsVideo"))
		end If
		if (trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate"))) And (trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate"))) then
			bchkbDealLineDate=DateDiff("d",rs1("BillFillDate"),rs1("DealLineDate"))
		end if
		
	end if 
	rs1.close
	set rs1=nothing
End If 

'高雄檢舉系統================================================
If trim(request("kinds"))="" And Trim(request("ReportCaseSn"))<>"" Then
	strSql1="select * from BillBaseTmp where Sn=" & Trim(request("ReportCaseSn"))
	set rs1=conn.execute(strSql1)
	If Not rs1.eof then
		otherbCarNo=Trim(rs1("CarNo"))
		if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			bIllegalDate=ginitdt(trim(rs1("IllegalDate")))
		end if
		if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			if hour(rs1("IllegalDate"))>9 then
				theChangeTime=theChangeTime&hour(rs1("IllegalDate"))
			else
				theChangeTime=theChangeTime&"0"&hour(rs1("IllegalDate"))
			end if
			if minute(rs1("IllegalDate"))>9 then
				theChangeTime=theChangeTime&minute(rs1("IllegalDate"))
			else
				theChangeTime=theChangeTime&"0"&minute(rs1("IllegalDate"))
			end if
			bIllegalTime=theChangeTime
		end If
		if sys_City="高雄市" Or sys_City="台中市" then
			if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
				bIllZip=trim(rs1("IllegalZip"))
			end	if
		end If
		if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
			bIllegalAddressID=trim(rs1("IllegalAddressID"))
		end	if
		if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
			bIllegalAddress=trim(rs1("IllegalAddress"))
		end	If
		if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then
			bReportCaseJurgeDay=ginitdt(trim(rs1("JurgeDay")))
		end If
		if sys_City="雲林縣" then
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				otherbCarSimpleID=trim(rs1("CarSimpleID"))
			end If
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				bRule1=trim(rs1("Rule1"))
			end	If
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				If Left(trim(rs1("Rule1")),5)="33101" Then
					bHighSpeedRoad="1"
				End If 
			end	If

			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				bRule2=trim(rs1("Rule2"))
			end	If
			
			if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
				strMem1="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID1"))
				set rsMem1=conn.execute(strMem1)
				if not rsMem1.eof then
					bLoginID1=trim(rsMem1("LoginID"))
				end if
				rsMem1.close
				set rsMem1=nothing
			end if
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				bBillMem1=trim(rs1("BillMem1"))
			end if
			if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
				bBillMemID1=trim(rs1("BillMemID1"))
			end if	
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				bBillUnitID=trim(rs1("BillUnitID"))
			end if
		End If 
		AlertString=""
		If Trim(rs1("IllegalDate"))<>"" then
			IllegalDateTemp=Year(rs1("IllegalDate"))&"/"&Month(rs1("IllegalDate"))&"/"&Day(rs1("IllegalDate"))
			strChkDbl="select * from BillBaseTmp where Sn<>"&Trim(request("ReportCaseSn")) &_
				" and CarNo='"&Trim(rs1("CarNo"))&"' and IllegalDate between to_date('"&IllegalDateTemp&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
				" and to_date('"&IllegalDateTemp&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and BillStatus<>'7' and RecordStateID=0"
			Set rsChkDbl=conn.execute(strChkDbl)
			While Not rsChkDbl.eof

				AlertString=AlertString&"該車號，於 "&Right("00"&hour(rsChkDbl("IllegalDate")),2)&":"&Right("00"&minute(rsChkDbl("IllegalDate")),2)&" "&Trim(rsChkDbl("IllegalAddress"))&" 已有其他民眾檢舉。違規項目:"&Trim(rsChkDbl("IllegalContent"))&"\n"

				rsChkDbl.MoveNext
			Wend 
			rsChkDbl.close
			Set rsChkDbl=Nothing 
		End if

		strChkDb2="select CloseDate,(select chName from memberdata where memberid=BillBaseTmp.RecordMemberid) as RecordName" &_
		" from BillBaseTmp where Sn="&Trim(request("ReportCaseSn")) &_
		" and BillStatus<>'1' and RecordStateID=0"
			'response.write strChkDb2
		Set rsChkDb2=conn.execute(strChkDb2)
		if Not rsChkDb2.eof Then 

			AlertString=AlertString&"此筆檢舉案件已於" &rsChkDb2("CloseDate")& "，由" & rsChkDb2("RecordName") & "作處理\n"

			
		End If  
		rsChkDb2.close
		Set rsChkDb2=Nothing 


%>
<script language="JavaScript">

<%	If AlertString<>"" Then%>
	alert("<%=AlertString%>");
	opener.myForm.submit();
<%	end if%>
</script>
<%
	'response.write AlertString

	end if 
	rs1.close
	set rs1=nothing
End If 
'=====================================================================
%>

<style type="text/css">
<!--
.style1 {font-size: 14px}
.style3 {font-size: 15px}
.style4 {
	color: #FF0000;
	font-size: 12px
	}
.style4b{
	font-size: 12px
	}
.style5 {
	font-size: 12px
	}
.style6 {font-size: 16px}
.style7 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px
	}
.style8 {
	color: #000000;
	font-size: 12px;
	line-height:14px;
}
.style9 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
	font-weight: bold;
}
.style10 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
}

.style11 {
	color: #FF0000;
	font-size: 18px;
}

.btn2 {font-size: 13px}
#LayerA1 {
	position:absolute;
	width:980px;
	height:182px;
	z-index:1;
	overflow: scroll;
}
.style99 {
	font-size: 15px;
	line-height:18px;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%if sys_City<>"台中縣" then%>
<!-- #include file="../Common/Bannernoimage.asp"-->
<%end if%>
	<form name="myForm" method="post">  

		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="6" height="45"><strong>逕舉資料建檔作業</strong>&nbsp; &nbsp; 日期格式：1150101 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				建檔日期：<%=ginitdt(now)%>
			<%if sys_City="花蓮縣" and trim(Session("Credit_ID"))="A001001" then%>
				<div id="Layer1" style="position:absolute;width:115px;height:26px;z-index:1;" onclick="OpenSpecCar();"></div>
			<%end if%>
				<br>
				<input type="checkbox" name="ReportChk" value="1" onclick="funcReportChk();" <%
				if sys_City="高雄市" and not (trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"") then
					bBillType="1"
				end if 
				if bBillType="1" then
					response.write "checked"
				ElseIf sys_City="保二總隊四大隊二中隊" And bBillType="" Then
					response.write "checked"
				end if
				%>>逕舉手開單&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				<input type="checkbox" name="CaseInByMem" value="1" <%
			if sys_City="嘉義縣" or sys_City="嘉義市" then
				if trim(request("CaseInByMem"))="1" then
					response.write "checked"
				end if
			end if
				%>>逾違規日期超過<%
			if sys_City="基隆市" then
				response.write "30天"
			Else
				response.write "二個月"
			end if
				%>強制建檔&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				<%
			if sys_City="南投縣" then
				CheckFlag=0
				str1="select * from apconfigure where id=777"
				Set rs1=conn.execute(str1)
				If Not rs1.eof Then
					CheckFlag=Trim(rs1("value"))
				End If
				rs1.close
				Set rs1=Nothing 
				If CheckFlag=1 Then
					response.write "<font color='#FF0000'><strong>六分鐘 : 不可建檔</strong></font>"
				Else
					response.write "<font color='#FF0000'><strong>六分鐘 : 可以建檔</strong></font>"
				End If 
			end if
				
				%>
			<%if sys_City="台中市" Then %>
				&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				批號
				<input type="checkbox" name="AcceptBatchNumberChk" value="1" <%
				If Trim(request("AcceptBatchNumberChk"))="1" Then
					response.write "checked"
				ElseIf Trim(request("AcceptBatchFlag"))="" Then
					response.write "checked"
				End If 
				%>>
				<input type="hidden" name="AcceptBatchFlag" value="1">
				<input type="text" size="14" name="AcceptBatchNumber" value="<%
				If Trim(request("AcceptBatchNumber"))<>"" Then
					session("AcceptBatchNumber_Report")=Trim(request("AcceptBatchNumber"))
					response.write Trim(request("AcceptBatchNumber"))
				ElseIf Trim(session("AcceptBatchNumber_Report"))<>"" Then
					response.write Trim(session("AcceptBatchNumber_Report"))
				End If 
				%>" onkeyup="this.value=this.value.toUpperCase()" >
			<%End if%>
				</td>
			</tr>	
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">單號</div></td>
				<td <%
				if sys_City<>"嘉義縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City<>"新竹市" and sys_City<>"雲林縣" and sys_City<>"台中市" and sys_City<>"台南市" then
					'response.write "colspan='3'"
				end if
				%>>
				<table >
				<tr>
				<td>
				<input name="Billno1" type="text" value="<%
				if sys_City="苗栗縣" and trim(request("ReCoverSn"))<>"" then
					response.write OtherBillNo
				else
					response.write theBillno
				end if
				%>" size="10" maxlength="9" onBlur="CheckBillNoExist();" onkeydown="funTextControl(this);" style=ime-mode:disabled <%
				if sys_City="苗栗縣" and trim(request("ReCoverSn"))<>"" then
					
				elseif bBillType<>"1" then
					response.write "disabled"
				end if
				%> <%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then '打滿跳下格%>
				onkeyup="FullToGoNextTag(9,'Billno1','CarNo')"
			<%end if%>>
			<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
				<input type="checkbox" value="1" name="isSave3" <%
				if trim(request("isSave3"))="1" then
					response.write "checked"
				end if
				%>><span class="style8">保留前三碼</span>
			<%end if%>
			<%If sys_City="基隆市" then%>
				&nbsp; &nbsp; &nbsp; 
				<span class="style11">
					請注意：
				</span>
			<%End If %>	
				</td>
				<td style="vertical-align:text-top;">
			<%If sys_City="基隆市" then%>
				
				<div id="Layer110_Street" style="position:absolute; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style10">法條53條、48條1項2款，違規地點只可用違規地點代碼代入，不可自行輸入</span>
				<br>
				<span class="style10">法條53條，違規地點只可用闖紅燈路段代碼，請先至『代碼維護系統-縣市路段代碼檔』設定</span>
				</div>
			<%End If %>	

				</td>
				</tr>
				</table>
				
			
			
				</td>
<%if sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="嘉義市" or sys_City="新竹市" or sys_City="台南市" or sys_City="保二總隊三大隊一中隊" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td>
				<input type="text" size="10" value="<%
				if bBillType<>"1" then
					response.write ginitdt(date)
				else
					if trim(bBillFillDate)="" then
						response.write ginitdt(date)
					else
						response.write bBillFillDate
					end if
				end if
				%>" maxlength="7" name="BillFillDate" onfocus="this.select()" onBlur="getDealLineDate()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
<%elseif sys_City="雲林縣" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td >
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="8" value="<%=otherbCarNo%>">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer7" style="position:absolute; width:140px; height:24px; z-index:0; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
					</td>
					</tr>
					</table>
				</td>
<%elseif sys_City="台中市" then%>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>告示單號</div></td>
				<td >
					<input type="text" size="10" name="ReportNo" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="11" onBlur="getAcceptData();" onkeyup="this.value=this.value.toUpperCase();">
				</td>
<%elseif sys_City="連江縣" then%>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>告示單號</div></td>
				<td >
					<input type="text" size="10" name="ReportNo" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="20" >
				</td>
<%end if%>
<%if sys_City<>"基隆市" And sys_City<>"南投縣" then%>
				<td bgcolor="#FFFFCC"><div align="right" class="style4b">有無全程錄影</div></td>
				<td >
					<input type="radio" name="IsVideo" value="1" <%
					If bIsVideo="1" Then
						'response.write "checked"
					End If 
					%>>有
					<input type="radio" name="IsVideo" value="0" <%
					If bIsVideo="0" Then
						'response.write "checked"
					End If 
					%>>無
					&nbsp; &nbsp; 
					<input type="button" value="清除" style="height: 22px; width: 43px; font-size: 10pt;"
					onclick="IsVideo[0].checked=false;IsVideo[1].checked=false;">
				</td>
<%End if%>
			</tr>
			<tr>
<%if sys_City<>"雲林縣" then%>
			  <td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td width="32%">
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="8" value="<%=otherbCarNo%>">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer7" style="position:absolute; width:140px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
					
					</td>
					</tr>
					</table>
				<%if sys_City="台南市" then%>
					<font style="color: #FF0000;">車號嚴禁輸入中文</font>
				<%End If %>
				</td>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="1" size="4" value="<%=otherbCarSimpleID%>" name="CarSimpleID" onBlur="getRuleAll();" onkeydown="funTextControl(this);" <%
					if sys_City="台中市" Then
						response.write "onkeyup=""chkAcceptBatch();"""
					End If 
					%> onfocus="this.select();" style=ime-mode:disabled>
					<font class="style7">1汽車 / 2拖車/ 3重機/ 4輕機/ 5動力機械/ 6臨時車牌/ 7試車牌</font>
					&nbsp;
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style8">550cc以上重機簡式<br>車種請選擇重機</span>
					</div>
					</td>
					</tr>
					</table>
				</td>
<%else%>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td width="32%">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="1" size="4" value="<%=otherbCarSimpleID%>" name="CarSimpleID" onBlur="getRuleAll();" onkeydown="funTextControl(this);" onfocus="this.select();" style=ime-mode:disabled>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style7">1汽車 / 2拖車/ 3重機<br>/ 4輕機/ 5動力機械/ 6臨時車牌/ 7試車牌</span>
					</div>
					&nbsp;<img src="/image/space.gif" width="120" height="8">
					<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style8">550cc以上重機簡式<br>車種請選擇重機</span>
					</div>
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="2" size="4" value="<%
				if sys_City="新竹市" then
					if trim(request("CarAddID"))="8" then
						response.write trim(request("CarAddID"))
					end if
				end if
					%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer110" style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機 /8拖吊<br>/9(550cc)重機 /10計程車/ 11危險物品<%
				if sys_City="苗栗縣" then
					Response.Write "/ F 檢舉案件"
				end if
				if sys_City="雲林縣" Then
					response.write " /12幼兒車(課輔車)"
				End If 
				%></span>
					</div>
					</td>
					</tr>
					</table>
				</td>
<%end if%>

			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" maxlength="7" name="IllegalDate" onkeydown="funTextControl(this);" onblur="getDealLineDate_Stop()" value="<%=bIllegalDate%>" style=ime-mode:disabled <%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then '打滿跳下格%>
				onkeyup="FullToGoNextTag(6,'IllegalDate','IllegalTime')"
			<%end if%>>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td colspan="3">
				<input type="text" size="4" maxlength="4" name="IllegalTime" value="<%=bIllegalTime%>" onkeydown="funTextControl(this);" style=ime-mode:disabled <%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then '打滿跳下格%>
				onkeyup="FullToGoNextTag(4,'IllegalTime','IllegalAddressID')"
			<%end if%><%
			If sys_City="苗栗縣" Then%>
				 onBlur="getBillData()"
			<%Else%>
				 onBlur="this.value=this.value.replace(/[^\d]/g,'')"
			<%End If 
			%>>
				</td>
			</tr>
<%if sys_City="雲林縣" or sys_City="新竹市" or sys_City="台南市" or sys_City="嘉義市"  then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%=bRuleSpeed%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled >
				</td>
			</tr>
<%end if%>
<%if sys_City<>"嘉義市" then %>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%=bIllegalAddressID%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<input type="hidden" name="OldIllegalAddressID" value="<%=bIllegalAddressID%>">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
			<%If sys_City="基隆市" then%>
					<input type="checkbox" name="LockIllegalAddress" value="1" onclick="changeLockIllegalAddress();" <%
				If Trim(request("CarNo"))<>"" Then
					If Trim(request("LockIllegalAddress"))="1" Then
						flag_LockIllegalAddress=1
						session("sessionLockIllegalAddress")="1"
					Else
						flag_LockIllegalAddress=0
						session("sessionLockIllegalAddress")="0"
					End if
				Else
					If Trim(session("sessionLockIllegalAddress"))="1" Then
						flag_LockIllegalAddress=1
					Else
						flag_LockIllegalAddress=0
					End If 
				End If 
				If flag_LockIllegalAddress=1 Then
					response.write "checked"
				End If 
					%>>
					<font color="red">鎖定違規地</font>
			<%End if%>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<%if sys_City="台南市" then %>
						<input type="text" class="btn5" size="3" value="<%=Trim(request("IllegalZip"))%>" name="IllegalZip" onBlur="getDriverZip(this,'IllegalAddress');" onkeydown="funTextControl(this);">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						區號
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
					<%end if%>
					<%if sys_City="高雄市" then %>
						區號
						<input type="text" class="btn5" size="3" value="<%=bIllZip%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
						<div id="LayerIllZip" style="position:absolute ; width:160px; height:30; z-index:0;  border: 1px none #000000;""><%
					if trim(bIllZip)<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&trim(bIllZip)&"'"
						set rsZip1=conn.execute(strZip1)
						if not rsZip1.eof then
							response.write trim(rsZip1("ZipName"))
						end if
						rsZip1.close
						set rsZip1=nothing
					end if
					%></div><br>
					<%end if%>
			<%	if sys_City="花蓮縣x" Then%>
					鄉鎮路段
					<input type="text" name="CityStreet" value="<%
					response.write Trim(request("CityStreet"))
					%>" onkeydown="funTextControl(this);" onblur="funChkCityStreet();">例如"花蓮市中正路"
			<%	End if%>
					<input type="text" size="<%
					if sys_City="台南市" Then
						response.write "21"
					Else
						response.write "29"
					End If 
					%>" value="<%
				if sys_City="花蓮縣x" Then
					response.write Replace(bIllegalAddress,Trim(request("CityStreet")),"")
				Else
					response.write bIllegalAddress
				End If 
					
					%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);" <%
					if sys_City="台南市" Then Response.Write " onfocus=""autoKeyEnd();"""
					%> >
					<input type="checkbox" name="chkHighRoad" value="1" <%
				if sys_City="雲林縣" then
					if trim(request("chkHighRoad"))="1" then response.write "checked" End If 
				Else
					if trim(request("chkHighRoad"))="1" Or bHighSpeedRoad="1" then response.write "checked" End If 
				End if
					%> onclick="setIllegalRule()" <%if sys_City="南投縣" then response.write "disabled"%>><span class="style1">快速道路</span>
					<%if sys_City="台中市" then %>
						
						<table >
						<tr>
						<td>
						區號
						<input type="text" class="btn5" size="3" value="<%=bIllZip%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
						</td>
						<td style="vertical-align:text-top;">
						<div id="LayerIllZip" style="position:absolute ; width:160px; height:30; z-index:0;  border: 1px none #000000;""><%
					if trim(bIllZip)<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&trim(bIllZip)&"'"
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
			</tr>
<%end if%>
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then%>

			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%=bRuleSpeed%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="5">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="9" size="10" value="<%=bRule1%>" name="Rule1" onkeyup="getRuleData1();" onfocus="this.select()" onkeydown="funTextControl(this);" onchange="DelSpace1();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>&sBillTypeID=2","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")' alt="查詢法條">
					<img src="../Image/BillLawPlusButton.jpg" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer1" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(bRule1)<>"" then
						strRule1="select IllegalRule from Law where ItemID='"&trim(bRule1)&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule1=conn.execute(strRule1)
						if not rsRule1.eof then
							response.write trim(rsRule1("IllegalRule"))
						end if
						rsRule1.close
						set rsRule1=nothing
						if trim(bRule4)<>"" then
							response.write "("&bRule4&")"
						end if
					end if
					%></div>
					<input type="hidden" name="ForFeit1" value="<%=trim(bForFeit1)%>">
					</td>
					</tr>
					</table>
				</td>
			</tr>
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then %>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%
						response.write bRuleSpeed
					%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled value="<%=otherbIllegalSpeed%>">
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規法條二</div></td>
				<td colspan="5">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="9" size="10" value="<%
				if sys_City<>"南投縣" And sys_City<>"台中市" then
					response.write bRule2
				end if
					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" onchange="DelSpace2();" style=ime-mode:disabled onBlur="TabFocus()">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer2" style="position:absolute ; width:590px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if sys_City<>"南投縣" And sys_City<>"台中市" then
					if trim(bRule2)<>"" then
						strRule2="select IllegalRule from Law where ItemID='"&trim(bRule2)&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule2=conn.execute(strRule2)
						if not rsRule2.eof then
							response.write trim(rsRule2("IllegalRule"))
						end if
						rsRule2.close
						set rsRule2=nothing
					end if
				end if
					%></div>
					<input type="hidden" name="ForFeit2" value="<%
				if sys_City<>"南投縣" And sys_City<>"台中市" then
					response.write trim(bForFeit2)
				end if
					%>">
					<img src="space.gif" width="590" height="2">
					<img src="../Image/Law3.jpg" width="45" height="25" onclick='InsertLaw()' alt="違規法條三">
					</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw1" align="right"></td>
				<td colspan="5" id="TDLaw2"></td>
			</tr>
<%if sys_City="嘉義市" then %>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%=bIllegalAddressID%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<input type="hidden" name="OldIllegalAddressID" value="<%=bIllegalAddressID%>">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="29" value="<%=bIllegalAddress%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()"><span class="style1">快速道路</span>
				</td>
			</tr>
<%end if%>
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" and sys_City<>"高雄市" And sys_City<>ApconfigureCityName then%>

			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%
						response.write bRuleSpeed
					%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" value="<%=otherbIllegalSpeed%>" onkeydown="funTextControl(this);" style=ime-mode:disabled >
				</td>
			</tr>
<%end if%>

			<tr>
				<td id="DLDate1" bgcolor="#FFFFCC" align="right"></td>
				<td id="DLDate2">
				<input type="hidden" size="6" value="" maxlength="7" name="DealLineDate" onBlur="DealLineDateReplace()" style=ime-mode:disabled>
				</td>
				<td id="DLDate3" bgcolor="#FFFFCC" align="right"></td>
				<td id="DLDate4" colspan="3"></td>
			</tr>

			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="BillMem1" value="<%=trim(bLoginID1)%>" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer12" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem1)%></div>
					<input type="hidden" value="<%=trim(bBillMemID1)%>" name="BillMemID1">
					<input type="hidden" value="<%=trim(bBillMem1)%>" name="BillMemName1">
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼2</div></td>
		  		<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="BillMem2" value="<%=trim(bLoginID2)%>" onkeyup="getBillMemID2();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer13" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem2)%></div>
					<input type="hidden" value="<%=trim(bBillMemID2)%>" name="BillMemID2">
					<input type="hidden" value="<%=trim(bBillMem2)%>" name="BillMemName2">
					</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
		<%If sys_City="台中市" then%>
				<td bgcolor="#FFFFCC"><div align="right">舉發人姓名</div></td>
		  		<td colspan="5">
					<input type="text" name="BillMemNameQry" value="" size="15"  onkeyup="getBillMemData();" onkeydown="funTextControl(this);">
					<div id="Layer50" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
					<input type="button" name="BillMemNameQry_But" value="代入" onclick="InnerBillMemData();">
					<input type="hidden" name="BillMemID1Qry" value="" >
					<input type="hidden" name="BillUnitTypeID1Qry" value="" >
					<input type="hidden" name="BillUnitTypeID1" value="" >
					<input type="hidden" name="BillUnitIDQry" value="" >
					<input type="hidden" name="BillUnitNameQry" value="" >

					<input type="hidden" size="10" name="BillMem3" value="<%=trim(bLoginID3)%>" >
					<div id="Layer14" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem3)%></div>
					<input type="hidden" value="<%=trim(bBillMemID3)%>" name="BillMemID3">
					<input type="hidden" value="<%=trim(bBillMem3)%>" name="BillMemName3">

					<input type="hidden" size="10" name="BillMem4" value="<%=trim(bLoginID4)%>" >
					<div id="Layer17" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem4)%></div>
					<input type="hidden" value="<%=trim(bBillMemID4)%>" name="BillMemID4">
					<input type="hidden" value="<%=trim(bBillMem4)%>" name="BillMemName4">
				</td>
		<%else%>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼3</div></td>
		  		<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="BillMem3" value="<%=trim(bLoginID3)%>" onkeyup="getBillMemID3();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer14" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem3)%></div>
					<input type="hidden" value="<%=trim(bBillMemID3)%>" name="BillMemID3">
					<input type="hidden" value="<%=trim(bBillMem3)%>" name="BillMemName3">
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼4</div></td>
		  		<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="BillMem4" value="<%=trim(bLoginID4)%>" onkeyup="getBillMemID4();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer17" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem4)%></div>
					<input type="hidden" value="<%=trim(bBillMemID4)%>" name="BillMemID4">
					<input type="hidden" value="<%=trim(bBillMem4)%>" name="BillMemName4">
					</td>
					</tr>
					</table>
				</td>
		<%End if%>
		
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td <%
				if sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="雲林縣" or sys_City="台南市" Or sys_City="新竹市" or sys_City="嘉義市" or sys_City="保二總隊三大隊一中隊" then
					response.write "colspan='5'"
				end if
				%>>
					<table >
					<tr>
					<td>
					<input type="text" size="10" name="BillUnitID" value="<%=trim(bBillUnitID)%>" onkeyup="getUnit();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer6" style="position:absolute ; width:180px; height:30px; z-index:0;  border: 1px none #000000;"><%
					if trim(bBillUnitID)<>"" then
						strUnitName="select UnitName from UnitInfo where UnitID='"&trim(bBillUnitID)&"'"
						set rsUnitName=conn.execute(strUnitName)
						if not rsUnitName.eof then
							response.write trim(rsUnitName("UnitName"))
						end if
						rsUnitName.close
						set rsUnitName=nothing
					end if
					%></div>
				<%if sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="嘉義市" or sys_City="新竹市" or sys_City="雲林縣" or sys_City="台南市" or sys_City="保二總隊三大隊一中隊" then%>

					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
					民眾檢舉時間
					<input type="text" name="JurgeDay" value="<%=bReportCaseJurgeDay%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" style=ime-mode:disabled onblur="this.value=this.value.replace(/[^\d]/g,'');">
					
					<span class="style10">不可超過違規日七天
					</div>
				<%End if%>
					</td>
					</tr>
					</table>
				</td>
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City<>"雲林縣" and sys_City<>"新竹市" and sys_City<>"台南市" and sys_City<>"保二總隊三大隊一中隊" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
					if bBillType<>"1" then
						response.write ginitdt(date)
					else
						if trim(bBillFillDate)="" then
							response.write ginitdt(date)
						else
							response.write bBillFillDate
						end if
					end if
					%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" onkeydown="funTextControl(this);" style=ime-mode:disabled onkeydown="funTextControl(this);" style=ime-mode:disabled <%
				if sys_City="高雄市" Or sys_City=ApconfigureCityName then '打滿跳下格%>
					onkeyup="FullToGoNextTag(6,'BillFillDate','JurgeDay')"
				<%end if%>>
					&nbsp; &nbsp; &nbsp; &nbsp; 民眾檢舉時間
					<input type="text" name="JurgeDay" value="<%=bReportCaseJurgeDay%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" <%
				if sys_City="高雄市" Or sys_City=ApconfigureCityName then '打滿跳下格%>
					onkeyup="FullToGoNextTag(6,'JurgeDay','ProjectID')"
				<%end if%> onblur="this.value=this.value.replace(/[^\d]/g,'');">
					</td>
					<td style="vertical-align:text-top;">
					<div style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style10">不可超過違規日七天</span>
					</div>
					</td>
					</tr>
					</table>
				</td>
<%end if%>
			</tr>
		<%'If sys_City="新竹市" or sys_City="嘉義市" or sys_City="花蓮縣" or sys_City="台南市" or sys_City="彰化縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="苗栗縣" or sys_City="雲林縣" or sys_City="保二總隊三大隊二中隊" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">身分證號<br><span class="style10">非轉歸責案件勿填</span></div></td>
		  		<td>
					<input type="text" size="10" name="DriverPID" onBlur="this.value=this.value.toUpperCase();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">應到案處所<br><span class="style10">非轉歸責案件勿填</span></div>
				
				</td>
		  		<td colspan="5">
					<table >
					<tr>
					<td>
					<input type="text" size="5" value="" name="MemberStation" onkeyup="getStation();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:120px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					</span>
					</td>
					</tr>
					</table>
				</td>
			</tr>
		<%'End if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">專案代碼</div></td>
				<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if sys_City="屏東縣" Then
					If trim(request("ProjectID"))<>"" then
						response.write trim(request("ProjectID"))
					End if
				end if
					%>" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
			<%if sys_City="苗栗縣" then%>
					<font style="font-size:12px;">檢舉達人1 / 拖吊9 </font>
			<%End If %>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					</td>
					</tr>
					</table>
				</td>
<%if sys_City="雲林縣" then%>	
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if bBillType<>"1" then
					response.write ginitdt(date)
				else
					if trim(bBillFillDate)="" then
						response.write ginitdt(date)
					else
						response.write bBillFillDate
					end if
				end if
				%>" maxlength="7" name="BillFillDate" onkeydown="funTextControl(this);" onBlur="getDealLineDate();" style=ime-mode:disabled>
				</td>
<%else%>
				<td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="2" size="4" value="<%
				if sys_City="新竹市" then
					if trim(request("CarAddID"))="8" then
						response.write trim(request("CarAddID"))
					end if

				elseif sys_City="高雄市" Or sys_City=ApconfigureCityName then
					response.write bCarAddId
				end if
					%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer110" style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機 /8拖吊<br>/9(550cc)重機 /10計程車/ 11危險物品</span>
					</div>
					</td>
					</tr>
					</table>
				</td>
<%end if%>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align='right'>採證工具</div></td>
				<td>
					<table >
					<tr>
					<td>
					<input maxlength="1" size="4" value="<%
				if sys_City="嘉義縣" or sys_City="花蓮縣" or sys_City="高雄縣" then
					response.write bUseTool
				end if
					%>" name="UseTool"  onBlur="getFixID();" onkeydown="funTextControl(this);" type='text' style=ime-mode:disabled> 
					</td>
					<td style="vertical-align:text-top;">
			        <div id="Layer11" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%=request("FixID")%>' onBlur="setFixEquip();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					</div>
					<font class="style7"> 1固定桿/ 2雷達三腳架/ 3相機/<%
				If sys_City="台南市" Then
					response.write " 4車載攝影機/ 5科技執法/"
				ElseIf sys_City="基隆市" Then
					response.write " 4雷射測速鎗/"
				End If 
					%> 8逕舉手開單</font>
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">備註</div></td>
				<td>
					<input type="text" size="15" value="" name="Note" onkeydown="funTextControl(this);" style=ime-mode:active>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">代保管物</div></td>
				<td>
					1. <input type="text" size="2" value="" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
					<div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>
					<input type="hidden" value="" name="Fastener1Val">

					2. <input type="text" size="2" value="" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
					 <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>
	                 <input type="hidden" value="" name="Fastener2Val">
				</td>
			</tr>
<tr>
<td bgcolor="#FFFFCC"><div align="right">來源平台</div></td>
<td><input type="text" size="20" value="" name="FromNote" onkeydown="funTextControl(this);" style=ime-mode:active></td>
<td bgcolor="#FFFFCC"><div align="right">平台單號/案號</div></td>
<td colspan="3"><input type="text" size="20" value="" name="FromNoteId" onkeydown="funTextControl(this);" style=ime-mode:active></td>
</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="6">
					<input type="button" name="save1" value="儲 存 <%
					if sys_City="台東縣" or sys_City="高雄縣" then
						response.write "F9"
					else
						response.write "F2"
					end if
					%>" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(223,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
<%if (sys_City<>"高雄市" and sys_City<>ApconfigureCityName) or trim(request("BillReCover"))<>"1" then%>
	<%'高雄檢舉系統不要出現按鈕
	If Trim(request("ReportCaseSn"))="" then%>
					<img src="/image/space.gif" width="29" height="8">
	
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_Car_Report.asp'" value="清 除 F4" class="btn1">
	<%end if%>
<%end if%>
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry()" value="查 詢 <%
					if sys_City="高雄市" Or sys_City=ApconfigureCityName then
						response.write "F5"
					else
						response.write "F6"
					end if
					%>" class="btn1">
					<input type="hidden" name="kinds" value="">
                    <span class="style1">
                    <span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
					<span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="funPrintCaseList_Report();" value="建檔清冊 <%
					if sys_City="高雄縣" then
						response.write "F2"
					else
						response.write "F10"
					end if
					%>" class="btn1">
                </span>
<%if (sys_City<>"高雄市" and sys_City=ApconfigureCityName) or trim(request("BillReCover"))<>"1" then%>
				<br>

				<div id="Layer1f69" style=" width:450px; height:14px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				
				</div>
	
				<img src="/image/space.gif" width="250" height="8">
	<%'高雄檢舉系統不要出現按鈕
	If Trim(request("ReportCaseSn"))="" then%>
				<input type="button" name="SubmitBack2" onClick="location='BillKeyIn_Report_Back.asp?PageType=First'" value="<< 第一筆 Home" class="btn1">
				<img src="/image/space.gif" width="29" height="8">
				<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Report_Back.asp?PageType=Back'" value="< 上一筆 PgUp" class="btn1">
	<%End if%>
				<!--<div id="Layer1c69" style=" width:360px; height:14px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style7">使用上一筆搜尋功能只能查詢到自己建檔且未入案的舉發單</span>
				</div>-->
				<img src="/image/space.gif" width="220" height="8">
<%end if%>
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案處所 -->
		<%'If sys_City<>"嘉義市" And sys_City<>"新竹市" And sys_City<>"花蓮縣" And sys_City<>"台南市" And sys_City<>"彰化縣" And sys_City<>"基隆市" And sys_City<>"澎湖縣" And sys_City<>"苗栗縣" And sys_City<>"雲林縣" And sys_City<>"保二總隊三大隊二中隊" then%>
				<input type="hidden" size="4" value="" name="XXXMemberStation" onkeyup="getStation();">
				<div id="XXXLayer5" style="position:absolute ; width:241px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
		<%'End If %>
				<!-- 附加說明 -->
				<input type="hidden" value="<%=bRule4%>" name="Rule4">
				<input type="hidden" value="<%=bRecordDate%>" name="otherRecordDate">
				<input type="hidden" name="Tmp_Order" value="<%=Session("BillCnt_Report")+1%>">

				</td>
			</tr>
			<tr>
				<td colspan="6">
				<font style="font-size: 10pt;">
				<span class="style7">＊ 使用上一筆搜尋功能只能查詢到自己建檔且未入案的舉發單</span>
				<br>
				＊ 臨時車牌、試車牌等前面有中文字的車號，輸入車號時，僅需輸入英數字，中文字不可輸入
				</font>
	<br>
<span class="style9">＊(重點工作報表針對特殊車種 需要在建檔時 輔助車種中 輸入   3砂石/ 8拖吊 /10計程車)</span>
				</td>
			</tr>
		</table>		
<%if (sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣") and trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"" then%>
		<table width='100%' border='0' align="center" cellpadding="0" cellspacing="0">
	<%
	strI="select * from BILLILLEGALIMAGE where billsn="&trim(request("ReCoverSn"))
	set rsI=conn.execute(strI)
	If Not rsI.eof Then
	%>
	<tr>
		<td>
		<input type="hidden" name="sys_IISImagePath" value="<%=trim(rsI("IISImagePath"))%>">
		<input type="hidden" name="sys_ImageFileNameA" value="<%=trim(rsI("ImageFileNameA"))%>">
		<input type="hidden" name="sys_ImageFileNameB" value="<%=trim(rsI("ImageFileNameB"))%>">
		<input type="hidden" name="sys_ImageFileNameC" value="<%=trim(rsI("ImageFileNameC"))%>">
	<%	if trim(rsI("ImageFileNameA"))<>"" then	%>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameA"))%>" width="460"><br>
	<%	end if%>
		</td>
		<td>
	<%	if trim(rsI("ImageFileNameB"))<>"" then	%>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameB"))%>" width="460"><br>
	<%	end if%>
		</td>
	</tr>
	<tr>
		<td>
	<%	if trim(rsI("ImageFileNameC"))<>"" then	%>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameC"))%>" width="400"><br>
	<%	end if%>
		</td>
	</tr>
	<%
	else
		response.write "查無違規影像"
	end if
	rsI.close
	set rsI=nothing
	%>
	
	</table>
<%end if%>
<%if sys_City="XX縣都不開" then%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr>
				<td>
				<div id="LayerA1">
				
				<table width='960' border='1' align="center" cellpadding="1">
				<tr bgcolor="#33FFFF">
					<td width="9%"><span class="style99">單號</span></td>
					<td width="7%"><span class="style99">車號</span></td>
					<td width="11%"><span class="style99">違規時間</span></td>
					<td width="19%"><span class="style99">違規地點</span></td>
					<td width="16%"><span class="style99">違規法條</span></td>
					<td width="7%"><span class="style99">限速、重</span></td>
					<td width="7%"><span class="style99">車速、重</span></td>
					<td width="16%"><span class="style99">舉發人</span></td>
					<td width="8%"><span class="style99">操作</span></td>
				</tr>
<%
	strBillView="select * from billbase where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID&" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is null order by RecordDate desc"
	set rsBillView=conn.execute(strBillView)
	If Not rsBillView.Bof Then rsBillView.MoveFirst 
	While Not rsBillView.Eof
		'for i=0 to 1000

%>
				<tr>
					<td><span class="style99"><%
					if trim(rsBillView("BillNo"))<>"" and not isnull(rsBillView("BillNo")) then
						response.write trim(rsBillView("BillNo"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%=trim(rsBillView("CarNo"))%></span></td>
					<td><span class="style99"><%=year(rsBillView("IllegalDate"))&"/"&Month(rsBillView("IllegalDate"))&"/"&Day(rsBillView("IllegalDate"))&" "&Hour(rsBillView("IllegalDate"))&":"&Minute(rsBillView("IllegalDate"))%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("IllegalAddress"))<>"" and not isnull(rsBillView("IllegalAddress")) then
						response.write trim(rsBillView("IllegalAddress"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("Rule1"))<>"" and not isnull(rsBillView("Rule1")) then
						response.write trim(rsBillView("Rule1"))
					else
						response.write "&nbsp;"
					end if
					if trim(rsBillView("Rule2"))<>"" and not isnull(rsBillView("Rule2")) then
						response.write "/"&trim(rsBillView("Rule2"))
					end if
					if trim(rsBillView("Rule3"))<>"" and not isnull(rsBillView("Rule3")) then
						response.write "/"&trim(rsBillView("Rule3"))
					end if
					if trim(rsBillView("Rule4"))<>"" and not isnull(rsBillView("Rule4")) then
						response.write "/"&trim(rsBillView("Rule4"))
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("RuleSpeed"))<>"" and not isnull(rsBillView("RuleSpeed")) then
						response.write trim(rsBillView("RuleSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("IllegalSpeed"))<>"" and not isnull(rsBillView("IllegalSpeed")) then
						response.write trim(rsBillView("IllegalSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("BillMem1"))<>"" and not isnull(rsBillView("BillMem1")) then
						response.write trim(rsBillView("BillMem1"))
					else
						response.write "&nbsp;"
					end if
					if trim(rsBillView("BillMem2"))<>"" and not isnull(rsBillView("BillMem2")) then
						response.write "/"&trim(rsBillView("BillMem2"))
					end if
					if trim(rsBillView("BillMem3"))<>"" and not isnull(rsBillView("BillMem3")) then
						response.write "/"&trim(rsBillView("BillMem3"))
					end if
					if trim(rsBillView("BillMem4"))<>"" and not isnull(rsBillView("BillMem4")) then
						response.write "/"&trim(rsBillView("BillMem4"))
					end if
					%></span></td>
					<td><span class="style99">
						<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN=<%=trim(rsBillView("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
					</span>
					
					</td>
				</tr>
<%		rsBillView.MoveNext
		'next
	Wend
	rsBillView.close
	set rsBillView=nothing
%>
				</table>
				
				</div>
				</td>
			</tr>
		</table>
<%end if%>
	</form>
<%
conn.close
set conn=nothing
%>
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
var TDVipCarErrorLog=0;
var TDIllZipErrorLog=0;
var SpeedError=0;
var TodayDate=<%=ginitdt(date)%>;
var ButtonSubmit=0;
var TDCityStreetErrorLog=0;

<%if sys_City="彰化縣" then %>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="保二總隊三大隊一中隊" then%>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="基隆市" or sys_City="苗栗縣" then%>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="新竹市" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="嘉義市" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="台南市" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="雲林縣" then %>
MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="高雄市" then %>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="花蓮縣" then%>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City=ApconfigureCityName then %>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");

<%elseif sys_City="台中市" then%>
MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMemNameQry||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="連江縣" then%>
MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");

<%else%>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%end if%>

//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.ReportChk.checked==true){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入單號。";
		}else if(ReadBillNo.length!=9){
			error=error+1;
			errorString=errorString+"\n"+error+"：單號不足九碼。";
		}
	}
<%if sys_City="高雄市" and not (trim(request("BillReCover"))="1" and trim(request("ReCoverSn"))<>"") then %>
	ReadBillNo=myForm.Billno1.value.replace(' ','');
	if (ReadBillNo==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：僅可建檔逕舉手開單案件，請勾選『逕舉手開單』，並輸入單號。";
	}
<%end if%>
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
	if (((myForm.Rule1.value.substr(0,2))=="35" || (myForm.Rule2.value.substr(0,2))=="35") && (myForm.IsVideo[0].checked==false && myForm.IsVideo[1].checked==false))
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：法條為35條時，請輸入有無全程錄影。";
	}
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
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%elseif sys_City="台中市" then%>
	}else if (!ChkIllegalDateTC89(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	}else if (!ChkIllegalDateTC89(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%elseif sys_City="高雄市" then%>
	}else if (!ChkIllegalDate2M_KS(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	}else if (!ChkIllegalDate2M_KS(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
	}
<%elseif sys_City="基隆市" then%>
	}else if (!ChkIllegalDateTC(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過30天期限，如確定要建檔請勾選上方強制建檔，並在備註輸入超過期限原因。";
	}else if (!ChkIllegalDateTC(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過30天期限原因。";
	}
<%else	'109/12/1 逕舉改成60天%>
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限。";
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
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
<%if sys_City="花蓮縣x" then %>
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
<%if sys_City<>"新竹市" and sys_City<>"嘉義縣" and sys_City<>"嘉義市" then%>
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%else%>
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value) && myForm.ReportChk.checked==true){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%end if%>
<%if sys_City="苗栗縣" then%>
	}else if (!ChkIllegalDateML(myForm.BillFillDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過二個月期限。";
	}
<%else%>
	}else if (!ChkIllegalDateML(myForm.BillFillDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過二個月期限。";
	}
<%end if%>
<%'If sys_City="新竹市" or sys_City="嘉義市" or sys_City="花蓮縣" or sys_City="雲林縣" then%>
	if (myForm.MemberStation.value!="" || myForm.DriverPID.value!=""){
		if ((myForm.Rule1.value.substr(0,2))=="35" && myForm.DriverPID.value=="")
		{
			//酒駕案件可以不填身分證號
		}
		else{
			if (myForm.MemberStation.value=="" || myForm.DriverPID.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：轉歸責案件，身分證號與應到案處所都要輸入。";
			}
		}
	}
<%'end if %>
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
	}else if (!ChkIllegalDateML(myForm.DealLineDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過二個月期限。";
	}
<%else%>
	}else if (!ChkIllegalDateML(myForm.DealLineDate.value) && myForm.CaseInByMem.checked==false){
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
	if (TDFastenerErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 輸入錯誤。";
	}
	if (TDFastenerErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物2 輸入錯誤。";
	}
	if (myForm.Fastener1.value==myForm.Fastener2.value && myForm.Fastener1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 與代保管物2 重複。";
	}
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
	if (((myForm.Rule1.value.substr(0,3))=="351") || ((myForm.Rule1.value.substr(0,3))=="352") || ((myForm.Rule1.value.substr(0,3))=="356") || ((myForm.Rule2.value.substr(0,3))=="351") || ((myForm.Rule2.value.substr(0,3))=="352") || ((myForm.Rule2.value.substr(0,3))=="356")){
		if (myForm.Rule1.value!="351000031" && myForm.Rule1.value!="352000021" && myForm.Rule2.value!="351000031" && myForm.Rule2.value!="352000021"){
			if ((myForm.ProjectID.value !="W1") && (myForm.ProjectID.value !="W2")){
				error=error+1;
				errorString=errorString+"\n"+error+"：酒駕案件，\n吹測  請於  專案代碼  輸入 W1 \n抽測  請於  專案代碼  輸入 W2";
			}
		}
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
			errorString=errorString+"\n"+error+"：告示單號不可少於11碼。";
		}
	}	
	if (myForm.AcceptBatchNumberChk.checked==true && myForm.AcceptBatchNumber.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：有勾選批號檢查，但是未輸入批號，請輸入批號或取消勾選。";
	}
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
<%if sys_City="雲林縣" then %>
	if (myForm.chkHighRoad.checked==true && myForm.IllegalAddress.value.indexOf('快速')==-1)
	{
		error=error+1;
		errorString=errorString+"\n"+error+"：違規地點如勾選快速道路，違規地點名稱必須包含『快速』兩字。";
	}
<%end if%>
	if (error==0 && ButtonSubmit==0){
			getChkCarIllegalDate();
	}else{
		alert(errorString);
	}
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function setChkCarIllegalDate(CarCnt,Illdate,RuleDetail)
{
	var ErrorStr="";
	if (CarCnt=="1"){
		ChkCarIlldateFlag="1";
	}else{
		ChkCarIlldateFlag="0";
	}
	<%if sys_City="台中市" then%>
	if (!ChkIllegalDateTC(myForm.IllegalDate.value)){
		ErrorStr=ErrorStr+"違規日期已超過30天。\n";
	}
	<%end if%>	
	if ((myForm.CarSimpleID.value=="1" || myForm.CarSimpleID.value=="2") && (myForm.Rule1.value.substr(0,3)=="316" || myForm.Rule2.value.substr(0,3)=="316" || myForm.Rule1.value=="4511301" || myForm.Rule2.value=="4511301"))
	{
		ErrorStr=ErrorStr+"您輸入的法條為機車法條，但是車種為汽車，請確認是否正確。\n";
	}
<%if sys_City="雲林縣" or sys_City="南投縣" or sys_City="屏東縣" then%>
	if (myForm.ReportChk.checked!=true){
		getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
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
		}
	}else{	//逕舉手開單+攔停天數
		getDealDateValue="45";
	<%if sys_City="屏東縣" then%>
		BFillDateTemp=myForm.IllegalDate.value;
	<%else%>
		BFillDateTemp=myForm.BillFillDate.value;
	<%end if%>
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
		}
	}
	<%if sys_City="南投縣" then%>
	if (eval(myForm.DealLineDate.value) < eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日小於填單日加"+getDealDateValue+"天，請確認是否正確。\n";
	}
	<%elseif sys_City="屏東縣" then%>
	if (eval(myForm.DealLineDate.value) < eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日小於"+getDealDateValue+"天，請確認是否正確。\n";
	}
	<%else%>
	if (eval(myForm.DealLineDate.value) != eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日不是填單日加"+getDealDateValue+"天，請確認是否正確。\n";
	}
	<%end if%>
<%end if%>
	<%if sys_City="台中市" then%>
	if (((myForm.Rule1.value.substr(0,2))=="55" || (myForm.Rule2.value.substr(0,2))=="55") && (myForm.ReportChk.checked==false)){
		ErrorStr=ErrorStr+"第55條不可逕行舉發，請確認是否正確。\n";
	}
	if (RuleDetail==8){
		ErrorStr=ErrorStr+"此車號已開過使用吊(註)銷之牌照行駛，請確認是否正確。\n";
	}
	<%end if %>
	if (RuleDetail==1 || RuleDetail==3){
		ErrorStr=ErrorStr+"違規事實與簡式車種不符，請確認是否正確。\n";
	}
	if (ChkCarIlldateFlag=="1"){
	<%if sys_City="新竹市" Or sys_City="基隆市" Or sys_City="台南市" Or sys_City="台東縣" Or sys_City="雲林縣" then%>
		ErrorStr=ErrorStr+"此車號於"+Illdate+"，有違規舉發記錄，請確認有無連續開單。\n";
	<%else%>
		ErrorStr=ErrorStr+"此車號於"+Illdate+"，有相同違規舉發，請確認有無連續開單。\n";
	<%end if %>
	}
	<%if sys_City="南投縣" then%>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) <= 10){
				ErrorStr=ErrorStr+"\n"+ErrorStr+"車速超過限速未超過10公里\n";
			}
		}				
	}
	<%end if%>
	<%if sys_City="高雄市" then%>
	if (myForm.IllegalAddressID.value=="00346" || myForm.IllegalAddressID.value=="00501")
	{
		if (myForm.Rule1.value.substr(0,2)=="53" || myForm.Rule2.value.substr(0,2)=="53")
		{
			ErrorStr=ErrorStr+"\n輕軌共用路口，注意引用法條，請確認是否正確。\n";
		}
	}
	<%end if%>	
	<%if sys_City="基隆市" then%>
	if (myForm.IllegalAddress.value != "" && myForm.IllegalSpeed.value!="" && myForm.RuleSpeed.value!=""){
		if ((myForm.IllegalAddress.value.indexOf("台62甲") == -1) || (myForm.IllegalAddress.value.indexOf("臺62甲") == -1)){
			if ((parseInt(myForm.IllegalSpeed.value)-parseInt(myForm.RuleSpeed.value) ) >= 60){
				ErrorStr=ErrorStr+"\n台62甲線，車速超過限速超過60公里，請注意!。";
			}
		}
	}
	<%end if%>	
	if (RuleDetail==2){
		alert("舉發單位代號輸入錯誤。\n");
<%if sys_City="高雄市" then%>
	}else if (RuleDetail==3 || RuleDetail==4){
		alert("此車號為業管車輛。\n");
<%end if%>
<%if sys_City="南投縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。\n");
<%elseif sys_City="新竹市" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。\n");
<%end if%>
<%if sys_City="台中市" or sys_City="雲林縣" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。\n");
	
<%elseif sys_City<>"台東縣" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間、違規地點，有相同違規法條，請確認是否正確。\n");
<%end if%>
	}else{
		if (ErrorStr!=""){
			if(confirm(ErrorStr+"\n是否確定要存檔？")){
				myForm.kinds.value="DB_insert";
				ButtonSubmit=1;
				myForm.submit();
			}
		}else{
			myForm.kinds.value="DB_insert";
			ButtonSubmit=1;
			myForm.submit();
		}
	}
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function getChkCarIllegalDate(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2=myForm.Rule2.value;
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	NewBillUnitID=myForm.BillUnitID.value;
	NewIllegalAddress=myForm.IllegalAddress.value;
	NewJurgeDay=myForm.JurgeDay.value;

	<%if sys_City="台中市" then%>
		var AcceptBatchNumberChk_Temp;
		if (myForm.AcceptBatchNumberChk.checked==true)
		{
			AcceptBatchNumberChk_Temp="1";
		}else{
			AcceptBatchNumberChk_Temp="0";
		}
		runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress+"&BillCheck=1&BillNO="+myForm.Billno1.value+"&AcceptBatchNumber="+myForm.AcceptBatchNumber.value+"&AcceptBatchNumberChk="+AcceptBatchNumberChk_Temp+"&JurgeDay="+NewJurgeDay+"&nowTime=<%=now%>");
	<%else%>
		runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress+"&JurgeDay="+NewJurgeDay+"&nowTime=<%=now%>");
	<%end if%>
}
<%if sys_City="花蓮縣x" then%>
function funChkCityStreet(){
	if (myForm.CityStreet.value != "")
	{
		runServerScript("getChkCityStreet.asp?CityStreet="+myForm.CityStreet.value);
	}
}
<%end if %>
//是否為特殊用車&檢查是否有同車號在同一天建檔
function getVIPCar(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			//alert("車牌格式錯誤，如該車輛無車牌或舊式車牌則可忽略此訊息！");
			runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=2");
		}else{
			runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=2");
		<%if sys_City<>"高雄市" and sys_City<>"苗栗縣" and sys_City<>"新竹市" and sys_City<>"連江縣" then%>
			myForm.CarSimpleID.value=CarType;
		<%end if%>
		<%if sys_City=ApconfigureCityName then%>
			myForm.IllegalDate.select();
		<%end if%>
		}
	}else{
		Layer7.innerHTML=" ";
		myForm.CarSimpleID.value="";
	}
<%If Trim(request("ReportCaseSn"))<>"" then%>
	myForm.CarSimpleID.focus();
<%end if %>
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

//檢查輔助車種
function getAddID(){
	//myForm.CarAddID.value=myForm.CarAddID.value.replace(/[^\d]/g,'');
	if (myForm.CarAddID.value.length>0){
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7" && myForm.CarAddID.value != "8" && myForm.CarAddID.value != "9" && myForm.CarAddID.value != "10" && myForm.CarAddID.value != "11" && myForm.CarAddID.value != "0"<%
			if sys_City="雲林縣" Then
				response.write " && myForm.CarAddID.value != ""12"""
			End If 
		%>){
			alert("輔助車種填寫錯誤!");
			//myForm.CarAddID.value = "";
			myForm.CarAddID.select();
		}
	}
}
//檢查簡式車種
function getRuleAll(){
	//myForm.CarSimpleID.value=myForm.CarSimpleID.value.replace(/[^\d]/g,'');
	if (myForm.CarSimpleID.value.length>0){
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "5" && myForm.CarSimpleID.value != "6" && myForm.CarSimpleID.value != "7"){
			alert("簡式車種填寫錯誤!");
			myForm.CarSimpleID.select();
			//myForm.CarSimpleID.value = "";
		}
	}

}
//法條刪掉其他符號
function DelSpace1(){
	myForm.Rule1.value=myForm.Rule1.value.replace(/[^\d]/g,'');
	myForm.Rule4.value="";
	getRuleData1();
}
function DelSpace2(){
	myForm.Rule2.value=myForm.Rule2.value.replace(/[^\d]/g,'');
	getRuleData2();
}
function DelSpace3(){
	myForm.Rule3.value=myForm.Rule3.value.replace(/[^\d]/g,'');
	getRuleData3();
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
<%if sys_City="台南市" then %>
		if ((myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,3))=="433") {

			alert("提醒：第43條第1項各款與第3項，應加開第4項吊扣牌照之舉發。\n( 本訊息僅為提醒，按確定後可繼續作業 )");
		}
<%elseif sys_City="澎湖縣" then %>
		if ((myForm.Rule1.value.substr(0,3))=="431") {

			alert("提醒：第43條第1項各款，應另單加開第4項處車主之舉發。\n( 本訊息僅為提醒，按確定後可繼續作業 )");
		}
<%end if%>
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
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
<%if sys_City="台南市" then %>
		if ((myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,3))=="433") {

			alert("提醒：第43條第1項各款與第3項，應加開第4項吊扣牌照之舉發。\n( 本訊息僅為提醒，按確定後可繼續作業 )");
		}
<%elseif sys_City="澎湖縣" then%>
		if ((myForm.Rule2.value.substr(0,3))=="431") {

			alert("提醒：第43條第1項各款，應另單加開第4項處車主之舉發。\n( 本訊息僅為提醒，按確定後可繼續作業 )");
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
//違規事實3(ajax)
function getRuleData3(){
	if (myForm.Rule3.value.length > 6){
		var Rule3Num=myForm.Rule3.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=3&RuleID="+Rule3Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		//CallChkLaw3();
	}else if (myForm.Rule3.value.length <= 6 && myForm.Rule3.value.length > 0){
		Layer3.innerHTML=" ";
		myForm.ForFeit3.value="";
		TDLawErrorLog3=1;
	}else{
		Layer3.innerHTML=" ";
		myForm.ForFeit3.value="";
		TDLawErrorLog3=0;
	}
}
//增加違規法條
function InsertLaw(){
	if (myForm.ReportChk.checked==false){
		alert("因逕舉舉發單格式只有兩個法條，如非逕舉手開單案件，請勿輸入超過兩個法條！");
	}else{
	TDLawNum=1;
	TDLaw1.innerHTML="違規法條三";
	TDLaw2.innerHTML="<table ><tr><td><input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();' onchange='DelSpace1();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw3()' alt='查詢法條'></td> <td style='vertical-align:text-top;'> <div id='Layer3' style='position:absolute ; width:589px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''></td></tr></table>";

	if (myForm.ReportChk.checked==true){
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="保二總隊三大隊一中隊" then%>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="基隆市" or sys_City="苗栗縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="新竹市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="花蓮縣" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMemNameQry||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
	}else{
	<%if sys_City="彰化縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="保二總隊三大隊一中隊" then%>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="基隆市" or sys_City="苗栗縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="新竹市" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
		MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="花蓮縣" then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
		MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMemNameQry||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
		MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
	}

	myForm.Rule3.focus();
	}
}
function OpenQueryLaw3(){
	window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes");
}
function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
<%if sys_City<>"南投縣" and sys_City<>"台中縣" and sys_City<>"雲林縣" and sys_City<>"彰化縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" and sys_City<>"高雄市" and sys_City<>ApconfigureCityName then %>
	if ((Rule1tmp.substr(0,5))!="33101" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,5))!="43102" && (Rule1tmp.substr(0,3))!="293" && (Rule2tmp.substr(0,5))!="33101" && (Rule2tmp.substr(0,2))!="40" && (Rule2tmp.substr(0,5))!="43102" && (Rule2tmp.substr(0,3))!="293"){
	<%if sys_City="屏東縣" then%>
		myForm.DealLineDate.select();
	<%else%>
		if (myForm.ReportChk.checked==false){
			myForm.BillMem1.select();
		}else{
			myForm.DealLineDate.select();
		}
	<%end if%>
	}
<%end if%>
	//MemberStationLaw="21,35,57,61,62";
	//法條代碼遇到21,35,57,61,62，應到案處所自動帶當地監理所
	//if (MemberStationLaw.indexOf(Rule1tmp.substr(0,2))!=-1 || MemberStationLaw.indexOf(Rule2tmp.substr(0,2))!=-1){
	//	myForm.MemberStation.value=<%=trim(BillLawMemberStation)%>;
	//	getStation();
	//}
}
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
		if (myForm.UseTool.value != "0" && myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3" && myForm.UseTool.value != "8" <%
	if sys_City="台南市" then
		response.write " && myForm.UseTool.value != ""4"" &&  myForm.UseTool.value != ""5"""
	elseif sys_City="基隆市" then
		response.write " && myForm.UseTool.value != ""4"""
	end if 
		%>){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.select();
			//myForm.UseTool.value = "";
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
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		if (myForm.IllegalAddressID.value!=myForm.OldIllegalAddressID.value){
			myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
		}
	}
<%end if%>
	if (event.keyCode!=13){
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
<%if sys_City="彰化縣" and session("Unit_ID")<>"JT00" then%>
		}else if (myForm.IllegalAddressID.value.length >= 3){
<%else%>
		}else if (myForm.IllegalAddressID.value.length >= 1){
<%end if%>
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	}
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	FullToGoNextTag(6,'IllegalAddressID','IllegalAddress');
	if (myForm.OldIllegalAddressID.value != myForm.IllegalAddressID.value)
	{
		myForm.IllegalZip.value="";
	}
	<%end if%>
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
//保管物品一(ajax)
function getFastener1(){
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_Fastener.asp?FaOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.Fastener1.value.length == 1){
		var FastenerNum=myForm.Fastener1.value;
		runServerScript("getFastener.asp?FastenerOrder=1&FastenerID="+FastenerNum);
	}else if (myForm.Fastener1.value.length == 0){
		Layer8.innerHTML=" ";
		TDFastenerErrorLog1=0;
		myForm.Fastener1Val.value="";
	}else{
		Layer8.innerHTML=" ";
		TDFastenerErrorLog1=1;
		myForm.Fastener1Val.value="";
	}
}
//保管物品二(ajax)
function getFastener2(){
	if (event.keyCode==<%
		if sys_City="高雄市" Or sys_City=ApconfigureCityName then
			response.write "117"
		else
			response.write "116"
		end if
		%>){	
		event.keyCode=0;
		window.open("Query_Fastener.asp?FaOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.Fastener2.value.length == 1){
		var FastenerNum=myForm.Fastener2.value;
		runServerScript("getFastener.asp?FastenerOrder=2&FastenerID="+FastenerNum);
	}else if (myForm.Fastener2.value.length == 0){
		Layer9.innerHTML=" ";
		TDFastenerErrorLog2=0;
		myForm.Fastener2Val.value="";
	}else{
		Layer9.innerHTML=" ";
		TDFastenerErrorLog2=1;
		myForm.Fastener2Val.value="";
	}
}

function getBillFillDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	if (myForm.IllegalDate.value.length >= 6 ){
		myForm.BillFillDate.value=myForm.IllegalDate.value;
		getDealLineDate();
	}
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	if (myForm.ReportChk.checked!=true){
	<%if sys_City<>"屏東縣" then%>
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
		
	<%end if%>
	}else{	//逕舉手開單+攔停天數
	
<%if (trim(Session("UnitLevelID"))<>"2" and sys_City="台中縣") or (sys_City<>"台中縣" and sys_City<>"高雄市") then%>
	<%if sys_City<>"南投縣" and sys_City<>"屏東縣" and sys_City<>"台中縣" and sys_City<>"台中市" then%>
		<%if sys_City="台中縣" or sys_City="彰化縣" or sys_City="新竹市" or sys_City="台南市" or sys_City="台東縣" or sys_City="嘉義市" or sys_City="嘉義縣" or sys_City="雲林縣" or sys_City="基隆市" or sys_City="保二總隊四大隊二中隊" or sys_City="保二總隊三大隊二中隊" then%>
			getDealDateValue="45";
			
		<%elseif sys_City="澎湖縣" or sys_City="保二總隊三大隊一中隊" then%>
			if (myForm.IsMail[0].checked==true){
				getDealDateValue=<%=getReportDealDateValue%>;
				
			}else{
				getDealDateValue=<%=getStopDealDateValue%>;
				
			}
		<%elseif sys_City="花蓮縣" then%>
			if (myForm.chkbDealLineDate.checked==true){
				getDealDateValue=30;
				
			}else{
				getDealDateValue=45;
				
			}
		<%elseif sys_City=ApconfigureCityName then%>
			getDealDateValue=<%=getReportDealDateValue%>;
		<%else%>
			getDealDateValue=<%=getStopDealDateValue%>;	//要加幾天
		<%end if%>
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
	<%end if%>
<%end if%>
	}
}
//逕舉手開單由違規日期帶入應到案日期+14
function getDealLineDate_Stop(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');

	if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}
<%if sys_City="屏東縣" or sys_City="澎湖縣" then%>
	if (myForm.ReportChk.checked!=false){
		<%if sys_City="澎湖縣" then%>
			if (myForm.IsMail[0].checked==true){
				getSDealDateValue=<%=getReportDealDateValue%>;
			}else{
				getSDealDateValue=<%=getStopDealDateValue%>;
			}
			BFillDateTemp=myForm.BillFillDate.value;
		<%else%>
			getSDealDateValue=<%
			'response.write getStopDealDateValue 99/9/8改為30天
			response.write getReportDealDateValue
		%>;
			//要加幾天
			BFillDateTemp=myForm.IllegalDate.value;
		<%end if%>
		if (BFillDateTemp.length >= 6){
			//myForm.BillFillDate.value=myForm.IllegalDate.value;
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday)
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getSDealDateValue,BFillDate);
			Dyear=parseInt(DLineDate.getFullYear())-1911;
			Dmonth=DLineDate.getMonth()+1;
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
<%end if%>
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
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.RuleSpeed.value > <%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：限速、限重超過<%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	'elseif sys_City="台東縣" then 
	'	response.write "41"
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
			<%'if sys_City="南投縣" then %>
//				if (myForm.Rule2.value=="" && myForm.Rule1.value!="4340003"){
//					myForm.Rule2.value="4340003";
//					getRuleData2();
//				}else if(TDLawNum==0 && myForm.Rule1.value!="4340003" && myForm.Rule2.value!="4340003"){
//					InsertLaw();
//					myForm.Rule3.value="4340003";
//					getRuleData3();
//				}
			<%'else%>
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
			<%'end if%>
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
			UrlStr="../BillKeyIn/ChkSpeedPW.asp";
			newWin(UrlStr,"ChkSpeedPW",350,200,300,100,"no","no","no","no");
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
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	'elseif sys_City="台東縣" then 
	'	response.write "41"
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
			<%'if sys_City="南投縣" then %>
//				if (myForm.Rule2.value=="" && myForm.Rule1.value!="4340003"){
//					myForm.Rule2.value="4340003";
//					getRuleData2();
//				}else if(TDLawNum==0 && myForm.Rule1.value!="4340003" && myForm.Rule2.value!="4340003"){
//					InsertLaw();
//					myForm.Rule3.value="4340003";
//					getRuleData3();
//				}
			<%'else%>
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
			<%'end if%>
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
			UrlStr="../BillKeyIn/ChkSpeedPW.asp";
			newWin(UrlStr,"ChkSpeedPW",350,200,300,100,"no","no","no","no");
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
			if ((thisLaw.substr(0,5))!="33101" && (thisLaw.substr(0,2))!="40" && (thisLaw.substr(0,5))!="43102" && (thisLaw.substr(0,3))!="293"){
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
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value!=""){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo.length==9){
		<%if sys_City="台中市" then%>
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum+"&AcceptBatchNumber="+myForm.AcceptBatchNumber.value+"&nowTime=<%=now%>");
		<%else%>
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum+"&nowTime=<%=now%>");
		<%end if %>
		}else{
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}
	}
}

//勾選後才可以輸入單號
function funcReportChk(){
	<%if trim(bDealLineDate)<>"" then%>
		var bDealLineDate=<%=trim(bDealLineDate)%>;
	<%else%>
		var bDealLineDate="";
	<%end if%>
	if (myForm.ReportChk.checked==true){
		myForm.Billno1.disabled=false;
		myForm.UseTool.value="8";
		//LayerDLDate.style.visibility = "visible"; 
		//LayerMStation.style.visibility = "visible";
		//myForm.MemberStation.disabled=false;
		DLDate1.innerHTML="應到案日期";
		DLDate3.innerHTML="是否郵寄";
		<%if sys_City="台中縣" then%>
			bDealLineDate=""
		<%end if%>
		<%if sys_City="雲林縣" then%>
			myForm.BillFillDate.value="";
			bDealLineDate=""
		<%end if%>
		DLDate2.innerHTML="<input type='text' size='6' value='"+bDealLineDate+"' maxlength='7' name='DealLineDate' onBlur='DealLineDateReplace()' onkeydown='funTextControl(this);' style=ime-mode:disabled <%if sys_City="高雄市" Or sys_City=ApconfigureCityName then response.write "onkeyup=FullToGoNextTag(6,'DealLineDate','BillMem1');"%> <%
				if sys_City="基隆市" or sys_City="花蓮縣" then '到案日不可修改
					response.write " readonly"
				End if%>><%
				if sys_City="基隆市" then '到案日不可修改
					response.write " <div id='LayerGL578' style='position:absolute; width:205px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;'><span class='style10'>因審計室審查，到案日不可修改</span></div>"
				End if%><%
				if sys_City="花蓮縣" then
					response.write "<input type='checkbox' name='chkbDealLineDate' value='1' onclick='getDealLineDate();'"
					if bchkbDealLineDate="30" then
						response.write "checked"
					end if 
					response.write ">30天"
				end if 
				%>";
		DLDate4.innerHTML="<input type='radio' name='IsMail' value='1' <% 
		If sys_City="澎湖縣" or sys_City="保二總隊三大隊一中隊" Then
				response.write "onclick='getDealLineDate();' " 
		End If

		if bEquipMent<>"-1" or isnull(bEquipMent) then
			response.write "checked"
		end if
		%>>是<input type='radio' name='IsMail' value='-1' <%
		If sys_City="澎湖縣" or sys_City="保二總隊三大隊一中隊" Then
				response.write "onclick='getDealLineDate();' "
		End If
		if bEquipMent="-1" then
			response.write "checked"
		end if
		%>>否";
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="保二總隊三大隊一中隊" then%>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="基隆市" or sys_City="苗栗縣" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="新竹市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="花蓮縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMemNameQry||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>

	}else{
<%if sys_City<>"屏東縣" then %>
		myForm.Billno1.value="";
		myForm.Billno1.disabled=true;
		if (myForm.UseTool.value=="8"){
			myForm.UseTool.value="";
		}
		//LayerDLDate.style.visibility = "hidden"; 
		//LayerMStation.style.visibility = "hidden"; 
		//myForm.MemberStation.disabled=true;
		myForm.MemberStation.Type="Text";
		DLDate1.innerHTML="";
		DLDate2.innerHTML="<input type='hidden' size='6' value='"+bDealLineDate+"' maxlength='7' name='DealLineDate' onBlur='DealLineDateReplace()' style=ime-mode:disabled >";
		DLDate3.innerHTML="";
		DLDate4.innerHTML="<input type='hidden' size='6' value='1' maxlength='7' name='IsMail' style=ime-mode:disabled>";
		getDealLineDate();
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="保二總隊三大隊一中隊" then%>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="基隆市" or sys_City="苗栗縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="新竹市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,JurgeDay||DriverPID,MemberStation||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="花蓮縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then %>
	MoveTextVar("CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMemNameQry||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
<%else%>
		myForm.Billno1.value="";
		myForm.Billno1.disabled=true;
		if (myForm.UseTool.value=="8"){
			myForm.UseTool.value="";
		}
		DLDate1.innerHTML="應到案日期";
		DLDate3.innerHTML="是否郵寄";
		DLDate2.innerHTML="<input type='text' size='6' value='"+bDealLineDate+"' maxlength='7' name='DealLineDate' onBlur='DealLineDateReplace()' onkeydown='funTextControl(this);' style=ime-mode:disabled  >";
		DLDate4.innerHTML="<input type='radio' name='IsMail' value='1' checked>是";
		getDealLineDate();

		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%end if%>
	}
}
function DealLineDateReplace(){
	myForm.DealLineDate.value=myForm.DealLineDate.value.replace(/[^\d]/g,'');

}


//逕舉建檔清冊
function funPrintCaseList_Report(){
	UrlStr="../Query/PrintCaseDataList_Report.asp?CallType=1";
	newWin(UrlStr,"CaseListWin2342",980,575,0,0,"yes","yes","yes","no");
}

function KeyDown(){ 
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if (event.keyCode==116){	//F5查詢
		event.keyCode=0;   
		event.returnValue=false;  
		funcOpenBillQry();
	}else if (event.keyCode==117){ //F6鎖死
		event.keyCode=0;   
		event.returnValue=false;  
<%else%>
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		event.returnValue=false;  
		funcOpenBillQry();
<%end if%>
<%if sys_City="台東縣" or sys_City="高雄縣" then%>
	}else if (event.keyCode==120){ //台東縣F9存檔
		event.keyCode=0;   
		InsertBillVase();
<%else%>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
<%end if%>
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		event.returnValue=false; 
		location='BillKeyIn_Car_Report.asp'
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
<%if sys_City="高雄縣" then%>
	}else if (event.keyCode==113){ //高雄縣F2查詢未建檔
		event.keyCode=0;   
		funPrintCaseList_Report();
<%else%>
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		funPrintCaseList_Report();
<%end if%>
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		location='BillKeyIn_Report_Back.asp?PageType=Back'
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		location='BillKeyIn_Report_Back.asp?PageType=First'
	}
}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=2;
<%if sys_City="台中市" then%>
	window.open("EasyBillQry_TC.asp","WebPage86_Update","left=300,top=200,location=0,width=350,height=160,resizable=yes,scrollbars=yes");
<%else%>
	window.open("EasyBillQry.asp","WebPage86_Update","left=300,top=200,location=0,width=350,height=160,resizable=yes,scrollbars=yes");
<%end if %>
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
//附加說明
function Add_LawPlus(){
	if (myForm.Rule1.value==""){
		alert("請先輸入違規法條一!!");
	}else{
	RuleID=myForm.Rule1.value;
	window.open("Query_LawPlus.asp?RuleID="+RuleID+"&theRuleVer=<%=theRuleVer%>","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
	}
}
function OpenSpecCar(){
	window.open("/traffic/SpecCar/SpecCar.asp","WebPage1","left=20,top=10,location=0,width=900,height=655,resizable=yes,scrollbars=yes");
}
function funGetSpeedRule(){
	<%if sys_City="基隆市" then%>
	var reqIllgealAressID="<%=trim(request("IllegalAddressID"))%>";
	if (myForm.IllegalAddressID.value=="RA743" || myForm.IllegalAddressID.value=="RA744" || myForm.IllegalAddressID.value=="RA694" || myForm.IllegalAddressID.value=="RA695" || myForm.IllegalAddressID.value=="RA696" || myForm.IllegalAddressID.value=="RA697" || myForm.IllegalAddressID.value=="RA746" || myForm.IllegalAddressID.value=="RA717" || myForm.IllegalAddressID.value=="RA734" || myForm.IllegalAddressID.value=="RA763" || myForm.IllegalAddressID.value=="RA781" || myForm.IllegalAddressID.value=="RA759" || myForm.IllegalAddressID.value=="RA762")
	{
		myForm.chkHighRoad.checked=true;
	}else if (reqIllgealAressID=="RA743" || reqIllgealAressID=="RA744" || reqIllgealAressID=="RA694" || reqIllgealAressID=="RA695" || reqIllgealAressID=="RA696" || reqIllgealAressID=="RA697" || reqIllgealAressID=="RA746" || reqIllgealAressID=="RA717" || reqIllgealAressID=="RA734" || reqIllgealAressID=="RA763" || reqIllgealAressID=="RA781" || reqIllgealAressID=="RA759" || reqIllgealAressID=="RA762")
	{
		if (myForm.IllegalAddressID.value!="RA743" && myForm.IllegalAddressID.value!="RA744" && myForm.IllegalAddressID.value!="RA694" && myForm.IllegalAddressID.value!="RA695" && myForm.IllegalAddressID.value!="RA696" && myForm.IllegalAddressID.value!="RA697" && myForm.IllegalAddressID.value!="RA746" && myForm.IllegalAddressID.value!="RA717" && myForm.IllegalAddressID.value!="RA734" && myForm.IllegalAddressID.value!="RA763" && myForm.IllegalAddressID.value!="RA781" && myForm.IllegalAddressID.value!="RA759" && myForm.IllegalAddressID.value!="RA762")
		{
			myForm.chkHighRoad.checked=false;
		}
	}
	<%end if %>
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
//用地點、車速抓違規法條
function setIllegalRule(){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
		if ((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
			if (myForm.IllegalDate.value>="1120630"){
		<%if sys_City="台中市" then%>
			IllegalRule=getIllegalRule2(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%elseif sys_City="台東縣" or sys_City="雲林縣" then%>
			IllegalRule=getIllegalRule3(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%else%>
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%end if%>
			}else{
		<%if sys_City="台中市" then%>
			IllegalRule=getIllegalRule2_Old1120630(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%elseif sys_City="台東縣" or sys_City="雲林縣" then%>
			IllegalRule=getIllegalRule3_Old1120630(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%else%>
			IllegalRule=getIllegalRule_Old1120630(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%end if%>
			}
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

//用車速，地點得到違規法條(台東縣 4310201 , 4000003)61以上才能開43條--違規日1120630帶舊法條
function getIllegalRule3_Old1120630(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;
//		if (Illaddr.indexOf("高速公路",0)!=-1){
//			if (Speed <= 20 && Speed > 0){
//				return "3310101";
//			}else if (Speed > 20 && Speed <= 40){
//				return "3310103";
//			}else if (Speed > 40 && Speed <= 60){
//				return "3310105";
//			}else if (Speed > 60 && Speed <= 80){
//				return "4310210";
//			}else if (Speed > 80 && Speed <= 100){
//				return "4310211";
//			}else if (Speed > 100){
//				return "4310212";
//			}else{
//				return "Null";
//			}
//		}else 
		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310102";
			}else if (Speed > 20 && Speed <= 40){
				return "3310104";
			}else if (Speed > 40 && Speed <= 60){
				return "3310106";
			}else if (Speed > 60 && Speed <= 80){
				return "4310210";
			}else if (Speed > 80 && Speed <= 100){
				return "4310211";
			}else if (Speed > 100){
				return "4310212";
			}else{
				return "Null";
			}
		}else{
			if (Speed <= 20 && Speed > 0){
				return "4000005";
			}else if (Speed > 20 && Speed <= 40){
				return "4000006";
			}else if (Speed > 40 && Speed <= 60){
				return "4000007";
			}else if (Speed > 60 && Speed <= 80){
				return "4310210";
			}else if (Speed > 80 && Speed <= 100){
				return "4310211";
			}else if (Speed > 100){
				return "4310212";
			}else{
				return "Null";
			}	
		}
	}
}

<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	//打滿跳下格
	function FullToGoNextTag(tagLen,tagName,togetTagName){
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
			if (tagName=="IllegalDate" || tagName=="DealLineDate" || tagName=="BillFillDate" || tagName=="JurgeDay")	{
				if (eval("myForm."+tagName).value.substr(0,1)=="1" ){
					if (eval("myForm."+tagName).value.length==7){
						eval("myForm."+togetTagName).select();
					}					
				}else if(eval("myForm."+tagName).value.length==6){
					eval("myForm."+togetTagName).select();
				}
				
			}else{
				if (eval("myForm."+tagName).value.length==tagLen){
					eval("myForm."+togetTagName).select();
				}
			}
		}
	}
<%end if%>
<%if sys_City="台中市" then%>
	function getAcceptData(){
		AcceptNo=myForm.ReportNo.value;
		event.keyCode=13;
		runServerScript("getAcceptData.asp?AcceptNo="+AcceptNo);
		
	}

function chkAcceptBatch(){
	if (myForm.ReportChk.checked==false)
			{
				if (myForm.AcceptBatchNumberChk.checked==true)
				{
					if (myForm.AcceptBatchNumber.value != "" && myForm.CarSimpleID.value != "")
					{
						runServerScript("getChkRunCarAccept.asp?AcceptBatchNumber="+myForm.AcceptBatchNumber.value+"&CarNo="+myForm.CarNo.value+"&CarSimpleID="+myForm.CarSimpleID.value);
					}	
				}
			}
}
<%end if%>
<%if sys_City="台南市" then%>
var sys_City="<%=sys_City%>";
function getDriverZip(obj,objName){
	if(obj.value!=''&&obj.value.length>2){
		if ((myForm.OldIllegalZip.value != "") && (myForm.OldIllegalZip.value != myForm.IllegalZip.value) && (myForm.IllegalAddressID.value == "")){
			myForm.IllegalAddress.value = "";
		}
		runServerScript("getZipNameForCar.asp?ZipID="+obj.value+"&getZipName="+objName+"&getIllegalAddress="+myForm.IllegalAddress.value);
	}else if(obj.value!=''&&obj.value.length<3){
		alert("郵遞區號錯誤!!");
		TDIllZipErrorLog=1;
	}
}
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");

}
<%elseif sys_City="高雄市" or sys_City="台中市" then%>
var sys_City="<%=sys_City%>";
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes,status=yes");

}
function getIllZip(){
	<%if sys_City="台中市" then%>
	if (event.keyCode==116){	
		event.keyCode=0;
		QryIllegalZip();
	}
	<%end if %>
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value);
}


<%end if %>
	//-----------上下左右-------------
	function funTextControl(obj){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeEnter(obj.name);
		}	
		/*if (event.keyCode==37){ //左換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}else */else if (event.keyCode==38){ //上換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}/*else if (event.keyCode==39){ //右換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveRight(obj.name);
		}*/else if (event.keyCode==40){ //下換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveRight(obj.name);
		}else if (event.keyCode==9){ 
			if (obj==myForm.IllegalAddress){
				event.keyCode=0;
				event.returnValue=false;
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then%>
				myForm.RuleSpeed.select();
<%elseif sys_City="嘉義市" then%>
			if (myForm.ReportChk.checked==false){
				myForm.BillMem1.select();
			}else{
				myForm.DealLineDate.select();
			}
<%else%>
				myForm.Rule1.select();
<%end if%>
			}
		}
	<%if sys_City="台南市" then%>

		if (obj.name=="IllegalZip"&&event.keyCode==116){	
			window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
	<%end if %>
<%if sys_City="雲林縣" then%>
		if (obj==myForm.BillFillDate){
			getDealLineDate();
		}
<%end if%>
	}
	//------------------------------
if (myForm.ReportChk.checked==false){
<%if sys_City="新竹市" then%>
	myForm.BillFillDate.focus();
<%elseif sys_City="台中市" then%>
	myForm.ReportNo.focus();
<%else%>
	myForm.CarNo.focus();
<%end if%>
}else{
	myForm.Billno1.focus();
}

<%if sys_City="苗栗縣" and trim(request("ReCoverSn"))<>"" then%>

<%else%>
<%	if trim(request("ReportCaseSn"))<>"" then%>
	myForm.ReportChk.checked=true;
<%	end if%>
funcReportChk();
<%end if%>

<%if sys_City<>"台中縣" then%>
getDealLineDate();
<%else%>
if (myForm.ReportChk.checked==false){
	getDealLineDate();
}
<%end if%>
<%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then
%>
if (myForm.isSave3.checked==true){
	myForm.Billno1.value="<%=left(trim(request("Billno1")),3)%>";
}
<%
			end if
%>
<%if sys_City="苗栗縣" then%>
function getBillData(){
	myForm.IllegalTime.value=myForm.IllegalTime.value.replace(/[^\d]/g,'');
	if ((myForm.IllegalTime.value.length==4) && (myForm.IllegalDate.value.length==6 || myForm.IllegalDate.value.length==7) && (myForm.CarNo.value!="")){
		//alert("getBillData.asp?CarNo="+myForm.CarNo.value+"&IllegalDate="+myForm.IllegalDate.value+"&IllegalTime="+myForm.IllegalTime.value);
		runServerScript("getBillData.asp?CarNo="+myForm.CarNo.value+"&IllegalDate="+myForm.IllegalDate.value+"&IllegalTime="+myForm.IllegalTime.value+"&nowTime=<%=now%>");
	}
	
}
<%end if%>
<%
	if sys_City="台中市" Then
	%>
	function getBillMemData(){
		if (event.keyCode==<%
			if sys_City="高雄市" Or sys_City=ApconfigureCityName then
				response.write "117"
			else
				response.write "116"
			end if
			%>){	
			event.keyCode=0;
			window.open("Query_MemID.asp?MemType=CarS&MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}else{
			if (myForm.BillMemNameQry.value.length > 1){
				var BillMemNum=myForm.BillMemNameQry.value;
				runServerScript("getBillMemData.asp?MemName="+BillMemNum);
			}else if (myForm.BillMemNameQry.value.length <= 1 && myForm.BillMemNameQry.value.length > 0){
				Layer50.innerHTML=" ";
				myForm.BillMemID1Qry.value="";
				myForm.BillUnitTypeID1Qry.value="";
				myForm.BillUnitIDQry.value="";
				myForm.BillUnitNameQry.value="";
			}else{
				Layer50.innerHTML=" ";
				myForm.BillMemID1Qry.value="";
				myForm.BillUnitTypeID1Qry.value="";
				myForm.BillUnitIDQry.value="";
				myForm.BillUnitNameQry.value="";
			}
		}
	}

	function InnerBillMemData(){
		if (myForm.BillMemID1Qry.value == "")
		{
			alert("請先查詢違規人姓名");
		}else{
			myForm.BillMem1.value=Layer50.innerHTML;
			Layer12.innerHTML=myForm.BillMemNameQry.value;
			myForm.BillMemID1.value=myForm.BillMemID1Qry.value;
			myForm.BillMemName1.value=myForm.BillMemNameQry.value;
			myForm.BillUnitTypeID1.value=myForm.BillUnitTypeID1Qry.value;
			myForm.BillUnitID.value=myForm.BillUnitIDQry.value;
			Layer6.innerHTML=myForm.BillUnitNameQry.value;
<%
	if sys_City="台中市" Then
	%>
	getUnit(1);
	<%
	End if
%>
		} 
	}
	<%
	End If 
	%>
<%
if sys_City="基隆市" Then
%>
	function changeLockIllegalAddress(){
		if (myForm.LockIllegalAddress.checked==true)
		{
			document.getElementById('IllegalAddress').readOnly = true;
		}else{
			document.getElementById('IllegalAddress').readOnly = false;
		}
	}
	
	changeLockIllegalAddress();
<%
End If 

If trim(request("kinds"))="" And Trim(request("ReportCaseSn"))<>"" Then
	if sys_City="雲林縣" Then
		if trim(bRule1 & "")<>"" then
			response.write "getRuleData1(1);"
		end If
		if trim(bRule2 & "")<>"" then
			response.write "getRuleData2(1);"
		end If
		if trim(bLoginID1 & "")<>"" then
			'response.write "getBillMemID1(1);"
		end if 
	end if 
end if 
%>
</script>
</html>
