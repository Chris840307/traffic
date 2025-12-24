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
	'smith remark
	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
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
'要抓本機照片或是server上的照片(0:本機 1:Server)or sys_City="苗栗縣" 
if sys_City="高雄市" Or sys_City=ApconfigureCityName Or sys_City="花蓮縣" then
	HowCatchPicture="0" 
else
	HowCatchPicture="1" 
end if
'本機路逕
if trim(request("ImageSaveLocation"))<>"" then
	Session("ImageSaveLocation")=trim(request("ImageSaveLocation"))
end if

if trim(Session("ImageSaveLocation"))="" Then
	If sys_City="花蓮縣" then
		UserPicturePath="C:/Image/ok/ok/"
	Else 
		UserPicturePath="C:/Image/ok/"
	End if
else
	UserPicturePath=trim(Session("ImageSaveLocation"))
end if
PicturePath="file:///"&UserPicturePath
'============================================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing

'併上案
if trim(request("kinds"))="BillMerge" then
	strM1="select * from ( " &_
		" select a.SN,b.FileName,b.ImageFileNameA,b.ImageFileNameB,b.ImageFileNameC,b.OperatorA " &_
		" from BillBase a,ProsecutionImage b,ProsecutionImageDetail c " &_
		" where a.BillTypeID='2' and a.BillStatus in ('0') and a.RecordStateID=0 " &_
		" and a.RecordMemberID="&theRecordMemberID&" and a.SN=c.BillSN " &_
		" and b.OperatorA=c.Operator " &_
		" and c.FileName=b.FileName and b.FixEquipType in (1,2,5,8,10) order by a.sn desc " &_
		" ) where rownum<=1 "
	set rsM1=conn.execute(strM1)
	if not rsM1.eof then
		if trim(rsM1("ImageFileNameA"))<>"" and trim(rsM1("ImageFileNameB"))<>"" and trim(rsM1("ImageFileNameC"))<>"" then
%>
<script language="JavaScript">
	alert("上筆資料已有三張照片!!");
</script>
<%
		else
			if trim(rsM1("ImageFileNameB"))="" or isnull(rsM1("ImageFileNameB")) then
				strM2="Update ProsecutionImage set ImageFileNameB='"&trim(request("gImageFileNameA"))&"'" &_
					" where FileName='"&trim(rsM1("FileName"))&"' and OperatorA='"&trim(rsM1("OperatorA"))&"'"
				conn.execute strM2
				strM3="Update BILLILLEGALIMAGE set ImageFileNameB='"&trim(request("gImageFileNameA"))&"'" &_
					" where Billsn="&trim(rsM1("SN"))&""
				conn.execute strM3
			else
				strM2="Update ProsecutionImage set ImageFileNameC='"&trim(request("gImageFileNameA"))&"'" &_
					" where FileName='"&trim(rsM1("FileName"))&"' and OperatorA='"&trim(rsM1("OperatorA"))&"'"
				conn.execute strM2
				strM3="Update BILLILLEGALIMAGE set ImageFileNameC='"&trim(request("gImageFileNameA"))&"'" &_
					" where Billsn="&trim(rsM1("SN"))&""
				conn.execute strM3
			end if
			'strUpdate2="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1,REALCARNO='"&UCase(trim(request("CarNo")))&"' where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
			strUpdate2="delete from PIDetail where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
			Conn.execute strUpdate2

			strUpdate3="delete from PI where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
			Conn.execute strUpdate3
		end if
	else
%>
<script language="JavaScript">
	alert("查無上筆資料!!");
</script>
<%
	end if
	rsM1.close
	set rsM1=nothing

end if

'新增告發單
if trim(request("kinds"))="DB_insert" Then
	If Trim(request("RuleSpeed"))<>"" And Trim(request("IllegalSpeed"))<>"" Then
		If Trim(request("RuleSpeed"))>300 Or Trim(request("IllegalSpeed"))>300 Then
			chkIsSpeedTooOver=1
		Else
			chkIsSpeedTooOver=0
		End If 
	Else
		chkIsSpeedTooOver=0
	End If 
	
	checkReportCaseFlag=0

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
		
		If Trim(request("ReportCaseNo"))<>"" then
			'高雄市將案件帶入民眾檢舉系統
			strchkKR="select CarNo,BillStatus,Billsn from BillbaseTmp where ReportCaseNo='"&Trim(request("ReportCaseNo"))&"' and recordstateid=0"
			Set rschkKR=conn.execute(strchkKR)
			If Not rschkKR.eof Then
				If Trim(rschkKR("BillStatus"))<>"1" Or Trim(rschkKR("Billsn") & "")<>"" Then
					checkReportCaseFlag=1
					chkAlertString=chkAlertString&"\n儲存失敗，此局信箱編號("&Trim(request("ReportCaseNo"))&")已經結案。"
				End If 
'				If UCase(Trim(rschkKR("CarNo")))<>UCase(Trim(request("CarNo"))) Then
'					checkReportCaseFlag=1
'					chkAlertString=chkAlertString&"\n儲存失敗，輸入車號("&Trim(request("CarNo"))&")與民眾檢舉系統車號("&Trim(rschkKR("CarNo"))&")不符。"
'				End If 
			Else 
				checkReportCaseFlag=1
				chkAlertString=chkAlertString&"\n儲存失敗，查無此局信箱編號("&Trim(request("ReportCaseNo"))&")。"
			End If 
			rschkKR.close
			Set rschkKR=Nothing 
			
		End If 
	End If

If chkIsSpeedTooOver=0 And checkReportCaseFlag=0 then
	'違規日期
	theIllegalDate=""
	if trim(request("IllegalDate"))<>"" then
		theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
	else
		theIllegalDate = "null"
	end if	

	
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
	if trim(request("UseTool"))="" then
		theUseTool=0
	else
		theUseTool=trim(request("UseTool"))
	end if
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
	'民眾檢舉時間
	theJurgeDay=""
	if trim(request("JurgeDay"))<>"" then
		theJurgeDay=DateFormatChange(trim(request("JurgeDay")))
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
		theCarAddID="null"
	else
		theCarAddID=trim(request("CarAddID"))
	end if
	'建檔影像
		theImageFileName=trim(request("gImageFileNameA"))
		theImagePathName=trim(request("gImagePathNameA"))
	'SN抓最大值
	sSQL = "select BillBase_seq.nextval as SN from Dual"
	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		sMaxSN = oRST("SN")
	end if
	oRST.close
	set oRST = nothing
If sys_City="彰化縣" Then
	'彰化民眾檢舉要先寫入
	'BillBaseTMP
	strInsert="insert into BillBaseTmp(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID4,BillMem4,BillMemID2,BillMem2,BillMemID3,BillMem3" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,ImageFileName"&ColAdd&",JurgeDay,CheckFlag)" &_
				" values("&sMaxSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
				",'"&UCase(trim(request("CarNo")))&"',"&trim(request("CarSimpleID")) &_						          
				","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
				",'"&trim(request("IllegalAddress"))&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
				","&theRuleSpeed&","&trim(request("ForFeit1"))&",'"&trim(request("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
				","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
				",'"&UCase(trim(request("DriverPID")))&"',"& theDriverBirth &",'"&trim(request("DriverName"))&"'" &_
				",'"&trim(request("DriverAddress"))&"','"&trim(request("DriverZip"))&"','"&trim(request("MemberStation"))&"'" &_
				",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
				",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
				",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
				",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'1',0,SYSDate,'" & theRecordMemberID &"'" &_
				",'"&trim(request("Note"))&"','1','"&theRuleVer&"'" &_
				",'"&trim(request("DriverSex"))&"','"&trim(theImageFileName)&"'" &_
				""&valueAdd&"," & theJurgeDay & ",'0')"
				conn.execute strInsert  

	'寫入BILLILLEGALIMAGE
	if trim(request("PicCount"))="1" then
		strBillImage="Insert Into BILLILLEGALIMAGETemp2(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'','','"&trim(theImagePathName)&"')"
	elseif trim(request("PicCount"))="2" then
		Tmp_gFileName=""
		Tmp_gImageFileName=""
		Tmp_gOperator=""
		if trim(request("SelectImage"))="1" then
			Tmp_gFileName=trim(request("gFileName1"))
			Tmp_gImageFileName=trim(request("gImageFileNameB"))
			Tmp_gOperator=trim(request("gOperator1"))
		else
			Tmp_gFileName=trim(request("gFileName2"))
			Tmp_gImageFileName=trim(request("gImageFileNameC"))
			Tmp_gOperator=trim(request("gOperator2"))
		end if
		strBillImage="Insert Into BILLILLEGALIMAGETemp2(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'"&Tmp_gImageFileName&"','','"&trim(theImagePathName)&"')"

		'strdel1="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&Tmp_gFileName&"' and Operator='" & Tmp_gOperator & "'"
		strdel1="delete from PIDetail where FileName='"&Tmp_gFileName&"' and Operator='" & Tmp_gOperator & "'"
		Conn.execute strdel1
		
		strdel1b="delete from PI where FileName='"&Tmp_gFileName&"' and OperatorA='" & Tmp_gOperator & "'"
		Conn.execute strdel1b

		strdel1B="Update PI set ImageFileNameB='"&trim(request("gImageFileNameB"))&"' where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
		Conn.execute strdel1B
	elseif trim(request("PicCount"))="3" then
		strBillImage="Insert Into BILLILLEGALIMAGETemp2(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'"&trim(request("gImageFileNameB"))&"','"&trim(request("gImageFileNameC"))&"','"&trim(theImagePathName)&"')"

		'strdel1="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&request("gFileName1")&"' and Operator='" & request("gOperator1") & "'"
		strdel1="Delete from PIDetail where FileName='"&request("gFileName1")&"' and Operator='" & request("gOperator1") & "'"
		Conn.execute strdel1
		
		strdel1b="Delete from PI where FileName='"&request("gFileName1")&"' and OperatorA='" & request("gOperator1") & "'"
		Conn.execute strdel1b

		'strdel2="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&request("gFileName2")&"' and Operator='" & request("gOperator2") & "'"
		strdel2="Delete from PIDetail where FileName='"&request("gFileName2")&"' and Operator='" & request("gOperator2") & "'"
		Conn.execute strdel2

		strdel2b="Delete from PI where FileName='"&request("gFileName2")&"' and OperatorA='" & request("gOperator2") & "'"
		Conn.execute strdel2b

		strdel1B="Update PI set ImageFileNameB='"&trim(request("gImageFileNameB"))&"',ImageFileNameC='"&trim(request("gImageFileNameC"))&"' where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
		Conn.execute strdel1B

		
	end if
	conn.execute strBillImage  
Else
	'BillBase
	If sys_City="高雄市" Then
		ColAdd=",IllegalZip"
		valueAdd=",'"&trim(request("IllegalZip"))&"'"
	End if	
	strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID4,BillMem4,BillMemID2,BillMem2,BillMemID3,BillMem3" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,ImageFileName"&ColAdd&",JurgeDay)" &_
				" values("&sMaxSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
				",'"&UCase(trim(request("CarNo")))&"',"&trim(request("CarSimpleID")) &_						          
				","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
				",'"&trim(request("IllegalAddress"))&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
				","&theRuleSpeed&","&trim(request("ForFeit1"))&",'"&trim(request("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
				","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
				",'"&UCase(trim(request("DriverPID")))&"',"& theDriverBirth &",'"&trim(request("DriverName"))&"'" &_
				",'"&trim(request("DriverAddress"))&"','"&trim(request("DriverZip"))&"','"&trim(request("MemberStation"))&"'" &_
				",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
				",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
				",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
				",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'0',0,SYSDate,'" & theRecordMemberID &"'" &_
				",'"&trim(request("Note"))&"','1','"&theRuleVer&"'" &_
				",'"&trim(request("DriverSex"))&"','"&trim(theImageFileName)&"'" &_
				""&valueAdd&"," & theJurgeDay & ")"
				conn.execute strInsert  

	'寫入BILLILLEGALIMAGE
	if trim(request("PicCount"))="1" then
		strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'','','"&trim(theImagePathName)&"')"
	elseif trim(request("PicCount"))="2" then
		Tmp_gFileName=""
		Tmp_gImageFileName=""
		Tmp_gOperator=""
		if trim(request("SelectImage"))="1" then
			Tmp_gFileName=trim(request("gFileName1"))
			Tmp_gImageFileName=trim(request("gImageFileNameB"))
			Tmp_gOperator=trim(request("gOperator1"))
		else
			Tmp_gFileName=trim(request("gFileName2"))
			Tmp_gImageFileName=trim(request("gImageFileNameC"))
			Tmp_gOperator=trim(request("gOperator2"))
		end if
		strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'"&Tmp_gImageFileName&"','','"&trim(theImagePathName)&"')"

		'strdel1="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&Tmp_gFileName&"' and Operator='" & Tmp_gOperator & "'"
		strdel1="delete from PIDetail where FileName='"&Tmp_gFileName&"' and Operator='" & Tmp_gOperator & "'"
		Conn.execute strdel1
		
		strdel1b="delete from PI where FileName='"&Tmp_gFileName&"' and OperatorA='" & Tmp_gOperator & "'"
		Conn.execute strdel1b

		strdel1B="Update PI set ImageFileNameB='"&trim(request("gImageFileNameB"))&"' where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
		Conn.execute strdel1B
	elseif trim(request("PicCount"))="3" then
		strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(request("Billno1")))&"','"&trim(theImageFileName)&"'" &_
		",'"&trim(request("gImageFileNameB"))&"','"&trim(request("gImageFileNameC"))&"','"&trim(theImagePathName)&"')"

		'strdel1="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&request("gFileName1")&"' and Operator='" & request("gOperator1") & "'"
		strdel1="Delete from PIDetail where FileName='"&request("gFileName1")&"' and Operator='" & request("gOperator1") & "'"
		Conn.execute strdel1
		
		strdel1b="Delete from PI where FileName='"&request("gFileName1")&"' and OperatorA='" & request("gOperator1") & "'"
		Conn.execute strdel1b

		'strdel2="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where FileName='"&request("gFileName2")&"' and Operator='" & request("gOperator2") & "'"
		strdel2="Delete from PIDetail where FileName='"&request("gFileName2")&"' and Operator='" & request("gOperator2") & "'"
		Conn.execute strdel2

		strdel2b="Delete from PI where FileName='"&request("gFileName2")&"' and OperatorA='" & request("gOperator2") & "'"
		Conn.execute strdel2b

		strdel1B="Update PI set ImageFileNameB='"&trim(request("gImageFileNameB"))&"',ImageFileNameC='"&trim(request("gImageFileNameC"))&"' where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
		Conn.execute strdel1B

		
	end if
	conn.execute strBillImage  
End if
	

	strPI1="select * from PI where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
	Set rsPI1=conn.execute(strPI1)
	If Not rsPI1.eof Then
		If IsNull(rsPI1("PROSECUTIONTIME")) then
			sysPROSECUTIONTIME="null"
		Else
			sysPROSECUTIONTIME="to_date('"&Year(rsPI1("PROSECUTIONTIME"))&"/"&month(rsPI1("PROSECUTIONTIME"))&"/"&day(rsPI1("PROSECUTIONTIME"))&" "&Hour(rsPI1("PROSECUTIONTIME"))&":"&Minute(rsPI1("PROSECUTIONTIME"))&":"&Second(rsPI1("PROSECUTIONTIME"))&"','YYYY/MM/DD/HH24/MI/SS')"
		End If 
		If IsNull(rsPI1("LIMITSPEED")) Then
			sysLIMITSPEED="null"
		Else
			sysLIMITSPEED=trim(rsPI1("LIMITSPEED"))
		End If 
		If IsNull(rsPI1("TRIGGERSPEED")) Then
			sysTRIGGERSPEED="null"
		Else
			sysTRIGGERSPEED=trim(rsPI1("TRIGGERSPEED"))
		End If 
		If IsNull(rsPI1("REPORTLINEA")) Then
			sysREPORTLINEA="null"
		Else
			sysREPORTLINEA=trim(rsPI1("REPORTLINEA"))
		End If 
		If IsNull(rsPI1("REPORTLINEB")) Then
			sysREPORTLINEB="null"
		Else
			sysREPORTLINEB=trim(rsPI1("REPORTLINEB"))
		End If 
		If IsNull(rsPI1("OVERSPEED")) Then
			sysOVERSPEED="null"
		Else
			sysOVERSPEED=trim(rsPI1("OVERSPEED"))
		End If 
		If IsNull(rsPI1("POSITION")) Then
			sysPOSITION="null"
		Else
			sysPOSITION=trim(rsPI1("POSITION"))
		End If 
		If IsNull(rsPI1("AMBERTIME")) Then
			sysAMBERTIME="null"
		Else
			sysAMBERTIME=trim(rsPI1("AMBERTIME"))
		End If 
		If IsNull(rsPI1("REDLIGHTTIME")) Then
			sysREDLIGHTTIME="null"
		Else
			sysREDLIGHTTIME=trim(rsPI1("REDLIGHTTIME"))
		End If 
		If IsNull(rsPI1("INTERVALTIME")) Then
			sysINTERVALTIME="null"
		Else
			sysINTERVALTIME=trim(rsPI1("INTERVALTIME"))
		End If 
		If IsNull(rsPI1("LINE")) Then
			sysLINE="null"
		Else
			sysLINE=trim(rsPI1("LINE"))
		End If 
		strPIadd="insert into ProsecutionImage(FileName,DIRECTORYNAME,FIXEQUIPTYPE,SITECODE,PROSECUTIONTIME" &_
			",PROSECUTIONTYPEID,LOGFILE,DISKID,SITENAME,LOCATION,DISTRICT,OPERATORA,OPERATORB,LIMITSPEED" &_
			",TRIGGERSPEED,REPORTLINEA,REPORTLINEB,RADARID,OVERSPEED,DIRECTION,POSITION,AMBERTIME,REDLIGHTTIME" &_
			",INTERVALTIME,LINE,REJECTCODE,REJECTREASON,VIDEOFILENAME,IMAGEFILENAMEA,IMAGEFILENAMEB" &_
			",CARDISTANCE,IMAGEFILENAMEC" &_
			") values('"&Trim(rsPI1("FileName"))&"','"&Trim(rsPI1("DIRECTORYNAME"))&"'" &_
			","&Trim(rsPI1("FIXEQUIPTYPE"))&",'"&Trim(rsPI1("SITECODE"))&"',"&sysPROSECUTIONTIME &_
			",'"&Trim(rsPI1("PROSECUTIONTYPEID"))&"','"&Trim(rsPI1("LOGFILE"))&"','"&Trim(rsPI1("DISKID"))&"'" &_
			",'"&Trim(rsPI1("SITENAME"))&"','"&Trim(rsPI1("LOCATION"))&"','"&Trim(rsPI1("DISTRICT"))&"'" &_
			",'"&Trim(rsPI1("OPERATORA"))&"','"&Trim(rsPI1("OPERATORB"))&"',"&sysLIMITSPEED &_
			","&sysTRIGGERSPEED&","&sysREPORTLINEA&","&sysREPORTLINEB&",'"&Trim(rsPI1("RADARID"))&"'" &_
			","&sysOVERSPEED&",'"&Trim(rsPI1("DIRECTION"))&"',"&sysPOSITION&","&sysAMBERTIME &_
			","&sysREDLIGHTTIME&","&sysINTERVALTIME&","&sysLINE&",'"&Trim(rsPI1("REJECTCODE"))&"'" &_
			",'"&Trim(rsPI1("REJECTREASON"))&"','"&Trim(rsPI1("VIDEOFILENAME"))&"'" &_
			",'"&Trim(rsPI1("IMAGEFILENAMEA"))&"','"&Trim(rsPI1("IMAGEFILENAMEB"))&"'" &_
			",'"&Trim(rsPI1("CARDISTANCE"))&"','"&Trim(rsPI1("IMAGEFILENAMEC"))&"'" &_
			")"
		'response.write strPIadd
		conn.execute strPIadd
	End If
	rsPI1.close
	Set rsPI1=Nothing 
	strPID1="select * from PIDetail where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
	Set rsPID1=conn.execute(strPID1)
	If Not rsPID1.eof Then	
		If IsNull(rsPID1("CARSIMPLEID")) Then
			sysCARSIMPLEID="null"
		Else
			sysCARSIMPLEID=trim(rsPID1("CARSIMPLEID"))
		End If 
		If IsNull(rsPID1("CARADDID")) Then
			sysCARADDID="null"
		Else
			sysCARADDID=trim(rsPID1("CARADDID"))
		End If 

		strPIDadd="insert into ProsecutionImageDetail(FILENAME,SN,CARNO,REALCARNO,CARSIMPLEID,LAWITEMID" &_
			",LPRRESULTID,VERIFYRESULTID,MEMBERID,NOTE,CARADDID,BILLSN,OPERATOR)" &_
			" values('"&Trim(rsPID1("FILENAME"))&"',(select nvl(max(SN),0)+1 from ProsecutionImageDetail),'"&Trim(rsPID1("CARNO"))&"'" &_
			",'"&UCase(trim(request("CarNo")))&"',"&sysCARSIMPLEID&",'"&Trim(rsPID1("LAWITEMID"))&"'" &_
			",'"&Trim(rsPID1("LPRRESULTID"))&"',0,"&theRecordMemberID&",'"&Trim(rsPID1("NOTE"))&"'" &_
			","&sysCARADDID&","&sMaxSN&",'"&Trim(rsPID1("OPERATOR"))&"'" &_
			")"
		Conn.execute strPIDadd		
	End If
	rsPID1.close
	Set rsPID1=Nothing 

	strUpdate2="delete from PIDetail where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
	Conn.execute strUpdate2
	strUpdate2b="delete from PI where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
	Conn.execute strUpdate2b
	
	If sys_City="高雄市" Then
		If Trim(request("ReportCaseNo"))<>"" and checkReportCaseFlag=0 Then
			strKR="Update BillBaseTmp set BillStatus='8',Billsn="&sMaxSN&" where ReportCaseNo='"&Trim(request("ReportCaseNo"))&"'"
			'response.write strKR
			conn.execute strKR
		End If 		
	End If 

	'檢舉案件檢查一周內是否有違規
		if trim(request("JurgeDay"))<>"" Then
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
			illegalDate1=DateAdd("d",-7,illegalDateTmp)
			illegalDate2=DateAdd("d",7,illegalDateTmp)
			strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" 0:0:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

			'If (left(trim(request("Rule1")),2)="56" And Len(trim(request("Rule1")))=7) Or (left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) then
				strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate,JurgeDay" &_
					" from Billbase where sn<>"&sMaxSN &_
					" and carno='"&UCase(trim(request("CarNo")))&"'" &_
					" and Recordstateid=0 " & strIllDate & " and JurgeDay is not null"
				'response.write strChk
				Set rsChk=conn.execute(strChk)
				If Not rsChk.eof Then	
	%>
		<script language="JavaScript">
			window.open("JurgeCaseAlert.asp?BillSn=<%=sMaxSN%>","JurgeCaseAlert","left=100,top=20,location=0,width=700,height=555,resizable=yes,scrollbars=yes")
		</script>
	<%		
				End If 
				rsChk.close
				Set rsChk=Nothing 
			'End If 
		End If
%>
<script language="JavaScript">
<%
	'交通隊劉小姐要求超過60公里要跳提示1030516
	if sys_City="南投縣" then
		if trim(request("IllegalSpeed"))<>"" and trim(request("RuleSpeed"))<>"" then
			if cdbl(request("IllegalSpeed"))-cdbl(request("RuleSpeed"))>60 then
				response.write "alert('超過最高限速60公里以上需另單舉發法條4340003(處車主)!!');"
			end if 
		end if 
	end if 
%>
</script>
<%
ElseIf chkIsSpeedTooOver=1 then
	%>
	<script language="JavaScript">
		alert("限速或實速超過300Km，請確認是否正確！！");
	</script>
	<%
End If
	If chkIllegalDateAndCar_KS=1 Or checkReportCaseFlag=1 Then
%>
	<script language="JavaScript">
		alert("<%=chkAlertString%>");
	</script>
<%
	End If 
end if
'無效
if trim(request("kinds"))="VerifyResultNull" then
	'strUpdate2="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1,REALCARNO='"&UCase(trim(request("CarNo")))&"' where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
	strUpdate2="delete from PIDetail where FileName='"&request("SelFileName")&"' and SN='" & request("SelSN") & "'"
	Conn.execute strUpdate2
	strUpdate2b="delete from PI where FileName='"&request("SelFileName")&"' and OperatorA='" & trim(request("SelOperator")) & "'"
	Conn.execute strUpdate2b
end if
'所有未建檔都設無效
if trim(request("kinds"))="AllNotKeyInVerifyResultNull" then
	'strUpdate2="update ProsecutionImageDetail set VERIFYRESULTID=-1 where Operator='"&trim(Session("Credit_ID"))&"' and VERIFYRESULTID=1 and billsn is null"
	strUpdate2="delete from PIDetail where Operator='"&trim(Session("Credit_ID"))&"' "
	Conn.execute strUpdate2
	strUpdate2b="delete from PI where OperatorA='" & trim(Session("Credit_ID")) & "'"
	Conn.execute strUpdate2b
end if

if trim(request("SessionFlag"))="" then
	Session.Contents.Remove("BillIgnore_Fix")
end if
'略過
if trim(request("kinds"))="BillIgnore" then
	if trim(request("SelFileName"))<>"" then
		if session("BillIgnore_Fix")<>"" then
			session("BillIgnore_Fix")=session("BillIgnore_Fix")&",'"&request("SelFileName")&"'"
		else
			session("BillIgnore_Fix")="'"&request("SelFileName")&"'"
		end if
	end if
end if
if session("BillIgnore_Fix")<>"" then
	strIgnorePlus=" and a.FileName not in ("&session("BillIgnore_Fix")&")"
	strIgnorePlus2=" and ProsecutionImage.FileName not in ("&session("BillIgnore_Fix")&")"
else
	strIgnorePlus=""
	strIgnorePlus2=""
end if
	
		
	If sys_City="屏東縣" Then
		strOrder=" order by  FIXEQUIPTYPE desc,DirectoryName,FileName,Location,PROSECUTIONTIME desc "
	Else
		strOrder=" order by  FIXEQUIPTYPE desc,FileName,Location,PROSECUTIONTIME desc "
	End If 

	strSQL="select * from ( select b.CarNo,b.CarSimpleID,a.ProsecutionTime,a.ProsecutionTypeID,a.SiteCode,a.FileName,a.DirectoryName,a.FIXEQUIPTYPE,a.IMAGEFILENAMEA,a.VideoFileName,a.IMAGEFILENAMEB,a .IMAGEFILENAMEC,a.Location,b.LawItemID,b.SN,a.LIMITSPEED,a.OVERSPEED,a.OperatorA,b.MemberID,b.Note from PI a, PIDetail b where a.FILENAME = b.FILENAME and a.OperatorA=b.Operator and b.Operator='"&trim(Session("Credit_ID"))&"' and FixEquipType in (1,2,5,8,10) and VERIFYRESULTID<>-1 and BillSn is null and (a.RejectCode<>'262' or a.RejectCode is null) "&strIgnorePlus & strOrder & ") where rownum<=1"
Session.Contents.Remove("BillTime_Image")

strTime="select sysdate from dual"
set rsTime=conn.execute(strTime)
if not rsTime.eof then
	BillTime_ImageTmp=DateAdd("s" , 1,rsTime("sysdate"))
else
	BillTime_ImageTmp=DateAdd("s" , 1, now)
end if
Session("BillTime_Image")=date&" "&hour(BillTime_ImageTmp)&":"&minute(BillTime_ImageTmp)&":"&second(BillTime_ImageTmp)
'response.write strSQL

'總共幾筆
if trim(request("Tmp_Order"))="" then
	Session.Contents.Remove("BillCnt_Image")
	Session.Contents.Remove("BillOrder_Image")
	
	'strSqlCnt="select count(*) as cnt from BillBase a,ProsecutionImage b,ProsecutionImageDetail c where a.BillTypeID='2' and a.BillStatus in ('0') and a.RecordStateID=0 and a.RecordMemberID="&theRecordMemberID&" and a.SN=c.BillSN and c.FileName=b.FileName and b.OperatorA=c.Operator and b.FixEquipType in (1,2,5,8,10) "
	If sys_City="彰化縣" Then
		strSqlCnt="select count(*) as cnt from BillBaseTmp a where a.BillTypeID='2' and a.BillStatus in ('1') and a.RecordStateID=0 and a.RecordMemberID="&theRecordMemberID&" and ImageFileName is not null "
		set rsCnt1=conn.execute(strSqlCnt)
			Session("BillCnt_Image")=trim(rsCnt1("cnt"))
			Session("BillOrder_Image")=trim(rsCnt1("cnt"))+1
		rsCnt1.close
		set rsCnt1=nothing
	Else
		strSqlCnt="select count(*) as cnt from BillBase a where a.BillTypeID='2' and a.BillStatus in ('0') and a.RecordStateID=0 and a.RecordMemberID="&theRecordMemberID&" and ImageFileName is not null "
		set rsCnt1=conn.execute(strSqlCnt)
			Session("BillCnt_Image")=trim(rsCnt1("cnt"))
			Session("BillOrder_Image")=trim(rsCnt1("cnt"))+1
		rsCnt1.close
		set rsCnt1=nothing
	End If 

	
else
	if trim(request("kinds"))="DB_insert" then
		Session("BillCnt_Image")=trim(request("Tmp_Order"))+1
		Session("BillOrder_Image")=Session("BillCnt_Image")+1
	else
		Session("BillCnt_Image")=trim(request("Tmp_Order"))
		Session("BillOrder_Image")=Session("BillCnt_Image")+1
	end if
end if

	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
if not rs1.eof then
	sImgWebPath = ""
	sImgWebPath=toImageDir(rs1("ProsecutionTime"))

	sFIXEQUIPTYPE=""
	if trim(rs1("FIXEQUIPTYPE"))="1" then
		sFIXEQUIPTYPE="Type1"
	elseif trim(rs1("FIXEQUIPTYPE"))="2" then
		sFIXEQUIPTYPE="Type2"
	elseif trim(rs1("FIXEQUIPTYPE"))="3" then
		sFIXEQUIPTYPE="Type3"
	elseif trim(rs1("FIXEQUIPTYPE"))="5" then
		sFIXEQUIPTYPE="Type5"
	end if
	
	if sFIXEQUIPTYPE="Type3" then
		RealFileName=right(rs1("FileName").value,4)
		WebPicPathTmp=left(rs1("FileName").value,14)
	end if
end if

bIllegalDate=""
bIllegalAddressID=""
bIllegalAddress=""
bRule1=""
bForFeit1=""
bLoginID1=""
bBillMem1=""
bBillMemID1=""
bLoginID2=""
bBillMem2=""
bBillMemID2=""
bBillUnitID=""
bBillType=""
bDealLineDate=""
bBillFillDate=""
bRuleSpeed=""
bCarAddId=""
bRule4=""
'抓上一筆的資料
If sys_City="彰化縣" Then
strSql3="select * from (select * from BillBaseTmp" &_
	" where BillTypeID='2' and BillStatus in ('1') and RecordStateID=0 and RecordMemberID="&theRecordMemberID &_
	" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is not null order by RecordDate desc)" &_
	" where rownum=1"
Else
strSql3="select * from (select * from BillBase" &_
	" where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID &_
	" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is not null order by RecordDate desc)" &_
	" where rownum=1"
End If 
set rs13=conn.execute(strSql3)
if not rs13.eof then
	if trim(rs13("BillNo"))<>"" and not isnull(rs13("BillNo")) then
		bBillType="1"
	else
		bBillType="2"
	end if
	if trim(rs13("RuleSpeed"))<>"" and trim(rs13("RuleSpeed"))<>"0" And not isnull(rs13("RuleSpeed")) then
		bRuleSpeed=trim(rs13("RuleSpeed"))
	end	if
	if trim(rs13("IllegalDate"))<>"" and not isnull(rs13("IllegalDate")) then
		bIllegalDate=ginitdt(trim(rs13("IllegalDate")))
	end if
	If sys_City="高雄市" Then
		if trim(rs13("IllegalZip"))<>"" and not isnull(rs13("IllegalZip")) then
			bIllZip=trim(rs13("IllegalZip"))
		end	if
	end if 
	if trim(rs13("IllegalAddressID"))<>"" and not isnull(rs13("IllegalAddressID")) then
		bIllegalAddressID=trim(rs13("IllegalAddressID"))
	end	if
	if trim(rs13("IllegalAddress"))<>"" and not isnull(rs13("IllegalAddress")) then
		bIllegalAddress=trim(rs13("IllegalAddress"))
	end	if
	if trim(rs13("Rule1"))<>"" and not isnull(rs13("Rule1")) then
		bRule1=trim(rs13("Rule1"))
	end	if
	if trim(rs13("Rule4"))<>"" and not isnull(rs13("Rule4")) then
		bRule4=trim(rs13("Rule4"))
	end	if
	if trim(rs13("ForFeit1"))<>"" and not isnull(rs13("ForFeit1")) then
		bForFeit1=trim(rs13("ForFeit1"))
	end	if
	if trim(rs13("BillMemID1"))<>"" and not isnull(rs13("BillMemID1")) then
		strMem1="select LoginID from MemberData where MemberID="&trim(rs13("BillMemID1"))
		set rsMem1=conn.execute(strMem1)
		if not rsMem1.eof then
			bLoginID1=trim(rsMem1("LoginID"))
		end if
		rsMem1.close
		set rsMem1=nothing
	end if
	if trim(rs13("BillMem1"))<>"" and not isnull(rs13("BillMem1")) then
		bBillMem1=trim(rs13("BillMem1"))
	end if
	if trim(rs13("BillMemID1"))<>"" and not isnull(rs13("BillMemID1")) then
		bBillMemID1=trim(rs13("BillMemID1"))
	end If
	if trim(rs13("BillMemID2"))<>"" and not isnull(rs13("BillMemID2")) then
		strMem2="select LoginID from MemberData where MemberID="&trim(rs13("BillMemID2"))
		set rsMem2=conn.execute(strMem2)
		if not rsMem2.eof then
			bLoginID2=trim(rsMem2("LoginID"))
		end if
		rsMem2.close
		set rsMem2=nothing
	end if
	if trim(rs13("BillMem2"))<>"" and not isnull(rs13("BillMem2")) then
		bBillMem2=trim(rs13("BillMem2"))
	end if
	if trim(rs13("BillMemID2"))<>"" and not isnull(rs13("BillMemID2")) then
		bBillMemID2=trim(rs13("BillMemID2"))
	end if
	if trim(rs13("BillUnitID"))<>"" and not isnull(rs13("BillUnitID")) then
		bBillUnitID=trim(rs13("BillUnitID"))
	end if
	if trim(rs13("DealLineDate"))<>"" and not isnull(rs13("DealLineDate")) then
		bDealLineDate=ginitdt(trim(rs13("DealLineDate")))
	end if
	if trim(rs13("BillFillDate"))<>"" and not isnull(rs13("BillFillDate")) then
		bBillFillDate=trim(ginitdt(rs13("BillFillDate")))
	end if
end if 
rs13.close
set rs13=nothing
%>
<title>數位固定桿違規影像建檔</title>
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
<%if sys_City<>"台中縣" then%>
<!-- #include file="../Common/Bannernoimage.asp"-->
<%end if%>
<table width='<%
If sys_City=ApconfigureCityName Then
	response.write "1150"
elseIf sys_City="苗栗縣" Then
	response.write "1200"
elseIf sys_City="高雄市" Then
	response.write "1150"
Else
	response.write "1000"
End If 
%>' border='1' align="left" cellpadding="0">
	<tr>
		<td rowspan="4" valign="top" >
		<!-- 影像大圖 -->
	<%if not rs1.eof then%>
		
		<%
		
		if trim(rs1("FixEquipType"))="8" Then
			If HowCatchPicture="0" then
				if len(trim(rs1("IMAGEFILENAMEA")))>14 then
					Type_IMAGEFILENAMEA=right(trim(rs1("IMAGEFILENAMEA")),len(trim(rs1("IMAGEFILENAMEA")))-14)
				else
					Type_IMAGEFILENAMEA=trim(rs1("IMAGEFILENAMEA"))
				end If
			Else
				Type_IMAGEFILENAMEA=trim(rs1("IMAGEFILENAMEA"))
			End If 
		else
			Type_IMAGEFILENAMEA=trim(rs1("IMAGEFILENAMEA"))
		end if
		%>
		</div>
		<div id="Layer5g7" style="position:absolute; width:auto; height:18px; z-index:0;  border: 1px none #000000; color: #336633; background-color: #FFFFFF; font-weight: bold; left:400px; top:3px;">
		<%
		if trim(rs1("FixEquipType"))="8" then
			response.write "&nbsp;"&Type_IMAGEFILENAMEA&"&nbsp;"
		end if
		%>
		</div>
		<%
		If sys_City="彰化縣" Then
			FileLocation="D:\image\finish"&trim(rs1("DirectoryName"))
			dim fso1 
			set fso1=Server.CreateObject("Scripting.FileSystemObject")
			if (fso1.FileExists(FileLocation & Type_IMAGEFILENAMEA)=false) Then
				arrType_IMAGEFILENAMEA=Split(Type_IMAGEFILENAMEA,".")
				Type_IMAGEFILENAMEA=arrType_IMAGEFILENAMEA(0)&".PNG"
			end if

			set fso1=nothing
		End If 
		'response.write Type_IMAGEFILENAMEA
		if HowCatchPicture="0" then
			bPicWebPath=PicturePath & Type_IMAGEFILENAMEA
		else
			bPicWebPath=replace(replace(sImgWebPath & trim(rs1("DirectoryName")),"\","/") & "/" & Type_IMAGEFILENAMEA,"//","/")
		end if
		'response.write bPicWebPath
		%>
		<%if bPicWebPath<>"" then%>
		<img src="<%=bPicWebPath%>" border=1 height="<%
	If sys_City=ApconfigureCityName Then
		response.write "570"
	elseIf sys_City="苗栗縣" Then
		response.write "570"
	Else
		response.write "460"
	End If 
		%>" <%
		'放大鏡功能
		if isBig="Y"  then
		%>onmousemove="show(this, '<%=bPicWebPath%>')" onmousedown="show(this, '<%=bPicWebPath%>')"<%
		end if
		%> id="imgSource" src="<%=bPicWebPath%>">

			<div id="div1" style="position:absolute; overflow:hidden; width:<%
			If sys_City="高雄市" Then
				response.write "330"
			elseIf sys_City=ApconfigureCityName Then
				response.write "230"
			Else
				response.write "210"
			End If 
			%>px; height:<%
			If sys_City=ApconfigureCityName Then
				response.write "110"
			Else
				response.write "90"
			End If 
			%>px; left:<%
			if trim(request("divX"))="" Then
				If sys_City="高雄市" Then
					response.write "420"
				elseIf sys_City=ApconfigureCityName Then
					response.write "650"
				elseIf sys_City="苗栗縣" Then
					response.write "1210"
				Else
					response.write "540"
				End If 
			else
				response.write trim(request("divX"))
			end if
			%>px; top:<%
			if trim(request("divY"))="" Then
				If sys_City=ApconfigureCityName Then
					response.write "490"
				elseIf sys_City="苗栗縣" Then
					response.write "40"
				Else
					response.write "400"
				End If 
			else
				response.write trim(request("divY"))
			end if
			%>px; z-index:1;border-right: white thin ridge; border-top: white thin ridge; border-left: white thin ridge; border-bottom: white thin ridge <%
		'放大鏡功能
		if isBig="N"  then
		%> ;visibility: hidden;<%
		end if
		%>" onMousedown="initializedragie( )">
				<img id="BigImg" style='position:relative' src="<%=bPicWebPath%>">
			
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
		<td height="80" width="24%" valign="bottom">
		
		<!--<span class="style4">路段</span>-->
	<%If sys_City="宜蘭縣" then%>
		<input type="button" name="uploadb1" value='上傳相片' onclick='window.open("SubunitImageUpload.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 60px; height: 28px">
		<input type="button" name="uploadgd1" value='上傳相片(分局用)' onclick='window.open("SubUnitImageUpload_YL.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 95px; height: 28px">
		<br>
		<span class="style66">上傳相片前，請先將相片縮小到適當大小(請勿超過800KB)，檔案過大會造成讀取問題</span><br>
		<span class="style67">(如使用上傳相片發生錯誤，請改用『分局用』)</span>
		<br>
		<input type="button" name="uploadb1A" value='上傳相片(新)' onclick='window.open("KeyInupload/default.asp","KeyInupload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 80px; height: 28px">
	<%ElseIf sys_City="嘉義縣" then%>
		<!-- <input type="button" name="uploadb1" value='上傳相片' onclick='window.open("SubunitImageUpload.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 60px; height: 28px"> -->
		<input type="button" name="uploadgd1" value='超速闖紅燈儀器舉發相片' onclick='window.open("SubUnitImageUpload_CY.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 155px; height: 28px">
		<br>
		<input type="button" name="uploadgd1" value='員警手持相機舉發相片' onclick='window.open("SubUnitImageUpload_YL.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 155px; height: 28px">
		<br>
		<!-- <br>
		<span class="style67">(如使用上傳相片發生錯誤，請改用『分局用』)</span> -->
	<%ElseIf sys_City="高雄市" or sys_City="苗栗縣" then%>
		<input type="button" name="uploadb1" value='上傳相片' onclick='window.open("SubunitImageUpload.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 60px; height: 28px">
		<input type="button" name="uploadgd1" value='上傳相片(分隊用)' onclick='window.open("SubUnitImageUpload_KS.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 95px; height: 28px">
	<%ElseIf sys_City="屏東縣" then%>
		<input type="button" name="uploadb1" value='上傳相片' onclick='window.open("SubunitImageUpload.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 60px; height: 28px">
		<input type="button" name="uploadgd1" value='上傳相片(IE10)' onclick='window.open("SubUnitImageUpload_YL.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 8pt; width: 95px; height: 28px">
	<%else%>
		<input type="button" name="uploadb1" value='上傳舉發數位相片' onclick='window.open("SubunitImageUpload.asp","WebPageUpload","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 9pt; width: 115px; height: 28px">
	<%End If %>
	<%If sys_City="基隆市"  then%>
		<input type="button" name="uploadb3" value='上傳相片(限Ie11)' onclick='window.open("http://10.104.10.244/IllegalImageUpload/UploadillegalImage.aspx?UserID=<%=Trim(Session("Credit_ID"))%>","UploadillegalImage","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 9pt; width: 115px; height: 28px">
	<%End If %>
	<%If sys_City="彰化縣"  then%>
		<input type="button" name="uploadb3" value='上傳相片(限Ie11)' onclick='window.open("http://10.119.1.6/UploadillegalImage/UploadillegalImage.aspx?UserID=<%=Trim(Session("Credit_ID"))%>","UploadillegalImage","left=0,top=0,location=0,width=770,height=555,resizable=yes,scrollbars=yes")' style="font-size: 9pt; width: 115px; height: 28px">
	<%End If %>
	<%if not rs1.eof then%>
		<input type="button" name="Submit2CF32" onClick="funAllNotKeyInVerifyResult();" value="未建檔案件設定為無效" <%
		if rs1.eof then
			response.write "disabled"
		end if
			%> style="font-size: 9pt; width: 125px; height: 28px">
		<%If sys_City="嘉義縣" then%>
		<input  type="button" onClick="ChangeImg2()" value="圖切換" class="style4" style="height: 28px">
		<%End if%>
		<div align="left">
		<%If sys_City<>"嘉義縣" then%>
		影像存放位置
		<%End If %>
		<span class="style4"><a href="../ProsecutionImage/ProsecutionImage.asp?Location=<%
		if trim(request("getStreetName"))<>"all" and trim(request("getStreetName"))<>"" then
			response.write trim(request("getStreetName"))
		else
			response.write ""
		end if
		%>" target="_blank">
		&nbsp;&nbsp;
		已建檔  <%=Session("BillCnt_Image")%> / 剩餘 <%
		if trim(request("getStreetName"))<>"all" and trim(request("getStreetName"))<>"" then
			StrStreetPlus=" and PI.Location='"&trim(request("getStreetName"))&"'"
		else
			StrStreetPlus=""
		end if
		strStreetCnt="select count(*) as locationCnt from PI,PIDetail where PI.FILENAME = PIDetail.FILENAME and PI.OperatorA=PIDetail.Operator and PIDetail.Operator='"&trim(Session("Credit_ID"))&"' and FixEquipType in (1,2,5,8,10) and VERIFYRESULTID<>-1 and BillSn is null and (RejectCode<>'262' or RejectCode is null)"&StrStreetPlus&""
		set rsStreetCnt=conn.execute(strStreetCnt)
		if not rsStreetCnt.eof then
			response.write rsStreetCnt("locationCnt")
		end if
		rsStreetCnt.close
		set rsStreetCnt=Nothing
		'response.write strStreetCnt
		%> 張</a></span>
		<%If sys_City<>"嘉義縣" then%>
		<input type="text" Name="ImageSaveLocation" value="<%=UserPicturePath%>" size="12">
		
		<input type="button" value="確定" onclick="funcUpdSaveLocation();" class="style4">
		<%End If %>
		<%
		If ((sys_City="宜蘭縣" And Trim(Session("Unit_ID"))="TQ00") Or sys_City="高雄市" Or Trim(Session("Credit_ID"))="A000000000") And sys_City<>"嘉義縣" Then
		%>
		<input type="button" onClick="ChangeImg2()" value="圖切換" class="style4">
		<%
		End If 
		%>
		<!-- <input type="button" onClick="ChangeImg2()" value="圖切換" class="style4"> -->
		</div>
		
	<%end if%>
		</td>
		
	</tr>
	<tr>
		<td height="110" align="center">
	<%if not rs1.eof then%>
		<!-- 影像小圖 2-->
		<%
		fileName1=""
		fileName2=""
		PicName1=""
		PicName2=""
		Operator1=""
		Operator2=""
		strSQL2="select * from (select b.CarNo,b.CarSimpleID,a.ProsecutionTime,a.ProsecutionTypeID,a.SiteCode,a.FileName,a.DirectoryName,a.FIXEQUIPTYPE,a.VideoFileName,a.IMAGEFILENAMEA,a.IMAGEFILENAMEB,a .IMAGEFILENAMEC,a.Location,b.LawItemID,b.SN,a.LIMITSPEED,a.OVERSPEED,a.OperatorA,b.MemberID,b.Note from PI a, PIDetail b where a.FILENAME = b.FILENAME and a.OperatorA=b.Operator and b.Operator='"&trim(Session("Credit_ID"))&"' and FixEquipType in (1,2,5,8,10) and VERIFYRESULTID<>-1 and BillSn is null and (a.RejectCode<>'262' or a.RejectCode is null)"&strIgnorePlus & strOrder &") where rownum<=5"
		set rsP2=conn.execute(strSQL2)
		If Not rsP2.Bof Then rsP2.MoveFirst 
		for qqq=0 to 2
			if rsP2.eof then exit for
			if qqq=1 then
				fileName1=trim(rsP2("FileName"))
				PicName1=trim(rsP2("IMAGEFILENAMEA"))
				Operator1=trim(rsP2("OperatorA"))
				if trim(rsP2("FixEquipType"))="8" Then
					If HowCatchPicture="0" then
						if len(trim(rsP2("IMAGEFILENAMEA")))>14 then
							PicName1img=right(trim(rsP2("IMAGEFILENAMEA")),len(trim(rsP2("IMAGEFILENAMEA")))-14)
						else
							PicName1img=trim(rsP2("IMAGEFILENAMEA"))
						end If
					Else
						PicName1img=trim(rsP2("IMAGEFILENAMEA"))
					End if
				else
					PicName1img=trim(rsP2("IMAGEFILENAMEA"))
				end If
				If sys_City="彰化縣" Then
					FileLocation="D:\image\finish"&trim(rsP2("DirectoryName"))
					dim fso2 
					set fso2=Server.CreateObject("Scripting.FileSystemObject")
					if (fso2.FileExists(FileLocation & PicName1img)=false) Then
						arrPicName1img=Split(PicName1img,".")
						PicName1img=arrPicName1img(0)&".PNG"
						PicName1=arrPicName1img(0)&".PNG"
					end if

					set fso2=nothing
				End If 
				PicName1imgPath=replace(sImgWebPath & replace(trim(rsP2("DirectoryName")),"\","/") & "/" & trim(PicName1img),"//","/")
			elseif qqq=2 then
				fileName2=trim(rsP2("FileName"))
				PicName2=trim(rsP2("IMAGEFILENAMEA"))
				Operator2=trim(rsP2("OperatorA"))
				if trim(rsP2("FixEquipType"))="8" Then
					If HowCatchPicture="0" then
						if len(trim(rsP2("IMAGEFILENAMEA")))>14 then
							PicName2img=right(trim(rsP2("IMAGEFILENAMEA")),len(trim(rsP2("IMAGEFILENAMEA")))-14)
						else
							PicName2img=trim(rsP2("IMAGEFILENAMEA"))
						end If
					Else
						PicName2img=trim(rsP2("IMAGEFILENAMEA"))
					End If 
				else
					PicName2img=trim(rsP2("IMAGEFILENAMEA"))
				end If
				If sys_City="彰化縣" Then
					FileLocation="D:\image\finish"&trim(rsP2("DirectoryName"))
					dim fso3 
					set fso3=Server.CreateObject("Scripting.FileSystemObject")
					if (fso3.FileExists(FileLocation & PicName2img)=false) Then
						arrPicName2img=Split(PicName2img,".")
						PicName2img=arrPicName2img(0)&".PNG"
						PicName2=arrPicName2img(0)&".PNG"
					end if

					set fso3=nothing
				End If 
				PicName2imgPath=replace(sImgWebPath & replace(trim(rsP2("DirectoryName")),"\","/") & "/" & trim(PicName2img),"//","/")
			end if

			rsP2.MoveNext
		next
		rsP2.close
		set rsP2=nothing
		if trim(PicName1)<>"" and not isnull(PicName1) then
			if HowCatchPicture="0" then
				sPicWebPath2=PicturePath & trim(PicName1img)
			else
				sPicWebPath2=PicName1imgPath
			end if
		else
			sPicWebPath2=""
		end if
		%>
		<%if sPicWebPath2<>"" then%>
		<div id="Layer5g7" style="position:absolute; width:22px; height:22px; z-index:0;  border: 1px none #000000; color: #FF0000; background-color: #CCFFFF; font-weight: bold;">
		<input type="checkbox" name="SelectImage" value="1" onclick="ChangeImageCount()" <%
		if instr(trim(request("SelectImage")),"1")>0 then
			response.write "checked"
		end if
		%>>
		</div>
		<img src="<%=sPicWebPath2%>" border=1 id="SmallImg2" width="<%
		If sys_City="苗栗縣" Then
			response.write "300"
		else
			response.write "230"
		end if
		%>" height="<%
		If sys_City="苗栗縣" Then
			response.write "200"
		else
			response.write "170"
		end if
		%>" <%
		If (sys_City="宜蘭縣" And Trim(Session("Unit_ID"))="TQ00") Or sys_City="高雄市" Then
			response.write "ondblclick=""ChangeImg2()"""
		Else
			response.write "ondblclick=""OpenPic('"&sPicWebPath2&"')"""
		End If 
		%>>
		<!-- ondblclick="ChangeImg2()" -->
		<%else%>
		<div id="Layer5g7" style="position:absolute; width:22px; height:22px; z-index:0;  border: 1px none #000000; color: #FF0000; background-color: #CCFFFF; font-weight: bold;">
		<input type="checkbox" name="SelectImage" value="1" onclick="ChangeImageCount()" disabled>
		</div>
		<%end if%>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="110" align="center">
	<%if not rs1.eof then%>
		<!-- 影像小圖 1-->
		<%
		if trim(PicName2)<>"" and not isnull(PicName2) then
			if HowCatchPicture="0" then
				sPicWebPath=PicturePath & trim(PicName2img)
			else
				sPicWebPath=PicName2imgPath
			end if
		else
			sPicWebPath=""
		end if
		%>
		<%if sPicWebPath<>"" then%>
		<div id="Layer5g4" style="position:absolute; width:22px; height:22px; z-index:0;  border: 1px none #000000; color: #FF0000; background-color: #CCFFFF; font-weight: bold;">
		<input type="checkbox" name="SelectImage" value="2" onclick="ChangeImageCount()" <%
		if instr(trim(request("SelectImage")),"2")>0 then
			response.write "checked"
		end if
		%>>
		</div>
		<img src="<%=sPicWebPath%>" border=1 id="SmallImg" width="<%
		If sys_City="苗栗縣" Then
			response.write "300"
		else
			response.write "230"
		end if
		%>" height="<%
		If sys_City="苗栗縣" Then
			response.write "190"
		else
			response.write "170"
		end if
		%>" onmousemove="show(this, '<%=bPicWebPath%>')" <%
		If (sys_City="宜蘭縣" And Trim(Session("Unit_ID"))="TQ00") Or sys_City="高雄市" Then
			response.write "ondblclick=""ChangeImg()"""
		Else
			response.write "ondblclick=""OpenPic('"&bPicWebPath&"')"""
		End If 
		%>>
		<!-- ondblclick="ChangeImg()"  -->
		<%else%>
		<div id="Layer5g4" style="position:absolute; width:22px; height:22px; z-index:0;  border: 1px none #000000; color: #FF0000; background-color: #CCFFFF; font-weight: bold;">
		<input type="checkbox" name="SelectImage" value="2" onclick="ChangeImageCount()" disabled>
		</div>
		<%end if%>
		<br>
			<input type="hidden" name="gImageFileNameA" value="<%
			piIMAGEPATHNAMEA = replace(sImgWebPath & replace(trim(rs1("DirectoryName")),"\","/") & "/" ,"//","/")
			If sys_City="彰化縣" Then
				response.write Type_IMAGEFILENAMEA
			Else
				response.write trim(rs1("IMAGEFILENAMEA"))
			End If 
			
			%>">
			<input type="hidden" name="gImagePathNameA" value="<%=piIMAGEPATHNAMEA%>">
		<%if (trim(PicName1)<>"" and not isnull(PicName1)) then%>
			<input type="hidden" name="gImageFileNameB" value="<%
			piIMAGEPATHNAMEB = replace(sImgWebPath & replace(trim(rs1("DirectoryName")),"\","/") & "/" ,"//","/") 
			response.write trim(PicName1)
			%>">
			<input type="hidden" name="gImagePathNameB" value="<%=piIMAGEPATHNAMEB%>">
			<input type="hidden" name="gImageFileNameC" value="<%
			piIMAGEPATHNAMEC = replace(sImgWebPath & replace(trim(rs1("DirectoryName")),"\","/") & "/" ,"//","/") 
			response.write trim(PicName2)
			%>">
			<input type="hidden" name="gImagePathNameC" value="<%=piIMAGEPATHNAMEC%>">
			<input type="hidden" name="gFileName1" value="<%=fileName1%>">
			<input type="hidden" name="gFileName2" value="<%=fileName2%>">
			<input type="hidden" name="gOperator1" value="<%=Operator1%>">
			<input type="hidden" name="gOperator2" value="<%=Operator2%>">
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="17" align="center">
			<input type="radio" name="PicCount" value="1" <%
			if trim(request("PicCount"))="" or trim(request("PicCount"))="1" then
				response.write "checked"
			end if
			%> onclick="ChangeImageCount2(1)">一張
			<input type="radio" name="PicCount" value="2" <%
			if trim(request("PicCount"))="2" then
				response.write "checked"
			end if
			%>>二張
			<input type="radio" name="PicCount" value="3" <%
			if trim(request("PicCount"))="3" then
				response.write "checked"
			end if
			%> onclick="ChangeImageCount2(3)">三張
		</td>
	</tr>
	<tr>
		<td height="100" colspan="2" valign="top">
		<%if not rs1.eof then%>
		<table width='100%' border='1' align="left" cellpadding="0">
			<tr>
				<td bgcolor="#FFFFCC" width="6%"><div align="right"> <span class="style3">＊</span>車號&nbsp;</div></td>
				<td width="12%">
				<input type="text" size="9" name="CarNo" onBlur="getVIPCar();" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) and trim(rs1("ProsecutionTypeID"))<>"R" then
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
				<td width="<%
			If sys_City="高雄市" Then
				response.write "19%"
			else
				response.write "14%"
			end if
				%>">
                    <!-- 簡式車種 -->
                    <input type="text" maxlength="1" size="2" value="<%
                    if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
                    	'response.write trim(rs1("CarSimpleID"))
                    end if
                    %>" name="CarSimpleID" onBlur="getRuleAll();" style=ime-mode:disabled onkeydown="funTextControl(this);">
                    <div id="Layer012" style="position:absolute; width:<%
				if sys_City="高雄市" Then
					response.write "175px"
				Else
					response.write "125px"
				End if
					%>; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1汽車 / 2拖車 / 3重機<br>&nbsp;/ 4輕機/5動力機械/6臨時車牌</font></div>
				</td>
				<td bgcolor="#FFFFCC" width="7%"><div align="right"><span class="style3">＊</span>違規時間</div></td>
				<td colspan="5" width="13%">
							<!-- 違規日期 -->
					<input type="text" size="6" maxlength="7" name="IllegalDate" class='Text1' value="<%
				if sys_City<>"苗栗縣" And sys_City<>"高港局" then
					if trim(rs1("ProsecutionTime"))<>"" and not isnull(rs1("ProsecutionTime")) then 
						response.write gInitDT(rs1("ProsecutionTime"))
					else
						response.write trim(bIllegalDate)
				End if
					end if%>" onBlur="getBillFillDate()" style=ime-mode:disabled onkeydown="funTextControl(this);" onkeyup="IllegalDateKeyUP()" >&nbsp;
							<!-- 違規時間 -->
					<input type="text" size="3" maxlength="4" name="IllegalTime" class='Text1' value="<%
					if trim(rs1("ProsecutionTime"))<>"" and not isnull(rs1("ProsecutionTime")) then 
						response.write Right("00"&hour(rs1("ProsecutionTime")),2)&Right("00"&minute(rs1("ProsecutionTime")),2)
					end if
					%>" onBlur="this.value=this.value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" onKeyUP="IllegalTimeKeyUP()">
<%
					if sys_City="花蓮縣" then	
						'if trim(rs1("SiteCode"))<>"" and not isnull(rs1("SiteCode")) then
							response.write "&nbsp; &nbsp; &nbsp; &nbsp;主機號碼："&trim(rs1("SiteCode"))
						'end If
					End If 
%>
				</td>
			</tr>
			<tr>
			<%if sys_City<>"苗栗縣" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>地點&nbsp;</div></td>
				<td colspan="3">
					<input type="text" size="4" value="<%
				If sys_City="花蓮縣" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						strAddressID="select StreetID from Street where Address='"&trim(rs1("Location"))&"'"
						set rsAddressID=conn.execute(strAddressID)
						if not rsAddressID.eof then
							response.write trim(rsAddressID("StreetID"))
						end if
						rsAddressID.close
						set rsAddressID=nothing
					end if		
				elseif sys_City="嘉義縣" Or sys_City="基隆市" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						strAddressID="select StreetID from Street where Address='"&trim(rs1("Location"))&"'"
						set rsAddressID=conn.execute(strAddressID)
						if not rsAddressID.eof then
							response.write trim(rsAddressID("StreetID"))
						end if
						rsAddressID.close
						set rsAddressID=nothing
					else
						response.write trim(bIllegalAddressID)
					end if	
				Else 
					response.write trim(bIllegalAddressID)
				End If 
					%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<input type="hidden" name="OldIllegalAddressID" value="<%=bIllegalAddressID%>">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<%if sys_City="高雄市" then %>
						區號
						<input type="text" class="btn5" size="3" value="<%=bIllZip%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButtonsmall.jpg" onclick="QryIllegalZip();">
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
					<input type="text" size="<%if sys_City="苗栗縣" then response.write "37" else response.write "28" end if%>" value="<%
				If sys_City="花蓮縣" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						response.write trim(rs1("Location"))
					end If
				elseif sys_City="嘉義縣" Or sys_City="基隆市" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						response.write trim(rs1("Location"))
					else
						response.write trim(bIllegalAddress)
					end If
				Else 
					response.write trim(bIllegalAddress)
				End If 
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()" <%if sys_City="南投縣" then response.write "disabled"%>>
					<div id="Layerert45" style="position:absolute ; width:30px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><span class="style1">快速道路</span>
                </td>
		<%end if%>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style3">＊</span>法條一</div></td>
				<td colspan="5">
		<%if sys_City="苗栗縣" then%>
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
				If left(Trim(bRule1),5)="33101" Or left(Trim(bRule1),5)="43102" Or left(Trim(bRule1),2)="40" then
					if trim(rs1("OVERSPEED"))<>"" and trim(rs1("OVERSPEED"))<>"0" and not isnull(rs1("OVERSPEED")) then
						response.write trim(rs1("OVERSPEED"))
					end If
				End if 
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
				If left(Trim(bRule1),5)="33101" Or left(Trim(bRule1),5)="43102" Or left(Trim(bRule1),2)="40" then
					if trim(rs1("LIMITSPEED"))<>"" and trim(rs1("LIMITSPEED"))<>"0" and not isnull(rs1("LIMITSPEED")) then
						response.write trim(rs1("LIMITSPEED"))
					else
						response.write trim(bRuleSpeed)
					end If
				End if
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
		<%End If %>
					<input type="text" maxlength="9" size="7" value="<%
				if sys_City<>"苗栗縣" then
						response.write trim(bRule1)
				End if
					%>" name="Rule1" onKeyUp="getRuleData1();" style=ime-mode:disabled onkeydown="funTextControl(this);" >
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<img src="../Image/BillLawPlusButton_Small.JPG" onclick="Add_LawPlus()" alt="附加說明">
		<%if sys_City<>"苗栗縣" then%>
			<%If sys_City="南投縣" Or sys_City="屏東縣" Or sys_City="花蓮縣" then%>
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
					if trim(rs1("LIMITSPEED"))<>"" and trim(rs1("LIMITSPEED"))<>"0" and not isnull(rs1("LIMITSPEED")) then
						response.write trim(rs1("LIMITSPEED"))
					else
						response.write trim(bRuleSpeed)
					end if
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
					if trim(rs1("OVERSPEED"))<>"" and trim(rs1("OVERSPEED"))<>"0" and not isnull(rs1("OVERSPEED")) then
						response.write trim(rs1("OVERSPEED"))
					end if
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
			<%ElseIf sys_City="嘉義縣" Or sys_City="基隆市" then%>
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
					if trim(rs1("OVERSPEED"))<>"" and trim(rs1("OVERSPEED"))<>"0" and not isnull(rs1("OVERSPEED")) then
						response.write trim(rs1("OVERSPEED"))
					end If
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
					if trim(rs1("LIMITSPEED"))<>"" and trim(rs1("LIMITSPEED"))<>"0" and not isnull(rs1("LIMITSPEED")) And trim(rs1("OVERSPEED"))<>"" And trim(rs1("OVERSPEED"))<>"0" then
						response.write trim(rs1("LIMITSPEED"))
					else
						response.write trim(bRuleSpeed)
					end If
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
			<%else%>
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
				If left(Trim(bRule1),5)="33101" Or left(Trim(bRule1),5)="43102" Or left(Trim(bRule1),2)="40" then
					if trim(rs1("OVERSPEED"))<>"" and trim(rs1("OVERSPEED"))<>"0" and not isnull(rs1("OVERSPEED")) then
						response.write trim(rs1("OVERSPEED"))
					end If
				End if 
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
				If left(Trim(bRule1),5)="33101" Or left(Trim(bRule1),5)="43102" Or left(Trim(bRule1),2)="40" then
					if trim(rs1("LIMITSPEED"))<>"" and trim(rs1("LIMITSPEED"))<>"0" and not isnull(rs1("LIMITSPEED")) then
						response.write trim(rs1("LIMITSPEED"))
					else
						response.write trim(bRuleSpeed)
					end If
				End if
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
			<%End If %>
		<%End If %>
					&nbsp;
					<span class="style5">
					<div id="Layer1" style="position:absolute ; width:230px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><%
				if sys_City<>"苗栗縣" then
						strRule1="select IllegalRule,Level1 from Law where ItemID='"&trim(bRule1)&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule1=conn.execute(strRule1)
						if not rsRule1.eof then
							response.write trim(rsRule1("IllegalRule"))
							gLevel1=trim(rsRule1("Level1"))
							if trim(bRule4)<>"" then
								response.write "("&bRule4&")"
							end if
						end if
						rsRule1.close
						set rsRule1=nothing		
				End If 
					%></div></span>
					<input type="hidden" name="ForFeit1" value="<%
				if sys_City<>"苗栗縣" then
						response.write bForFeit1
				End If 
					%>">
					
				</td>
			<%if sys_City="苗栗縣" then%>
				<td bgcolor="#FFFFCC" ><div align="right">法條二</div></td>
				<td colspan="3">
					<input type="text" maxlength="9" size="7" value="<%

					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer2" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%

					%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%

					%>">

				</td>
			<%end if%>
		    </tr>
			<tr>
			<%if sys_City="苗栗縣" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>地點&nbsp;</div></td>
				<td colspan="5">
					<input type="text" size="4" value="<%
				If sys_City="花蓮縣" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						strAddressID="select StreetID from Street where Address='"&trim(rs1("Location"))&"'"
						set rsAddressID=conn.execute(strAddressID)
						if not rsAddressID.eof then
							response.write trim(rsAddressID("StreetID"))
						end if
						rsAddressID.close
						set rsAddressID=nothing
					end if		
				Else 
					response.write trim(bIllegalAddressID)
				End If 
					%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<input type="text" size="<%if sys_City="苗栗縣" then response.write "37" else response.write "28" end if%>" value="<%
				If sys_City="花蓮縣" Then
					if trim(rs1("Location"))<>"" and not isnull(rs1("Location")) then
						response.write trim(rs1("Location"))
					end If
				Else 
					response.write trim(bIllegalAddress)
				End If 
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()">
					<div id="Layerert45" style="position:absolute ; width:30px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><span class="style1">快速道路</span>
                </td>
			<%end if%>
			<%if sys_City<>"苗栗縣" then%>
				<td bgcolor="#FFFFCC" ><div align="right">法條二</div></td>
				<td colspan="3">
					<input type="text" maxlength="8" size="7" value="<%

					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer2" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%

					%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%

					%>">

				</td>
			<%end if%>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><span class="style3">＊</span>舉發人&nbsp;</div></td>
		  		<td colspan="<%
				If sys_City="高雄市" or sys_City="苗栗縣" Then
					response.write "3"
				Else
					response.write "5"
				End If 
				%>">
					<input type="text" size="<%If sys_City="苗栗縣" then response.write "9" Else response.write "7" end if%>" name="BillMem1" value="<%
					response.write bLoginID1
					BillRecordID=bBillMemID1
				%>" onKeyUp="getBillMemID1();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer12" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #00000; "><%=bBillMem1%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID%>" name="BillMemID1">
					<input type="hidden" value="<%
						response.write bBillMem1
					%>" name="BillMemName1">
<%If sys_City<>"宜蘭縣" And sys_City<>"苗栗縣" then%>
			<%If sys_City="高雄市" then%>
					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">舉發人二</div></td>
					<td >
						
						<input type="text" size="7" name="BillMem2" value="<%
						response.write bLoginID2
						BillRecordID2=bBillMemID2
					%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
						<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
						<span class="style5">
						<div id="Layer13" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=bBillMem2%></div>
						</span>
						<input type="hidden" value="<%=BillRecordID2%>" name="BillMemID2">
						<input type="hidden" value="<%
							response.write bBillMem2
						%>" name="BillMemName2">
					</td>
			<%else%>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
					<input type="hidden" size="4" name="BillMem2" value="<%
					response.write trim(request("BillMem2"))
					BillRecordID2=trim(request("BillMemID2"))
				%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
			
					<span class="style5">
					<div id="Layer13" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName2"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID2%>" name="BillMemID2">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName2"))
					%>" name="BillMemName2">
			<%End If %>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
					<input type="hidden" size="4" name="BillMem3" value="<%
					response.write trim(request("BillMem3"))
					BillRecordID3=trim(request("BillMemID3"))
				%>" onKeyUp="getBillMemID3();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					
					<span class="style5">
					<div id="Layer14" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName3"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID3%>" name="BillMemID3">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName3"))
					%>" name="BillMemName3">

					<input type="hidden" size="4" name="BillMem4" value="<%
					response.write trim(request("BillMem4"))
					BillRecordID4=trim(request("BillMemID4"))
				%>" onKeyUp="getBillMemID4();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<div id="Layer17" style="position:absolute ; width:130px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName4"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID4%>" name="BillMemID4">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName4"))
					%>" name="BillMemName4">
<%End if%>
				</td>
			</tr>
<%If sys_City="宜蘭縣" Or sys_City="苗栗縣" then%>
			<tr>
				<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">舉發人二</div></td>
		  		<td colspan="2">
					
					<input type="text" size="<%If sys_City="苗栗縣" then response.write "9" Else response.write "7" end if%>" name="BillMem2" value="<%
					response.write trim(request("BillMem2"))
					BillRecordID2=trim(request("BillMemID2"))
				%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer13" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName2"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID2%>" name="BillMemID2">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName2"))
					%>" name="BillMemName2">
				</td>
				<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">舉發人三</div></td>
		  		<td colspan="2" >
					<input type="text" size="<%If sys_City="苗栗縣" then response.write "9" Else response.write "7" end if%>" name="BillMem3" value="<%
					response.write trim(request("BillMem3"))
					BillRecordID3=trim(request("BillMemID3"))
				%>" onKeyUp="getBillMemID3();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer14" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName3"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID3%>" name="BillMemID3">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName3"))
					%>" name="BillMemName3">
				</td>
				<td bgcolor="#FFFFCC" height="30" colspan="1"><div align="right" style="font-size: 12px ;">舉發人四</div></td>
				<td colspan="3">
					<input type="text" size="<%If sys_City="苗栗縣" then response.write "9" Else response.write "7" end if%>" name="BillMem4" value="<%
					response.write trim(request("BillMem4"))
					BillRecordID4=trim(request("BillMemID4"))
				%>" onKeyUp="getBillMemID4();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer17" style="position:absolute ; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%=trim(request("BillMemName4"))%></div>
					</span>
					<input type="hidden" value="<%=BillRecordID4%>" name="BillMemID4">
					<input type="hidden" value="<%
						response.write trim(request("BillMemName4"))
					%>" name="BillMemName4">
				</td>
			</tr>
<%End if%>
			<tr>

				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span><span class="style4">舉發單位</span></div></td>
				<td colspan="3">
					<input type="text" size="4" name="BillUnitID" value="<%=bBillUnitID%>" onKeyUp="getUnit();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:200px; height:30px; z-index:0;  border: 1px none #000000; "><%
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
					</span>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span class="style4">民眾檢舉時間</span>
					<input type="text" name="JurgeDay" value="" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:60px;" onblur="this.value=this.value.replace(/[^\d]/g,'');">
		<%if sys_City="高雄市" then%>
					<span class="style4">局信箱</span>
					<input type="text" name="ReportCaseNo" value="" style=ime-mode:disabled onkeydown="funTextControl(this);" style="width:110px;" >
		<%End if %>
				</td>
				<td bgcolor="#FFFFCC" width="8%">
		<%if sys_City<>"苗栗縣" then%>
				<div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>
		<%End if%>
				<div align="right"><span class="style3">＊</span>填單日期</div></td>
				<td width="<%
			If sys_City="高雄市" Then
				response.write "6%"
			else
				response.write "9%"
			end if
				%>">
				
				&nbsp;<input type="text" size="6" value="<%=ginitdt(date)%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">
				<input type="hidden" name="SelFileName" value="<%=trim(rs1("FileName"))%>">
				<input type="hidden" name="SelSN" value="<%=trim(rs1("SN"))%>">
				<input type="hidden" name="SelOperator" value="<%=trim(rs1("OperatorA"))%>">
				</td>
		<%if sys_City<>"苗栗縣" then%>
				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種&nbsp;</td>
				<td width="6%">
                &nbsp;<input type="text" maxlength="2" size="4" value="" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
                
				</td>
		<%End If %>
				<td bgcolor="#FFFFCC" width="8%">
		
				<div align="right">專案代碼&nbsp;</div></td>
				<td width="12%">
					&nbsp;<input type="text" size="5" value="" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
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
			<%if sys_City<>"宜蘭縣" then%>
			    <!-- 備註 -->
					<input type="hidden" size="29" value="<%
					if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
						response.write trim(rs1("Note"))
					end If
					if sys_City="花蓮縣" then	
						if trim(rs1("SiteCode"))<>"" and not isnull(rs1("SiteCode")) then
							response.write trim(rs1("SiteCode"))
						end If
					End If 
					%>" name="Note" style=ime-mode:active>
			<%End if%>
				</td>
		<%if sys_City="苗栗縣" then%>
				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種&nbsp;</td>
				<td width="6%">
                &nbsp;<input type="text" maxlength="2" size="4" value="" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
				<div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>
				</td>
		<%End If %>
			</tr>
		<%if sys_City="宜蘭縣" then%>
			<tr>
				<td bgcolor="#FFFFCC" align="right" width="8%">備註&nbsp;</td>
				<td colspan="5">
                &nbsp;
				<input type="Text" size="29" value="" name="Note" style=ime-mode:active>
				</td>
			</tr>
		<%End If%>
		</table>
		<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
		<%end if%>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td height="28" colspan="2" align="center">
		<%
			if sys_City="南投縣" then
				CheckFlag=0
				str1x="select * from apconfigure where id=777"
				Set rs1x=conn.execute(str1x)
				If Not rs1x.eof Then
					CheckFlag=Trim(rs1x("value"))
				End If
				rs1x.close
				Set rs1x=Nothing 
				If CheckFlag=1 Then
					response.write "<font color='#FF0000'><strong>六分鐘 : 不可建檔</strong></font>"
				Else
					response.write "<font color='#FF0000'><strong>六分鐘 : 可以建檔</strong></font>"
				End If 
			end if
				
				%>
	<%if sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="苗栗縣" then%>
			<input type="button" value="併上筆" onclick="BillMerge();"  <%
		if Session("BillCnt_Image")="0" then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 70px; height: 27px">
	<%end if%>
			<input type="button" value="儲 存 F2" onclick="InsertBillVase();"  <%
		if rs1.eof then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 70px; height: 27px">
			
			<input type="button" name="Submit1343" onClick="location='BillKeyIn_Image_Fix_CH.asp'" value="清 除 F4" style="font-size: 10pt; width: 70px; height: 27px">
			
			<input type="button" name="Submit5322" onClick="funcOpenBillQry()" value="查 詢 <%
			If sys_City="南投縣" Or sys_City="屏東縣" Or sys_City="嘉義縣" Then
				response.write "F6"
			Else
				response.write "F5"
			End If 
			%>" style="font-size: 10pt; width: 70px; height: 27px">
			<input type="hidden" name="kinds" value="">
		   
			<input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" style="font-size: 10pt; width: 70px; height: 27px">
			
			<input type="button" name="Submit2932" onClick="funVerifyResult();" value="無 效 F9" <%
		if rs1.eof then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 70px; height: 27px">
			
			<input type="button" name="Submit4232" onClick="funPrintCaseList_Report();" value="建檔清冊 F10" style="font-size: 10pt; width: 100px; height: 27px"> 
			<input type="button" name="Submit4f32" onClick="funIgnore();" value="略過 F11" style="font-size: 10pt; width: 70px; height: 27px">
			<img src="/image/space.gif" width="29" height="8">
<%'抓本機就不下載(暫不開放)
if HowCatchPicture="xxx" then %>
			<input type="button" name="Submite3f2"  onclick='window.open("<%
	 strType1="select Value from Apconfigure where ID=51"
	 set rsType1=conn.execute(strType1)
	 if not rsType1.eof then
		response.write trim(rsType1("value"))&Session("User_ID")
	 end if
	 rsType1.close
	 set rsType1=nothing
	  %>","WebPageUpload_Fix","location=0,width=770,height=455,resizable=yes,scrollbars=yes,toolbar=yes")' value="下載沖洗照片" style="font-size: 9pt; width: 100px; height: 27px">
<%end if%>
			<input type="button" name="SubmitBack2" onClick="location='BillKeyIn_Image_Fix_Back_CH.asp?PageType=First'" value="<< 第一筆 Home" style="font-size: 9pt; width: 100px; height: 27px">
			
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_Fix_Back_CH.asp?PageType=Back'" value="< 上一筆 PgUp" style="font-size: 9pt; width: 100px; height: 27px">
		<%If sys_City="高雄市" then%>
			&nbsp; &nbsp; 
			<input type="button" name="Submit4f32" onClick="funGetReportCase();" value="帶入檢舉資料" style="font-size: 10pt; width: 120px; height: 27px">
		<%End If%>
             <input type="hidden" name="Tmp_Order" value="<%=Session("BillCnt_Image")%>">
			<%If sys_City="嘉義縣" then%>
				<input type="checkbox" name="CaseInByMem" ><font style="font-size: 10pt">違規日逾三個月強制建檔</font>
			<%End If %>
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案日期 -->
				<input type="hidden" size="12" maxlength="7" name="DealLineDate">
				<!-- 應到案處所 -->
				<input type="hidden" size="10" value="" name="MemberStation">
				<input type="hidden" value="<%=bRule4%>" name="Rule4">
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
					Else
						response.write "540"
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
					Else
						response.write "400"
					End If 
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
var ProsecutionTypeID="<%
if not rs1.eof then
	response.write trim(rs1("ProsecutionTypeID"))
end if
%>";
var InsertFlag=0;
<%if sys_City="宜蘭縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID||Note");
<%elseif sys_City="南投縣" Or sys_City="屏東縣" or sys_City="花蓮縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,RuleSpeed,IllegalSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="高雄市" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1,BillMem2||BillUnitID,JurgeDay,ReportCaseNo,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="苗栗縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalSpeed,RuleSpeed,Rule1,Rule2||IllegalAddressID,IllegalAddress,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,ProjectID,CarAddID");
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
<%if sys_City="南投縣" then %>
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}
<%else%>
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}
	//else if (chkCarNoFormat(myForm.CarNo.value)==0){
	//	error=error+1;
	//		errorString=errorString+"\n"+error+"：違規車號格式錯誤。";
	//}
<%end if%>
	if (myForm.CarSimpleID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簡式車種。";
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
	}else if (!ChkIllegalDate(myForm.IllegalDate.value)){
	<%If sys_City="嘉義縣" then%>
		if (myForm.CaseInByMem.checked==false)
		{
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期已超過三個月期限。";
		}	
	<%else%>
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過三個月期限。";
	<%end if %>
	}
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
	}else if( myForm.BillFillDate.value.substr(0,1)=="0"  ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤，請直接輸入年份，開頭不須補0。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="9" && myForm.BillFillDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="1" && myForm.BillFillDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
	}else if (!ChkIllegalDate(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過三個月。";
	}
<%If sys_City="彰化縣" Then%>
	if (myForm.JurgeDay.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入民眾檢舉時間。";	
	}
<%end if%>
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
	}else if( myForm.DealLineDate.value.substr(0,1)=="0"  ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤，請直接輸入年份，開頭不須補0。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.DealLineDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過三個月。";
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
			errorString=errorString+"\n"+error+"：請輸入舉發人代碼。";
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
	if (eval(myForm.BillFillDate.value) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
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
<%if sys_City="台中市" then %>
	if ((myForm.Rule1.value.substr(0,2))=="55"){
		error=error+1;
		errorString=errorString+"\n"+error+"：第55條不可逕行舉發。";
	}
<%end if%>
	if ((myForm.Rule1.value.substr(0,3))=="293" && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
	if (myForm.PicCount(1).checked==true){
		if (myForm.SelectImage(0).checked==true && myForm.SelectImage(1).checked==true){
			error=error+1;
			errorString=errorString+"\n"+error+"：請確認要合併的照片。";
		}if (myForm.SelectImage(0).checked==false && myForm.SelectImage(1).checked==false){
			error=error+1;
			errorString=errorString+"\n"+error+"：請確認要合併的照片。";
		}
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 60){
				if ((myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,5))=="33101"){
					if (myForm.Rule1.value=="4340003" || myForm.Rule2.value=="4340003"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340003需另單舉發。";
					}
				}
			}
		}
	}
	if (error==0){
		if (InsertFlag==0){
			InsertFlag=1;
			getChkCarIllegalDate();
		}
	}else{
		alert(errorString);
	}
}
//併上筆
function BillMerge(){
	if(confirm('是否要將此張相片檔併入上案？')){
		myForm.kinds.value="BillMerge";
		myForm.submit();
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
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
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
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			//alert("車牌格式錯誤，如該車輛無車牌或舊式車牌則可忽略此訊息！");
			runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=2");
		}else{
			runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=2");
		<%if sys_City<>"高雄市" and sys_City<>"苗栗縣" and sys_City<>"宜蘭縣" then%>
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
		
		//event.keyCode=35;
		//	event.returnValue=true;
			//alert("sdfs");
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
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "5" && myForm.CarSimpleID.value != "6"){
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
		runServerScript("getRuleDetail_forLawPlus.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
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
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
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
		runServerScript("getMemberStation.asp?StationID="+StationNum);
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
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}

//舉發單位民眾檢舉用(ajax)
function getUnit_Report(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
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
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	
		if (myForm.IllegalAddressID.value.length == 6){
		<%if sys_City="苗栗縣" then %>
			myForm.IllegalAddress.select();
		<%else%>
			myForm.Rule1.select();
		<%end if%>
		}
		<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
		if (myForm.OldIllegalAddressID.value != myForm.IllegalAddressID.value)
		{
			myForm.IllegalZip.value="";
		}
		<%end if%>
	}
}
//舉發人一(ajax)
function getBillMemID1(){
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
	if (myForm.BillMem1.value.length > 2){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
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
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
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
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
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
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
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
		Dyear=parseInt(DLineDate.getYear())-1911;
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
		response.write "100"
	else
		response.write "61"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
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
		response.write "100"
	else
		response.write "61"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
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
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value.length >= 9){
		runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
	}
}

//檢查單號是否有在GETBILLBASE內
function setCheckBillNoExist(GetBillFlag,BillBaseFlag,BillSN,BillType,MLoginID,MMemberID,MMemName,MUnitID,MUnitName)
{
	if (GetBillFlag==0){
		alert("此單號不存在於領單紀錄中！");
		document.myForm.Billno1.value="";
	}else{
		document.myForm.BillMem1.value=MLoginID;
		document.myForm.BillMemID1.value=MMemberID;
		document.myForm.BillMemName1.value=MMemName;
		Layer12.innerHTML=MMemName;
		TDMemErrorLog1=0;
		if (document.myForm.BillUnitID.value==""){
			document.myForm.BillUnitID.value=MUnitID;
			Layer6.innerHTML=MUnitName;
			TDUnitErrorLog=0;
		}
		if (BillBaseFlag==1){
			alert("此單號已建檔！");
			document.myForm.Billno1.value="";
		}else if (BillBaseFlag==0){
			alert('此單號已建檔！');
			document.myForm.Billno1.value="";
		}
	}
}

//逕舉建檔清冊
function funPrintCaseList_Report(){
	UrlStr="../Query/PrintCaseDataList_Report.asp?CallType=1";
	newWin(UrlStr,"CaseListWin2342",980,575,0,0,"yes","yes","yes","no");
}

//審核無效
function funVerifyResult(){
	if(confirm('確定要將此筆違規影像設為無效？')){
		myForm.kinds.value="VerifyResultNull";
		myForm.submit();
	}
}
//
function funAllNotKeyInVerifyResult(){
	if(confirm('確定要將所有未建檔的違規影像設為無效？')){
		myForm.kinds.value="AllNotKeyInVerifyResultNull";
		myForm.submit();
	}
}
function ChangeImageCount(){
	if (myForm.SelectImage(0).checked==true && myForm.SelectImage(1).checked==true){
		myForm.PicCount(2).checked=true;
	}else if (myForm.SelectImage(0).checked==true && myForm.SelectImage(1).checked==false){
		myForm.PicCount(1).checked=true;
	}else if (myForm.SelectImage(0).checked==false && myForm.SelectImage(1).checked==true){
		myForm.PicCount(1).checked=true;
	}else{
		myForm.PicCount(0).checked=true;
	}
}

function ChangeImageCount2(PCnt){
	if (PCnt=="1"){
		myForm.SelectImage(0).checked=false;
		myForm.SelectImage(1).checked=false;
	}else if (PCnt=="3"){
		myForm.SelectImage(0).checked=true;
		myForm.SelectImage(1).checked=true;
	}
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
		funcOpenBillQry();
<%if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then %>
	}else if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%end if %>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
<%
	if not rs1.eof then
%>
		InsertBillVase();
<%
	end if
%>
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		event.returnValue=false;  
		location='BillKeyIn_Image_Fix_CH.asp'
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
%>
		funVerifyResult();
<%
	end if
%>
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		event.returnValue=false;  
		funPrintCaseList_Report();
	}else if (event.keyCode==122){ //F11略過
		event.keyCode=0;   
		event.returnValue=false;  
<%
	if not rs1.eof then
%>
		funIgnore();
<%
	end if
%>
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		location='BillKeyIn_Image_Fix_Back_CH.asp?PageType=Back'
<%if sys_City<>"苗栗縣" then %>
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		location='BillKeyIn_Image_Fix_Back_CH.asp?PageType=First'
<%end if%>
	}
}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=2;
	window.open("EasyBillQry.asp","WebPage86_Update","left=300,top=200,location=0,width=350,height=160,resizable=yes,scrollbars=yes");
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
function setIllegalRule(flag){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
	<%if not rs1.eof then%>
		if ((myForm.Rule1.value.substr(0,2))!="29"){
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
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
<%if sys_City="高雄市" then%>
var sys_City="<%=sys_City%>";
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes,status=yes");

}
function getIllZip(){
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value);
}

function funGetReportCase(){
	if (myForm.ReportCaseNo.value==""){
		alert("請先輸入局信箱編號!!");
	}else{
		runServerScript("getGetReportCase.asp?ReportCaseNo="+myForm.ReportCaseNo.value);
	}
}
<%end if %>
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
//略過
function funIgnore(){
	myForm.kinds.value="BillIgnore";
	myForm.submit();
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
		response.write "330"
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

function ChangeImg(){
<%if sPicWebPath<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImg.src;

	myForm.SmallImg.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;
<%end if%>
}
//============================================================

function ChangeImg2(){
<%if sPicWebPath2<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImg2.src;

	myForm.SmallImg2.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;
<%end if%>
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
		//if (CarType!=0){
			myForm.CarSimpleID.value=CarType;
		//}
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
