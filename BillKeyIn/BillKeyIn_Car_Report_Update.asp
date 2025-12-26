<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>逕舉資料修改</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
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
'告發類別
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

'修改告發單
if trim(request("kinds"))="DB_insert" then
	'有改單號的話，先檢查有沒有重覆的單號
	if trim(request("Billno1"))<>trim(request("OldBillNo")) and trim(request("Billno1"))<>"" then
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
	'違規日期	
	theIllegalDate=""
	if trim(request("IllegalDate"))<>"" Then
		If len(trim(request("IllegalTime")))=6 Then
			theIllegalDate="TO_DATE('"&gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2) &"','YYYY/MM/DD/HH24/MI/SS')"
		else
			theIllegalDate=funGetDate(gOutDT(request("IllegalDate")) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
		End if
	else
		theIllegalDate = "null"
	end if	
	'response.write theIllegalDate
	'檢查有沒有相同車號同時間同違規法條
	if sys_City<>"台東縣" then
		if trim(request("Rule1"))<>"" then
			strRule1=" and Rule1=to_char('"&trim(request("Rule1"))&"')"
		End If
		if trim(request("Rule2"))<>"" then
			strRule2=" and Rule2='"&trim(request("Rule2"))&"'"
		End If
		strChkCIL="select count(*) as cnt from billbase where Sn<>"&trim(request("BillSN")) &_
			" and CarNo='"&UCase(trim(request("CarNo")))&"'" &_
			" and IllegalDate="&theIllegalDate & strRule1 & strRule2 & " and RecordstateID=0"
		Set rsChkCIL=conn.execute(strChkCIL)
		If Not rsChkCIL.eof Then
			If Trim(rsChkCIL("cnt"))>0 then
				chkIsExistBillNumFlag=2
				If len(trim(request("IllegalTime")))=6 Then
					Illdate2=gOutDT(request("IllegalDate")) &" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)
				else
					Illdate2=gOutDT(request("IllegalDate")) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)
				End if
			End If 
		End If
		rsChkCIL.close
		Set rsChkCIL=Nothing 
	End If 

	IllegalTimeCheck_Str=""
	if sys_City="台中市" then
		If Trim(request("ReportNo"))<>Trim(request("OldReportNo")) Then
			strChkReportNo="select count(*) as cnt from BillBase a,BillReportNo b where b.ReportNo='" & trim(request("ReportNo")) & "'" &_
				" and a.RecordStateID=0 and a.SN=b.BillSN"
			Set rsCPNIE=conn.execute(strChkReportNo)
			If Not rsCPNIE.eof Then
				If CInt(rsCPNIE("cnt"))>0 Then
					chkIsExistBillNumFlag=3
				End If 
			End If 
			rsCPNIE.close
			Set rsCPNIE=Nothing 

		End If 

		If len(trim(request("IllegalTime")))=6 Then
			illegalDateTmp=gOutDT(request("IllegalDate")) &" "&left(trim(request("IllegalTime")),2)&":"&mid(trim(request("IllegalTime")),3,2)&":"&right(trim(request("IllegalTime")),2)
		else
			illegalDateTmp=gOutDT(request("IllegalDate"))&" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2)&":00"
		End if

		illegalDate1=DateAdd("h",-2,illegalDateTmp)
		illegalDate2=DateAdd("h",2,illegalDateTmp)
		strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"

		If ((Left(trim(request("Rule1")),2)="55" And Len(trim(request("Rule1")))=7) Or (Left(trim(request("Rule2")),2)="55" And Len(trim(request("Rule2")))=7) Or (Left(trim(request("Rule1")),2)="56"  And Len(trim(request("Rule1")))=7) Or (Left(trim(request("Rule2")),2)="56") And Len(trim(request("Rule2")))=7) And FlagRuleDetail<>5 Then
			
			strRep2="select IllegalDate from BillBase where CarNo='"&UCase(trim(request("CarNo")))&"' and RecordstateID=0 "&strIllDate&" and ((Rule1 like '55%' and length(Rule1)=7) or (Rule2 like '55%' and length(Rule1)=7) or (Rule1 like '56%' and length(Rule1)=7) or (Rule2 like '56%' and length(Rule1)=7)) and sn<>"&Trim(request("BillSN"))
			Set rsRep2=conn.execute(strRep2)
			If Not rsRep2.eof Then 
				IllegalTimeCheck_Str="此車號於 "& Trim(rsRep2("IllegalDate")) & " 有相同違規，請確認是否重複建檔。"

			End If 
			rsRep2.close
			Set rsRep2=Nothing 
		Else
			if trim(request("Rule1"))<>"" then
				strRule1=" and Rule1=to_char('"&trim(request("Rule1"))&"')"
			End If
			if trim(request("Rule2"))<>"" then
				strRule2=" and Rule2='"&trim(request("Rule2"))&"'"
			End If
			strRep="select IllegalDate from BillBase where CarNo='"&UCase(trim(request("CarNo")))&"' and RecordstateID=0 "&strIllDate&strRule1&strRule2&" and sn<>"&Trim(request("BillSN"))
			set rsRep=conn.execute(strRep)
			If Not rsRep.eof Then 
				IllegalTimeCheck_Str="此車號於 "& Trim(rsRep("IllegalDate")) & " 有相同違規，請確認是否重複建檔。"
			end if
			rsRep.close
			set rsRep=Nothing
		End If 
	End If 
	if chkIsExistBillNumFlag=0 then

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
		'只做查車的案件如果修改車牌就把狀態改未處理，並刪除查車的Dcilog
		if trim(request("CarNo"))<>trim(request("OldCarNo")) and trim(request("BillStatus"))="1" then
			if sys_City="苗栗縣" Then
				strDel="Update DciLog set FileName='',SeqNo='',DciReturnStatusID='' where BillSn="&trim(request("BillSN"))&" and ExchangeTypeID='A'"
				conn.execute strDel
			else
				strBillStatusPlus=",BillStatus='0'"

				strDel="delete from DciLog where BillSn="&trim(request("BillSN"))&" and ExchangeTypeID='A'"
				conn.execute strDel
			end if 
		else
			strBillStatusPlus=""
		end if
		'BillBase
		ColAdd=""
		If sys_City="高雄市" Or sys_City="台中市" Then
			ColAdd=",IllegalZip='"&trim(request("IllegalZip"))&"'"
		End if	
		If Trim(request("StartIllegalDate"))<>"" Then
			ColAdd=ColAdd&",StartIllegalDate=to_date('" & gOutDT(request("StartIllegalDate") ) &" "&left(trim(request("StartIllegalTime")),2)&":"&mid(trim(request("StartIllegalTime")),3,2)&":"&right(trim(request("StartIllegalTime")),2) & "','YYYY/MM/DD/HH24/MI/SS')"
		End If 
		If Trim(request("Distance"))<>"" Then
			ColAdd=ColAdd&",Distance="&trim(Request("Distance"))
		End If 

		strUpd="update BillBase set BillTypeID='"&trim(request("BillType"))&"',BillNo='"&UCase(trim(request("Billno1")))&"'" &_
			",CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&trim(request("CarSimpleID")) &_
			",CarAddID="&theCarAddID&",IllegalDate="&theIllegalDate&_
			",IllegalAddressID='"&trim(request("IllegalAddressID"))&"',IllegalAddress='"&trim(request("IllegalAddress"))&"'" &_
			",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed&",RuleSpeed="&theRuleSpeed &_
			",ForFeit1="&trim(request("ForFeit1"))&",Rule2='"&trim(request("Rule2"))&"',ForFeit2="&theForFeit2 &_
			",Rule3='"&trim(request("Rule3"))&"',ForFeit3="&theForFeit3&",Rule4='"&trim(request("Rule4"))&"'" &_
			",ForFeit4="&theForFeit4&",Insurance="&theInsurance&",UseTool="&theUseTool &_
			",ProjectID='"&theProjectID&"',DriverID='"&UCase(trim(request("DriverPID")))&"'" &_
			",DriverBirth="&theDriverBirth&",Driver='"&trim(request("DriverName"))&"'" &_
			",DriverAddress='"&trim(request("DriverAddress"))&"',DriverZip='"&trim(request("DriverZip"))&"'" &_
			",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
			",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
			",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
			",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
			",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
			",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
			",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
			",Note='"&trim(request("Note"))&"',FromNote='"&trim(request("FromNote"))&"',FromNoteId='"&trim(request("FromNoteId"))&"',EquipmentID='"&trim(request("IsMail"))&"',DriverSex='"&trim(request("DriverSex"))&"'" &_
			",RECORDSTATEID=0,TrafficAccidentType=''"&strBillStatusPlus&ColAdd &_
			",JurgeDay="&theJurgeDay &_
			",IsVideo='"&Trim(request("IsVideo"))&"'" &_
			" where SN="&trim(request("BillSN"))

			conn.execute strUpd
			if sys_City="台中市" then
				ConnExecute strUpd&"~!@"&trim(request("OldBillData")),353
			else
				ConnExecute strUpd,353
			end if

			'舉發單扣件明細檔 BillFastenerDetail
			strDel="delete from BillFastenerDetail where BillSN="&trim(request("BillSN"))
			conn.execute strDel
			if trim(request("Fastener1"))<>"" then
				strInsFastene1="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
							" values(BillFastenerDetail_seq.nextval,"&trim(request("BillSN"))&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener1"))&"','')"
				conn.execute strInsFastene1
				ConnExecute strInsFastene1,353
			end if
			if trim(request("Fastener2"))<>"" then
				strInsFastene2="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
							" values(BillFastenerDetail_seq.nextval,"&trim(request("BillSN"))&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener2"))&"','')"
				conn.execute strInsFastene2
				ConnExecute strInsFastene2,353
			end If
			
			'台中市要填告發單號
			if sys_City="台中市" Or sys_City="連江縣" Then
				strDelR="delete from BILLREPORTNO where BillSN="&trim(request("BillSN"))
				conn.execute strDelR
				If Trim(request("ReportNo"))<>"" Then
					strReportNo="insert into BillReportNo(BillSN,ReportNo)" &_
						" values("&trim(request("BillSN"))&",'"&trim(request("ReportNo"))&"')"
					conn.execute strReportNo
				End If 
			End If
				
			if sys_City="苗栗縣" Then
				If Trim(request("QryCarUpdate"))="Yes" Then
					strDciUpdate="Update Dcilog set FileName='',SeqNo='',DciReturnStatusID='' where BillSn="&Trim(request("BillSN"))&" and ExchangeTypeID='A'"
					conn.execute(strDciUpdate)
				End If 
			End If 
			'========================titan==========================
				thewhere=""
				CaseInDate=""
				CaseINCnt=0
				strCType="select a.DciCaseInDate from BillBaseDCIReturn a " &_
				" where a.BillNo='"&trim(request("Billno1"))&"' and a.CarNo='"&trim(request("CarNo"))&"'" &_
				" and ExchangeTypeID='W'" &_
				" and Status in ('Y','S','n') order by DciCaseInDate Desc"
				set rsCType=conn.execute(strCType)
				if not rsCType.eof then
					'response.write trim(rsCType("DciCaseInDate"))
					CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
					CaseINCnt=1
				else
					strCType2="select ExchangeDate from DciLog where BillSn='"&trim(request("BillSN"))&"' " &_
						" and ExchangeTypeID='W' " &_
						" and (DciReturnStatusID in ('Y','S','n') or DciReturnStatusID is Null)" &_
						" order by ExchangeDate Desc"
					set rsCType2=conn.execute(strCType2)
					if not rsCType2.eof then
						CaseInDate=trim(rsCType2("ExchangeDate"))
						CaseINCnt=1
					end if
					rsCType2.close
					set rsCType2=nothing
				end if
				rsCType.close
				set rsCType=nothing
				'計算入案幾天
				CountCaseIN=0
				if CaseInDate<>"" then
					CountCaseIN=dateDiff("d",CaseInDate,now)
				end if
				if Not ifnull(request("OldBillNo")) then
					thewhere=" where BillSn="&trim(request("BillSN"))
					'----入案成功超過7天才能修改----
					'if (CountCaseIN>7 and CaseINCnt=1) or CaseINCnt=0 then
						strUpd4="Update BillBaseDciReturn  set BillNo='"&trim(request("Billno1"))&"',CarNo='"&trim(request("CarNo"))&"',Rule1='"&trim(request("Rule1"))&"',Rule2='"&trim(request("Rule2"))&"' where BillNo='"&trim(request("OldBillNo"))&"' and CarNo='"&trim(request("OldCarNo"))&"' and ExchangeTypeID<>'A'"
						conn.execute strUpd4
						strSQL="update DCILog set BillNo='"&UCase(trim(request("Billno1")))&"',CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
						conn.execute strSQL
					'end if
					'-------------------------------
					strSQL="update BillMailHistory set BillNo='"&UCase(trim(request("Billno1")))&"',CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
					conn.execute strSQL
				else
					thewhere1=" where BillSn="&trim(request("BillSN"))
					'----入案成功超過7天才能修改----
					'if (CountCaseIN>7 and CaseINCnt=1) or CaseINCnt=0 then
						'thewhere2=" where CarNo='"&trim(request("OldCarNo"))&"'"
						'strSQL="update BillBaseDCIReturn set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
						'conn.execute strSQL
						strSQL="update DCILog set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere1
						conn.execute strSQL
					'end if
					'-------------------------------
					strSQL="update BillMailHistory set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere1
					conn.execute strSQL
				end if
			'=======================================================
%>
<script language="JavaScript">
<%	'入案超過七天
	if CountCaseIN>7 and CaseINCnt=1 then
%>
	alert("此舉發單入案 [ 已超過七天 ] ，須請監理站修改資料，以避免與監理站資料不同!\n(此訊息僅是提醒。)");
<%	'未到七天
	elseif CountCaseIN<=7 and CaseINCnt=1 then
%>
	alert("此舉發單入案 [ 未超過七天 ] ，必須刪除後，再重新建檔入案!\n(此訊息僅是提醒。)");
<%	'未入案
	elseif CaseINCnt=0 then
%>
	alert("修改完成!!");
<%

	end if
		If IllegalTimeCheck_Str<>"" Then
%>
		alert("修改完成，\n<%=IllegalTimeCheck_Str%>");
<%	
		End If 
%>
</script>
<%
	elseIf chkIsExistBillNumFlag=1 then
%>
<script language="JavaScript">
	alert("儲存失敗，此單號：<%=UCase(trim(request("Billno1")))%>，已經建檔!");
</script>
<%	
	Elseif chkIsExistBillNumFlag=2 then
%>
<script language="JavaScript">
	alert("儲存失敗，此車號於<%=Illdate2%>，有相同違規舉發，請確認有無連續開單。");
</script>
<%
	Elseif chkIsExistBillNumFlag=3 then
%>
<script language="JavaScript">
	alert("儲存失敗，此告示單號：<%=UCase(trim(request("ReportNo")))%>，已建檔！！");
</script>
<%

	end If
	
	If strBillStatusPlus<>"" Then
%>
<script language="JavaScript">
	alert("因為有修改車號，請至『上傳監理所-車籍查詢』重新上傳！！");
</script>
<%	
	End If 

end if

dcicheckWflag="0"
strDciChk="select sn from dcilog where billsn="&trim(request("BillSN"))
set rsDciChk=conn.execute(strDciChk)
if not rsDciChk.eof then
	dcicheckWflag="1"
end if 
rsDciChk.close
set rsDciChk=nothing

strSql="select * from BillBase where SN="&trim(request("BillSN"))
set rs1=conn.execute(strSql)
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
	font-size: 13px
}
.style8 {
	color: #FF0000;
	font-size: 15px
}
.style7 {
	color: #FF0000;
	font-size: 12px;
	line-height:15px
	}
.style10 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px
	}
.style11 {
	color: #FF0000;
	font-size: 18px
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<%if sys_City<>"台中縣" then%>
<!-- #include file="../Common/Bannernoimage.asp"-->
<%end if%>
	<form name="myForm" method="post">  
		<table width='985' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="6"><strong>逕舉資料修改</strong>&nbsp; &nbsp; 日期格式：1150101 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; 
				<input type="checkbox" name="CaseInByMem" value="1" <%if trim(request("CaseInByMem"))="1" then response.write "checked"%>>逾違規日期超過二個月強制建檔</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">單號</div></td>
				<td <%
		if sys_City<>"台中市" then
			'response.write "colspan='5'"
		end if
		%>>
				<input name="Billno1" type="text" value="<%
				if trim(rs1("Billno"))<>"" and not isnull(rs1("Billno")) then
					response.write trim(rs1("Billno"))
					OldBillData=OldBillData&",Billno="&trim(rs1("Billno"))
				else
					OldBillData=OldBillData&",Billno="
				end if
				%>" size="13" maxlength="9" <%
			If sys_City<>ApconfigureCityName then
				if trim(rs1("UseTool"))<>"8" then
					response.write "readonly"
				end If
			End if
				%> onblur="CheckBillNoExist()" onkeydown="funTextControl(this);">
				<input name="OldBillNo" type="hidden" value="<%
				if trim(rs1("Billno"))<>"" and not isnull(rs1("Billno")) then
					response.write trim(rs1("Billno"))
				end if
				%>">
	<%if sys_City="台中市" Or sys_City="連江縣" then%>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>告示單號</div></td>
				<td >
					<input type="text" size="10" name="ReportNo" onkeydown="funTextControl(this);" value="<%
				strRNo="select * from BillReportNo where billsn="&trim(rs1("SN"))
				Set rsRNO=conn.execute(strRNo)
				If Not rsRNO.eof Then
					response.write Trim(rsRNO("ReportNo"))
					OldReportNo=Trim(rsRNO("ReportNo"))
				End If
				rsRNO.close
				Set rsRNO=nothing
					
					%>" style=ime-mode:disabled maxlength="11">
					<input type="hidden" size="10" name="OldReportNo" value="<%
				response.write OldReportNo
					
					%>" style=ime-mode:disabled maxlength="11">
				</td>
	<%end if%>	
	<%if sys_City<>"基隆市" And sys_City<>"南投縣" then%>
				<td bgcolor="#FFFFCC"><div align="right" class="style4b">有無全程錄影</div></td>
				<td >
					<input type="radio" name="IsVideo" value="1" <%
				If Trim(rs1("IsVideo"))="1" Then
					response.write "checked"
				End If 
					%>>有
					<input type="radio" name="IsVideo" value="0" <%
				If Trim(rs1("IsVideo"))="0" Then
					response.write "checked"
				End If 
					%>>無
					&nbsp; &nbsp; 
					<input type="button" value="清除" style="height: 22px; width: 43px; font-size: 10pt;"
					onclick="IsVideo[0].checked=false;IsVideo[1].checked=false;">
	<%end if%>	
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td>
					<table >
					<tr>
					<td>
					<input type="text" size="12" maxlength="8" value="<%
					if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
						response.write trim(rs1("CarNo"))
						OldBillData="CarNo="&trim(rs1("CarNo"))
					else
						OldBillData="CarNo="
					end if
					%>" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);">
					<input type="hidden" value="<%
					if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
						response.write trim(rs1("CarNo"))
					end if
					%>" name="OldCarNo">
					<input type="hidden" value="<%
					if trim(rs1("BillStatus"))<>"" and not isnull(rs1("BillStatus")) then
						response.write trim(rs1("BillStatus"))
					end if
					%>" name="BillStatus">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer7" style="position:absolute; width:130px; height:25px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"><%
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
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="10" value="<%
				if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
					response.write trim(rs1("CarSimpleID"))
					OldBillData=OldBillData&",CarSimpleID="&trim(rs1("CarSimpleID"))
				else
					OldBillData=OldBillData&",CarSimpleID="
				end if
				%>" name="CarSimpleID" onBlur="getRuleAll();" onfocus="this.select();" onkeydown="funTextControl(this);">
				<font class="style7">1汽車/2拖車/3重機/4輕機/5動力機械/6臨時車牌/7試車牌</font>
				</td>
				
			</tr>
<%
	theStartTime=""
	if sys_City="台中市" or sys_City="高雄市" or sys_City="苗栗縣" or sys_City="基隆市" or sys_City="台東縣" or sys_City="花蓮縣" Then
		if trim(rs1("StartIllegalDate"))<>"" and not isnull(rs1("StartIllegalDate")) then
%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期(起)</div></td>
				<td>
				<input type="text" size="10" value="<%
				if trim(rs1("StartIllegalDate"))<>"" and not isnull(rs1("StartIllegalDate")) then
					response.write ginitdt(trim(rs1("StartIllegalDate")))
				end if
				%>" maxlength="7" name="StartIllegalDate" onblur="this.value=this.value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間(起)</div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if trim(rs1("StartIllegalDate"))<>"" and not isnull(rs1("StartIllegalDate")) then
					theStartTime=Right("00"&hour(rs1("StartIllegalDate")),2)&Right("00"&minute(rs1("StartIllegalDate")),2)&Right("00"&Second(rs1("StartIllegalDate")),2)
					response.write theStartTime
					OldBillData=OldBillData&",StartIllegalDate="&year(rs1("StartIllegalDate"))&"/"&month(rs1("StartIllegalDate"))&"/"&day(rs1("StartIllegalDate"))&" "&hour(rs1("StartIllegalDate"))&":"&minute(rs1("StartIllegalDate"))&":0"
				end if
				%>" maxlength="6" name="StartIllegalTime" onblur="this.value=this.value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%
		End if
	End If 
%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期<%
				If theStartTime<>"" Then
					response.write "(迄)"
				End If 
				%></div></td>
				<td>
				<input type="text" size="10" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					response.write ginitdt(trim(rs1("IllegalDate")))
				end if
				%>" maxlength="7" name="IllegalDate" onblur="getDealLineDate_Stop()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間<%
				If theStartTime<>"" Then
					response.write "(迄)"
				End If 
				%></div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				If theStartTime<>"" Then
					LenIllegalDate=6
				Else
					LenIllegalDate=4
				End If 
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					If Second(rs1("IllegalDate"))<>0 Or theStartTime<>"" Then
						theChangeTime=Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)&Right("00"&Second(rs1("IllegalDate")),2)
						LenIllegalDate=6
					Else
						theChangeTime=Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
						LenIllegalDate=4
					End If 
					
					response.write theChangeTime
					OldBillData=OldBillData&",IllegalDate="&year(rs1("IllegalDate"))&"/"&month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&hour(rs1("IllegalDate"))&":"&minute(rs1("IllegalDate"))&":0"
				end if
				%>" maxlength="<%
				response.write LenIllegalDate
				%>" name="IllegalTime" onblur="this.value=this.value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%
	if sys_City="台中市" or sys_City="高雄市" or sys_City="苗栗縣" or sys_City="基隆市" or sys_City="台東縣" or sys_City="花蓮縣" Then
		if trim(rs1("Distance"))<>"" and not isnull(rs1("Distance")) then
%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>距離(公尺)</div></td>
				<td colspan="5">
				<input type="text" size="10" value="<%
				if trim(rs1("Distance"))<>"" and not isnull(rs1("Distance")) then
					response.write trim(rs1("Distance"))
				end if
				%>" maxlength="7" name="Distance" onblur="this.value=this.value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>

			</tr>
<%
		End if
	End If 
%>
<%if sys_City="雲林縣" or sys_City="宜蘭縣" or sys_City="台南市" or sys_City="嘉義市" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
					OldBillData=OldBillData&",IllegalAddressID="&trim(rs1("IllegalAddressID"))
				else
					OldBillData=OldBillData&",IllegalAddressID="
				end if
				%>" name="IllegalAddressID" onkeyup="getillStreet();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<%if sys_City="台南市" then %>
						<input type="text" class="btn5" size="3" value="" name="IllegalZip" onBlur="getDriverZip(this,'IllegalAddress');" onkeydown="funTextControl(this);">
						區號
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
					<%end if%>
					<%if sys_City="高雄市" then %>
						區號
						<input type="text" class="btn5" size="3" value="<%
				if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
					bIllZip=trim(rs1("IllegalZip"))
					response.write trim(rs1("IllegalZip"))
				else
					bIllZip=""
				end if 
						%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(rs1("IllegalZip"))%>">
						
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
					<input type="text" size="<%
					if sys_City="台南市" Then
						response.write "22"
					Else
						response.write "31"
					End If 
					%>" value="<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write trim(rs1("IllegalAddress"))
					OldBillData=OldBillData&",IllegalAddress="&trim(rs1("IllegalAddress"))
				else
					OldBillData=OldBillData&",IllegalAddress="
				end if
				%>" name="IllegalAddress" onkeydown="funTextControl(this);">
				<input type="checkbox" name="chkHighRoad" value="1" onclick="setIllegalRule()" <%
					if Left(trim(rs1("Rule1")),2)="33" then
						response.write "checked"
					elseif trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						if Left(trim(rs1("Rule2")),2)="33" then
							response.write "checked"
						elseif trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
							if Left(trim(rs1("Rule3")),2)="33" then
								response.write "checked"
							end if
						end if
					end if
					%> <%if sys_City="南投縣" then response.write "disabled"%>><span class="style1">快速道路</span>
					<%if sys_City="台中市" then %>
						<table >
						<tr>
						<td>
						區號
						<input type="text" class="btn5" size="3" value="<%
				if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
					bIllZip=trim(rs1("IllegalZip"))
					response.write trim(rs1("IllegalZip"))
				else
					bIllZip=""
				end if 
						%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(rs1("IllegalZip"))%>">
						
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
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="5">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="9" size="10" value="<%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
					OldBillData=OldBillData&",Rule1="&trim(rs1("Rule1"))
				else
					OldBillData=OldBillData&",Rule1="
				end if
				%>" name="Rule1" onKeyUp="getRuleData1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=trim(rs1("RuleVer"))%>&sBillTypeID=2","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")' alt="查詢法條">
					<img src="../Image/BillLawPlusButton.jpg" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer1" style="position:absolute ; width:530px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					strCarImple=""
					if left(trim(rs1("Rule1")),4)="2110" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
						gLevel1=trim(rsR1("Level1"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					response.write "("&trim(rs1("Rule4"))&")"
				end if
				%></div>
					<input type="hidden" name="ForFeit1" value="<%
				if trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1")) then
					response.write trim(rs1("ForFeit1"))
				else
					if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
						response.write gLevel1
					end if
				end if
				%>">
					</td>
					</tr>
					</table>
				</td>
			</tr>
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
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
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
					OldBillData=OldBillData&",Rule2="&trim(rs1("Rule2"))
				else
					OldBillData=OldBillData&",Rule2="
				end if
				%>" name="Rule2" onKeyUp="getRuleData2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer2" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					strCarImple=""
					if left(trim(rs1("Rule2")),4)="2110" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
						gLevel2=trim(rsR1("Level1"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit2" value="<%
				if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
					response.write trim(rs1("ForFeit2"))
				else
					if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						response.write gLevel2
					end if
				end if
				%>">
					<img src="space.gif" width="565" height="2">
<%if trim(rs1("Rule3"))="" or isnull(rs1("Rule3")) then%>
					<img src="../Image/Law3.jpg" width="45" height="25" onclick='InsertLaw()' alt="違規法條三">
<%end if%>
					</td>
					</tr>
					</table>
				</td>
			</tr>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw1" align="right"><div align="right">違規法條三</div></td>
				<td colspan="5" id="TDLaw2">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="8" size="10" value="<%
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					response.write trim(rs1("Rule3"))
					OldBillData=OldBillData&",Rule3="&trim(rs1("Rule3"))
				else
					OldBillData=OldBillData&",Rule3="
				end if
				%>" name="Rule3" onKeyUp="getRuleData3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer3" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					strCarImple=""
					if left(trim(rs1("Rule3")),4)="2110" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
						gLevel3=trim(rsR1("Level1"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit3" value="<%
				if trim(rs1("ForFeit3"))<>"" and not isnull(rs1("ForFeit3")) then
					response.write trim(rs1("ForFeit3"))
				else
					if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
						response.write gLevel3
					end if
				end if
				%>">
					</td>
					</tr>
					</table>
				</td>
				
			</tr>
<%else%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw1" align="right"></td>
				<td colspan="5" id="TDLaw2"></td>
			</tr>
<%end if%>
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" and sys_City<>"高雄市" and sys_City<>ApconfigureCityName then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					strMem1="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
					end if
					rsMem1.close
					set rsMem1=nothing
					OldBillData=OldBillData&",BillMemID1="&trim(rs1("BillMemID1"))
				else
					OldBillData=OldBillData&",BillMemID1="
				end if
				%>" name="BillMem1" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer12" style="position:absolute ; width:130px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					response.write trim(rs1("BillMem1"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					response.write trim(rs1("BillMemID1"))
				end if
				%>" name="BillMemID1">
					<input type="hidden" value="<%
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					response.write trim(rs1("BillMem1"))
				end if
				%>" name="BillMemName1">
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼2</div></td>
		  		<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					strMem2="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID2"))
					set rsMem2=conn.execute(strMem2)
					if not rsMem2.eof then
						response.write trim(rsMem2("LoginID"))
					end if
					rsMem2.close
					set rsMem2=nothing
					OldBillData=OldBillData&",BillMemID2="&trim(rs1("BillMemID2"))
				else
					OldBillData=OldBillData&",BillMemID2="
				end if
				%>" name="BillMem2" onkeyup="getBillMemID2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer13" style="position:absolute ; width:130px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					response.write trim(rs1("BillMem2"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					response.write trim(rs1("BillMemID2"))
				end if
				%>" name="BillMemID2">
					<input type="hidden" value="<%
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					response.write trim(rs1("BillMem2"))
				end if
				%>" name="BillMemName2">
					</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼3</div></td>
		  		<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					strMem3="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID3"))
					set rsMem3=conn.execute(strMem3)
					if not rsMem3.eof then
						response.write trim(rsMem3("LoginID"))
					end if
					rsMem3.close
					set rsMem3=nothing
					OldBillData=OldBillData&",BillMemID3="&trim(rs1("BillMemID3"))
				else
					OldBillData=OldBillData&",BillMemID3="
				end if
				%>" name="BillMem3" onkeyup="getBillMemID3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer14" style="position:absolute ; width:130px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					response.write trim(rs1("BillMem3"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					response.write trim(rs1("BillMemID3"))
				end if
				%>" name="BillMemID3">
					<input type="hidden" value="<%
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					response.write trim(rs1("BillMem3"))
				end if
				%>" name="BillMemName3">
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼4</div></td>
		  		<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					strMem4="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID4"))
					set rsMem4=conn.execute(strMem4)
					if not rsMem4.eof then
						response.write trim(rsMem4("LoginID"))
					end if
					rsMem4.close
					set rsMem4=nothing
					OldBillData=OldBillData&",BillMemID4="&trim(rs1("BillMemID4"))
				else
					OldBillData=OldBillData&",BillMemID4="
				end if
				%>" name="BillMem4" onkeyup="getBillMemID4();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer17" style="position:absolute ; width:130px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
					response.write trim(rs1("BillMem4"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					response.write trim(rs1("BillMemID4"))
				end if
				%>" name="BillMemID4">
					<input type="hidden" value="<%
				if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
					response.write trim(rs1("BillMem4"))
				end if
				%>" name="BillMemName4">
					</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td>
					<table >
					<tr>
					<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					response.write trim(rs1("BillUnitID"))
					OldBillData=OldBillData&",BillUnitID="&trim(rs1("BillUnitID"))
				else
					OldBillData=OldBillData&",BillUnitID="
				end if
				%>" name="BillUnitID" onKeyUp="getUnit();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage_CarRep_UPdate","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer6" style="position:absolute ; width:190px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					strU="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
					set rsU=conn.execute(strU)
					if not rsU.eof then
						response.write trim(rsU("UnitName"))
					end if
					rsU.close
					set rsU=nothing
				end if
					%></div>
					</td>
					</tr>
					</table>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
					<input type="text" size="10" value="<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write ginitdt(trim(rs1("BillFillDate")))
					OldBillData=OldBillData&",BillFillDate="&trim(rs1("BillFillDate"))
				else
					OldBillData=OldBillData&",BillFillDate="
				end if
				%>" maxlength="7" name="BillFillDate" onblur="getDealLineDate();" onkeydown="funTextControl(this);">
				&nbsp; &nbsp; &nbsp; &nbsp; 民眾檢舉時間
				<input type="text" name="JurgeDay" value="<%
				if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then
					response.write ginitdt(trim(rs1("JurgeDay")))
					OldBillData=OldBillData&",JurgeDay="&trim(rs1("JurgeDay"))
				else
					OldBillData=OldBillData&",JurgeDay="
				end if
				%>" size="10" maxlength="7" onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">
				</td>
				
			</tr>
			
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">應到案日期</div></td>
				<td >
				
					<input type="text" size="6" value="<%
					if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
						response.write ginitdt(trim(rs1("DealLineDate")))
						OldBillData=OldBillData&",DealLineDate="&trim(rs1("DealLineDate"))
					else
						OldBillData=OldBillData&",DealLineDate="
					end if
					%>" maxlength="7" name="DealLineDate" onblur="this.value=this.value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);" <%
				if sys_City="基隆市" Or sys_City="花蓮縣" then '到案日不可修改
					response.write " readonly"
				End if%>>
			<%	if sys_City="基隆市" then%>
					<span class="style10">
					因審計室審查，到案日不可修改
					</span>
			<%	End if%>
			<%	if sys_City="花蓮縣" then%>
					<input type="checkbox" name="chkbDealLineDate" value="1" onclick='getDealLineDate();' <%
					If DateDiff("d",rs1("BillFillDate"),rs1("DealLineDate"))=30 Then
						response.write "checked"
					End if
					%>>30天
			<% End if%>
			<%'If sys_City<>"嘉義市" And sys_City<>"宜蘭縣" And sys_City<>"花蓮縣" And sys_City<>"台南市" And sys_City<>"彰化縣" And sys_City<>"基隆市" And sys_City<>"澎湖縣" And sys_City<>"苗栗縣" And sys_City<>"雲林縣" And sys_City<>"保二總隊三大隊二中隊" then%>
					<!-- 應到案處所： -->
					<input type="hidden" size="4" value="<%
					if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
						response.write trim(rs1("MemberStation"))
					end if
					%>" name="XXXMemberStation" onKeyup="getStation();">
						<div id="XXXLayer5" style="position:absolute ; width:241px; height:24px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000; visibility :visible;"></div>
			<% 'End if%>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">是否郵寄</div></td>
				<td colspan="3">
				<%
				if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
					OldBillData=OldBillData&",EquipMentID="&trim(rs1("EquipMentID"))
				else
					OldBillData=OldBillData&",EquipMentID="
				end if
				  'if trim(rs1("UseTool"))<>"8" then%>
					<!-- <input type="radio" name="IsMail" value="1" checked>是 -->
				<%'else%>
					<input type="radio" name="IsMail" value="1" <%
					If sys_City="澎湖縣" Then
							response.write "onclick='getDealLineDate();' " 
					End If
					if trim(rs1("EquipMentID"))<>"-1" or isnull(rs1("EquipMentID")) then
						response.write "checked"
					end if
					%>>是
					<input type="radio" name="IsMail" value="-1" <%
					If sys_City="澎湖縣" Then
							response.write "onclick='getDealLineDate();' " 
					End If
					if trim(rs1("EquipMentID"))="-1" then
						response.write "checked"
					end if
					%>>否
				<%'end if%>
				</td>
				
			</tr>
		<%'If sys_City="宜蘭縣" or sys_City="嘉義市" or sys_City="花蓮縣" or sys_City="台南市" or sys_City="彰化縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="苗栗縣" or sys_City="雲林縣" or sys_City="保二總隊三大隊二中隊" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">身分證號<br><span class="style10">非轉歸責案件勿填</span></div></td>
		  		<td>
					<input type="text" size="10" name="DriverPID" value="<%=Trim(rs1("DriverID"))%>" onBlur="this.value=this.value.toUpperCase();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">應到案處所<br><span class="style10">非轉歸責案件勿填</span></div>
				
				</td>
		  		<td colspan="5">
					<table >
					<tr>
					<td>
					<input type="text" size="5" value="<%
					if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
						response.write trim(rs1("MemberStation"))
						OldBillData=OldBillData&",MemberStation="&trim(rs1("MemberStation"))
					else
						OldBillData=OldBillData&",MemberStation="
					end if
					%>" name="MemberStation" onkeyup="getStation();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>
					</td>
					<td style="vertical-align:text-top;">
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:120px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					strS="select DciStationName from Station where StationID='"&trim(rs1("MemberStation"))&"'"
					set rsS=conn.execute(strS)
					if not rsS.eof then
						response.write trim(rsS("DciStationName"))
						If trim(rs1("MemberStation"))="41" Then
							response.write "(中和辦公室)"
						ElseIf trim(rs1("MemberStation"))="46" Then
							response.write "(蘆洲辦公室)"
						ElseIf trim(rs1("MemberStation"))="60" Then
							response.write "(大肚辦公室)"
						ElseIf trim(rs1("MemberStation"))="61" Then
							response.write "(北屯辦公室)"
						ElseIf trim(rs1("MemberStation"))="63" Then
							response.write "(豐原辦公室)"
						End if
					end if
					rsS.close
					set rsS=nothing
				end if
				%></div>
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
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					response.write trim(rs1("ProjectID"))
					OldBillData=OldBillData&",ProjectID="&trim(rs1("ProjectID"))
				else
					OldBillData=OldBillData&",ProjectID="
				end if
				%>" name="ProjectID" onkeyup='ProjectF5();' onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
			<%if sys_City="苗栗縣" then%>
					<font style="font-size:12px;">檢舉達人1 / 拖吊9 </font>
			<%End If %>
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer001" style="position:absolute ; width:190px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					strProject="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"
					set rsProject=conn.execute(strProject)
					if not rsProject.eof then
						response.write trim(rsProject("Name"))
					end if
					rsProject.close
					set rsProject=nothing
				end if
					%></div>
					</td>
					</tr>
					</table>
				</td>
				
				<td bgcolor="#FFFFCC"><div align="right">輔助車種</div></td>
				<td colspan="3">
					<table >
					<tr>
					<td>
					<input type="text" maxlength="2" size="3" value="<%
					if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
						response.write trim(rs1("CarAddID"))
						OldBillData=OldBillData&",CarAddID="&trim(rs1("CarAddID"))
					else
						OldBillData=OldBillData&",CarAddID="
					end if
					%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);">
					</td>
					<td style="vertical-align:text-top;">
					<div id="Layer110" style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機 /8拖吊<br>/9(550cc)重機 /10計程車/ 11危險物品<%
				if sys_City="雲林縣" Then
					response.write " /12幼兒車(課輔車)"
				End If 
				%></span>
					</div>
					</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align="right">採証工具</div></td>
				<td width="32%"><input type="text" maxlength="1" size="3" value="<%
				if trim(rs1("UseTool"))<>"" and not isnull(rs1("UseTool")) then
					response.write trim(rs1("UseTool"))
					OldBillData=OldBillData&",UseTool="&trim(rs1("UseTool"))
				else
					OldBillData=OldBillData&",UseTool="
				end if
				%>" name="UseTool" onBlur="getFixID();" onkeydown="funTextControl(this);" <%
				if trim(rs1("UseTool"))="8" then
				%>readonly<%
				End if
				%>>
				<div id="Layer11" style="position:absolute; width:278px; height:24px; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden">
				  <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>				  
				  <input type='text' size='6' name='FixID' value='<%
				if trim(rs1("EQUIPMENTID"))<>"" and not isnull(rs1("EQUIPMENTID")) then
					response.write trim(rs1("EQUIPMENTID"))
				end if
				%>' onkeyup="setFixEquip();">
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
				<td bgcolor="#FFFFCC" align="right" width="13%">備註</td>
				<td width="16%">
					<input type="text" size="15" value="<%
				if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
					response.write trim(rs1("Note"))
					OldBillData=OldBillData&",Note="&trim(rs1("Note"))
				else
					OldBillData=OldBillData&",Note="
				end if
				%>" name="Note" onkeydown="funTextControl(this);">
					
				</td>
				<td bgcolor="#FFFFCC" align="right" width="8%">代保管物</td>
				<td width="18%">
					1. <input type="text" size="2" value="<%
				itemTemp=""
				strItem="select * from BillFastenerDetail where BillSN="&trim(rs1("SN"))
				set rsItem=conn.execute(strItem)
				If Not rsItem.Bof Then rsItem.MoveFirst 
				While Not rsItem.Eof
					if itemTemp="" then
						itemTemp=trim(rsItem("FastenerTypeID"))
					else
						itemTemp=itemTemp&"&"&trim(rsItem("FastenerTypeID"))
					end if
				rsItem.MoveNext
				Wend
				rsItem.close
				set rsItem=nothing
				ItemVal=split(itemTemp,"&")
				if ubound(ItemVal)>=0 then
					response.write ItemVal(0)
				end if
				%>" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;">
                </div>
                <input type="hidden" value="<%
					response.write FVal1
					%>" name="Fastener1Val">
				2. <input type="text" size="2" value="<%
				if ubound(ItemVal)>=1 then
					response.write ItemVal(1)
				end if
				%>" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;">
                </div>
                <input type="hidden" value="<%=FVal2%>" name="Fastener2Val">
				</td>

			</tr>
<tr>
<td bgcolor="#FFFFCC"><div align="right">來源平台</div></td>
<td><input type="text" size="20" value="<%
				if trim(rs1("FromNote"))<>"" and not isnull(rs1("FromNote")) then
					response.write trim(rs1("FromNote"))
					OldBillData=OldBillData&",FromNote="&trim(rs1("FromNote"))
				else
					OldBillData=OldBillData&",FromNote="
				end if
				%>" name="FromNote" onkeydown="funTextControl(this);" style=ime-mode:active></td>
<td bgcolor="#FFFFCC"><div align="right">平台單號/案號</div></td>
<td colspan="3"><input type="text" size="20" value="<%
				if trim(rs1("FromNoteId"))<>"" and not isnull(rs1("FromNoteId")) then
					response.write trim(rs1("FromNoteId"))
					OldBillData=OldBillData&",FromNoteId="&trim(rs1("FromNoteId"))
				else
					OldBillData=OldBillData&",FromNoteId="
				end if
				%>" name="FromNoteId" onkeydown="funTextControl(this);" style=ime-mode:active></td>
</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="6">
					<span style="color: #FF0000;font-size: 18px;"><strong>※重要：如果要刪除案件，請勿在刪除案件前修改單號、車號，修改後會無法刪除案件</strong></span>
					<br>
					<input type="button" value="儲 存 <%
					if sys_City="台東縣" or sys_City="高雄縣" then
						response.write "F9"
					else
						response.write "F2"
					end if
					%>" onclick="InsertBillVase();" class="btn1" <%
					dciReturnCheck=0
					strDci="select count(*) as cnt from dcilog where BillSN="&trim(rs1("SN"))&" and exchangetypeid='W' and dcireturnstatusid is null"
					set rsDci=conn.execute(strDci)
					If Not rsDci.Bof Then 
						if cdbl(rsDci("cnt"))>0 then
							response.write "disabled"
							dciReturnCheck=1
						end if 
					end if 
					rsDci.close
					set rsDci=nothing
					
					%>>
					&nbsp; &nbsp; 
					<input type="button" value="修改監理站回傳資料" onclick="UpdateDciData();" class="btn1">
                    &nbsp; &nbsp; 
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉" class="btn1">
					<input type="hidden" value="<%=trim(rs1("RuleVer"))%>" name="RuleVerSion">
					<input type="hidden" value="" name="kinds">
					<input type="hidden" value="<%=OldBillData%>" name="OldBillData">
					<input type="hidden" value="<%=trim(request("BillSN"))%>" name="BillSN">
                </span>
				
				<%
				if dciReturnCheck=1 then
				%>
				<br>
				<span style="color: #9933FF;font-size: 20px;"><strong>(不可修改，案件上傳監理站尚未回傳，請等候案件回傳後再修改)</strong></span>	
				<%
				end if 
				%>
				<br>
				<span class="style8"><strong>(已入案成功七天內的舉發單必須刪除後，再重新建檔入案，入案七天後的舉發單，須請監理站修改資料，以避免與監理站資料不同)</strong></span>
			
					<!-- 告發類別 -->
					<input type="hidden" size="3" maxlength="1" value="<%
					if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
						response.write trim(rs1("BillTypeID"))
					end if
					%>" name="BillType" readonly>
					<!-- 附加說明 -->
				<input type="hidden" name="Rule4" value="<%=trim(rs1("Rule4"))%>">
				</td>
			</tr>
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

<%if trim(rs1("Rule3"))="" or isnull(rs1("Rule3")) then%>
	<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="宜蘭縣" or sys_City="嘉義市" or sys_City="花蓮縣" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="苗栗縣" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
<%else%>
	<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="宜蘭縣" or sys_City="嘉義市" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
	MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="苗栗縣" then%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
<%end if%>

//修改告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");

	if (myForm.UseTool.value=="8"){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (myForm.Billno1.value==""){
			error=error+1;
			errorString=error+"：逕舉手開單請輸入單號。";
		}else if(ReadBillNo.length!=9){     
			error=error+1;
			errorString=error+"：單號不足九碼。";
		}
	}
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}
<% if dciReturnCheck=1 then%>
	error=error+1;
	errorString=errorString+"\n"+error+"：不可修改，案件上傳監理站尚未回傳，請等候案件回傳後再修改。";	
<%end if %>
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
	}else if( myForm.IllegalDate.value.substr(0,1)=="9" && myForm.IllegalDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if( myForm.IllegalDate.value.substr(0,1)=="1" && myForm.IllegalDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		IllDateFlag=1;
	}else if (!ChkIllegalDateML(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過二個月期限，如確定要修改請勾選上方強制建檔，並在備註輸入超過期限原因。";
	}else if (!ChkIllegalDate(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==true && myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請於備註欄填寫違規日期超過二個月期限原因。";
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
<%If theStartTime<>"" Then%>
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
		}else if (!ChkIllegalDate(myForm.StartIllegalDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期(起)已超過三個月期限。";
		}

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
<%end if %>
	/*if (myForm.IllegalAddressID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點代碼。";
	}*/
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
<%end if%>
	}else if (!ChkIllegalDate(myForm.BillFillDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過三個月期限。";
	}
<%If sys_City="宜蘭縣" or sys_City="嘉義市" then%>
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
	}else if( myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.DealLineDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過三個月期限。";
	}
	if (myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發單位代號。";
	}else if (TDUnitErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發單位代號輸入錯誤。";
	}
	if (myForm.BillMem1.value==""){
		//固定桿不需要輸入舉發人
		if (myForm.UseTool.value!="1"){
		    error=error+1;
			errorString=errorString+"\n"+error+"：請輸入舉發人代碼1。";
		}
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
				//error=error+1;
				//errorString=errorString+"\n"+error+"：民眾檢舉時間已超過七天，民眾檢舉發生超過七日之交通違規，依法不得舉發。";	
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
	}
<%if sys_City="台中市" then%>
	if (myForm.ReportNo.value!=""){
		if (myForm.ReportNo.value.length<11){
			error=error+1;
			errorString=errorString+"\n"+error+"：告示單號不可少於11碼。";
		}
	}	
<%end if%>
<%if sys_City="雲林縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then %>
	if (TDVipCarErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：車號 "+myForm.CarNo.value+" 為業管車輛。";
	}
<%end if%>
<%if sys_City="台中市" then %>
	//if (((myForm.Rule1.value.substr(0,2))=="55" || (myForm.Rule2.value.substr(0,2))=="55") && (myForm.UseTool.value!="8")){
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
<%if dcicheckWflag="1" then%>
	if (myForm.Note.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：本案件已經上傳查車或入案，如要修改請在備註欄位填寫修改原因。";
	}
<%end if %>
<%if sys_City="台東縣" then %>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 61){
				if ((myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,3))=="431" || (myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,3))=="431" || (myForm.Rule2.value.substr(0,5))=="33101"){
					if (myForm.Rule1.value=="4340003" || myForm.Rule2.value=="4340003"){
						error=error+1;
						errorString=errorString+"\n"+error+"：法條4340003需另單舉發。";
					}
				}
			}
		}
	}
<%else%>
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
<%end if%>
	if (error==0){
			getChkCarSimpleIDandRule();
	}else{
		alert(errorString);
	}
}
//檢查車種跟法條內容相不相符
function getChkCarSimpleIDandRule(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2=myForm.Rule2.value;
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	runServerScript("getChkCarSimpleIDandRule.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID);
}
function setChkCarSimpleIDandRule(RuleDetail){
	ErrorStr="";
	if (RuleDetail==1){
		ErrorStr="違規事實與簡式車種不符，請確認是否正確。";
	}
<%if sys_City="雲林縣" or sys_City="南投縣" or sys_City="屏東縣" then%>
	if (myForm.UseTool.value!="8"){
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
	<%if sys_City="台中市" then%>
	if (((myForm.Rule1.value.substr(0,2))=="55" || (myForm.Rule2.value.substr(0,2))=="55") && (myForm.UseTool.value!="8")){
		ErrorStr=ErrorStr+"\n第55條不可逕行舉發，請確認是否正確。";
	}
	<%end if %>
	<%if sys_City="南投縣" then%>
	if (eval(myForm.DealLineDate.value) < eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日小於填單日加"+getDealDateValue+"天，請確認是否正確。";
	}
	<%elseif sys_City="屏東縣" then%>
	if (eval(myForm.DealLineDate.value) < eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日小於"+getDealDateValue+"天，請確認是否正確。";
	}
	<%else%>
	if (eval(myForm.DealLineDate.value) != eval(Dyear+Dmonth+Dday)){
		ErrorStr=ErrorStr+"應到案日不是填單日加"+getDealDateValue+"天，請確認是否正確。";
	}
	<%end if%>
<%end if%>
	<%if sys_City="南投縣" then%>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) <= 10){
				ErrorStr=ErrorStr+"\n"+ErrorStr+"車速超過限速未超過10公里";
			}
		}				
	}
	<%end if%>
	
	if (ErrorStr!=""){
		if(confirm(ErrorStr+"\n是否確定要存檔？")){
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}else{
		document.myForm.kinds.value="DB_insert";
		document.myForm.submit();
	}
}
//是否為特殊用車
function getVIPCar(){
	strSpecUser=<%=trim(Session("SpecUser"))%>;
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			//alert("車牌格式錯誤，如該車輛無車牌或舊式車牌則可忽略此訊息！");
			//myForm.CarNo.select();
			runServerScript("getVIPCar.asp?CarID="+CarNum);
		}else{
			if (strSpecUser=="1"){
				runServerScript("getVIPCar.asp?CarID="+CarNum);
			}
		<%if sys_City<>"高雄市" and sys_City<>"苗栗縣" and sys_City<>"宜蘭縣" and sys_City<>"連江縣" then%>
			myForm.CarSimpleID.value=CarType;
		<%end if%>
			//myForm.CarSimpleID.select();
		}
	}else{
		Layer7.innerHTML=" ";
		myForm.CarSimpleID.value="";
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
		}
	}
}
//法條刪掉其他符號
function DelSpace1(){
	myForm.Rule1.value=myForm.Rule1.value.replace(/[^\d]/g,'');
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
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=myForm.RuleVerSion.value;
		myForm.Rule4.value="";
		runServerScript("getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw1();
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
	AutoGetRuleID(1);
}
//違規事實2(ajax)
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw2();
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
	AutoGetRuleID(2);
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
	TDLawNum=1;
	TDLaw1.innerHTML="違規法條三";
	TDLaw2.innerHTML="<input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw3()' alt='查詢法條'> <div id='Layer3' style='position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''>";

	<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DriverPID,MemberStation||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="宜蘭縣" or sys_City="嘉義市" then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="高雄市" then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City=ApconfigureCityName then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||RuleSpeed,IllegalSpeed||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台中市" then%>
		MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,IllegalZip||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="連江縣" then%>
		MoveTextVar("Billno1,ReportNo||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="苗栗縣" then%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||DriverPID,MemberStation||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%else%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate,JurgeDay||DealLineDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>

	myForm.Rule3.focus();
}
function OpenQueryLaw3(){
	window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes");
}
function AutoGetRuleID(LawOrder){	//按F5可以直接顯示相關法條
	if (event.keyCode==116){	
		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
		window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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
function getUnit(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (myForm.BillUnitID.value.length > 0){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}
//保管物品一(ajax)
function getFastener1(){
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
//保管物品三(ajax)
function getFastener3(){
	if (myForm.Fastener3.value.length == 1){
		var FastenerNum=myForm.Fastener3.value;
		runServerScript("getFastener.asp?FastenerOrder=3&FastenerID="+FastenerNum);
	}else if (myForm.Fastener3.value.length == 0){
		Layer10.innerHTML=" ";
		TDFastenerErrorLog3=0;
		myForm.Fastener3Val.value="";
	}else{
		Layer10.innerHTML=" ";
		TDFastenerErrorLog3=1;
		myForm.Fastener3Val.value="";
	}
}
//逕舉不一定要輸入固定桿編號. 除了是下方選擇使用固定桿
function getFixID(){
	if (myForm.UseTool.value.length == "1"){
		if (myForm.UseTool.value != "0" && myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3" && myForm.UseTool.value != "8" <%
	if sys_City="台南市"  then
		response.write " && myForm.UseTool.value != ""4"" && myForm.UseTool.value != ""5"""
	elseif sys_City="基隆市" then
		response.write " && myForm.UseTool.value != ""4"""
	end if 
		%>){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.select();
		}
	}
	if (myForm.UseTool.value == "1"){
		//Layer11.style.visibility = "visible";  
	}else{
		//Layer11.style.visibility = "hidden"; 
	}
}
//違規地點代碼(ajax)
function getillStreet(){
<%if sys_City<>"基隆市" and sys_City<>"彰化縣" then%>
	myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
<%end if%>
	if (event.keyCode==116){	
		event.keyCode=0;
		OstreetID=myForm.IllegalAddressID.value;

		window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}else if (myForm.IllegalAddressID.value.length >= 1){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
	
}
//舉發人一(ajax)
function getBillMemID1(){
<%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem1.value=myForm.BillMem1.value.toUpperCase();
	}
<%end if%>
	if (myForm.BillMem1.value.length > 1){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
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
	if (myForm.BillMem2.value.length > 1){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
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
	if (myForm.BillMem3.value.length > 1){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
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
	if (myForm.BillMem4.value.length > 1){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
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
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	if (myForm.UseTool.value!="8"){
	<%if sys_City<>"屏東縣" then%>
		getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
		myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday)
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getDealDateValue,BFillDate);
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
	<%end if%>
	}else{	//逕舉手開單+攔停天數
<%if (trim(Session("UnitLevelID"))<>"2" and sys_City="台中縣") or (sys_City<>"台中縣" and sys_City<>"高雄市") then%>
	<%if sys_City<>"基隆市" and sys_City<>"南投縣" and sys_City<>"台中縣" and sys_City<>"屏東縣" and sys_City<>"台中市" then%>
	<%if sys_City="台中縣" or sys_City="彰化縣" or sys_City="宜蘭縣" or sys_City="台南市" or sys_City="台東縣" or sys_City="嘉義市" or sys_City="嘉義縣" or sys_City="雲林縣" or sys_City="保二總隊四大隊二中隊" or sys_City="保二總隊三大隊二中隊" then%>
		getDealDateValue="45";
	<%elseif sys_City="澎湖縣" then%>
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
	if(TodayDate < myForm.IllegalDate.value){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}
<%if sys_City="屏東縣" or sys_City="澎湖縣" then%>
	if (myForm.UseTool.value=="8"){
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
function DateAdd(timeU,byMany,dateObj) {
	var millisecond=1;
	var second=millisecond*1000;
	var minute=second*60;
	var hour=minute*60;
	var day=hour*24;
	var year=day*365;

	var newDate;
	var dVal=dateObj.valueOf();
	switch(timeU) {
	case "ms": newDate=new Date(dVal+millisecond*byMany); break;
	case "s": newDate=new Date(dVal+second*byMany); break;
	case "mi": newDate=new Date(dVal+minute*byMany); break;
	case "h": newDate=new Date(dVal+hour*byMany); break;
	case "d": newDate=new Date(dVal+day*byMany); break;
	case "y": newDate=new Date(dVal+year*byMany); break;
	}
	return newDate;
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
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.RuleSpeed.value > 100){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：限速、限重超過100。";
	}
<%'if sys_City="台東縣" or sys_City="雲林縣" then%>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 61){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
			}
		}
	}
<%'else%>
//	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
//		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
//			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 60){
//				IntError=IntError+1;
//				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
//				IntError=IntError+1;
//				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
//			}
//		}
//	}
<%'end if%>
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
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > 100){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過100。";
	}
<%'if sys_City="台東縣" or sys_City="雲林縣" then%>
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 61){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
			}
		}
	}
<%'else%>
//	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
//		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
//			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= 60){
//				IntError=IntError+1;
//				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
//				IntError=IntError+1;
//				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
//			}
//		}
//	}
<%'end if%>
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

function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value!=""){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo.length != 9 ){
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}else if(myForm.Billno1.value != myForm.OldBillNo.value){
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
		}
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
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%if sys_City="台東縣" or sys_City="高雄縣" then%>
	}else if (event.keyCode==120){ //台東縣F9存檔
		event.keyCode=0;   
		InsertBillVase();
	}
<%else%>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	}
<%end if%>
}
function ProjectF5(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}/*else if (event.keyCode==13){	
		event.keyCode=0;
		OnKeyUpNextTag("CarAddID");
	}*/
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

//附加說明
function Add_LawPlus(){
	if (myForm.Rule1.value==""){
		alert("請先輸入違規法條一!!");
	}else{
		RuleID=myForm.Rule1.value;
	window.open("Query_LawPlus.asp?RuleID="+RuleID+"&theRuleVer=<%=theRuleVer%>","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
	}
}
//修改監理站回傳資料
function UpdateDciData(){
	window.open("BillKeyIn_Dci_Update.asp?BillSN="+myForm.BillSN.value+"&theUpdVer=1","WebPage1","left=20,top=10,location=0,width=900,height=600,resizable=yes,scrollbars=yes");
}
<%if sys_City="台南市" then%>
var sys_City="<%=sys_City%>";
function getDriverZip(obj,objName){
	if(obj.value!=''&&obj.value.length>2){
		runServerScript("getZipName.asp?ZipID="+obj.value+"&getZipName="+objName);
	}else if(obj.value!=''&&obj.value.length<3){
		alert("郵遞區號錯誤!!");
	}
}
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");

}
<%elseif sys_City="高雄市" Or sys_City="台中市" then%>
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
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="嘉義市" then%>
				myForm.RuleSpeed.select();
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
	}
<%
	if trim(rs1("CarSimpleID"))="" or isnull(rs1("CarSimpleID")) then
		if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
%>
		CarType=chkCarNoFormat(myForm.CarNo.value);
		//if (CarType!=0){
			myForm.CarSimpleID.value=CarType;
		//}
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
