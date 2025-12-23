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
<!--#include virtual="/traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size:11pt; 
	font-weight: bold;
	font-family: "標楷體";
}
.style2 {
	font-size:11pt; 
}
.style3 {
	font-size:11pt; 
	font-weight: bold;
}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:20px;
	font-family: "標楷體";
}
-->
</style>
<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<%	

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQLTemp=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp=strSQLTemp&" and a.ImageFileNameB='"&Right("0000000000000000"&trim(request("BillNo")),16)&"'"
	end if
	if trim(request("CarNo"))<>"" then
		strSQLTemp=strSQLTemp&" and a.CarNo='"&trim(request("CarNo"))&"'"
	end if
'	if trim(request("IllegalName"))<>"" then
'		strSQLTemp=strSQLTemp&" and (b.Owner='"&trim(request("IllegalName"))&"' or b.Driver='"&trim(request("IllegalName"))&"')"
'	end if
'	if trim(request("IllegalID"))<>"" then
'		strSQLTemp=strSQLTemp&" and (b.OwnerID='"&trim(request("IllegalID"))&"' or b.DriverID='"&trim(request("IllegalID"))&"' or a.DriverID='"&trim(request("IllegalID"))&"')"
'	end if
	if trim(request("BillSn"))<>"" then
		strSQLTemp=strSQLTemp&" and a.SN='"&trim(request("BillSn"))&"'"
	end If
	Cnt=0
	strSQLA="Select a.imagefilenameb" &_
		" from BillBase a" &_
		" where ((a.RecordStateID<>-1 and a.BillStatus='0')" &_
		" or a.BillStatus<>'0') and a.ImagePathName is not null  and a.BillNo is null "&strSQLTemp&" Group by imagefilenameb order by a.ImageFileNameB"
	Set rsArr1=conn.execute(strSQLA)
	If Not rsArr1.Bof Then
		rsArr1.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("查無資料！");
	window.close();
</script>	
<%
	end if
	While Not rsArr1.Eof
	If IsNull(rsArr1("imagefilenameb")) Then
		STRimagefilenameb=" and imagefilenameb is Null"
	Else
		STRimagefilenameb=" and imagefilenameb='"&Trim(rsArr1("imagefilenameb"))&"'"
	End If 
	strSQL="Select a.imagefilenameb,a.BillNo,a.Sn,a.CarNo,a.BillTypeID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.MemberStation,a.EquipMentID" &_
		",a.Recorddate,a.RecordMemberID,a.RecordStateID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2" &_
		",a.BillMem2,a.BillMemID3,a.BillMem3" &_
		",a.BillMemID4,a.BillMem4,a.RuleVer,a.IllegalAddressID,a.IllegalAddress,a.BillFillDate,a.BillUnitID" &_
		",a.DealLineDate,a.Note,a.CarSimpleID,a.OwnerAddress,a.OwnerZip,a.DriverAddress,a.DriverZip,a.Owner,a.OwnerAddress,a.OwnerZip,a.ImageFileNameB" &_
		" from BillBase a" &_
		" where ((a.RecordStateID<>-1 and a.BillStatus='0')" &_
		" or a.BillStatus<>'0') and a.ImagePathName is not null  and a.BillNo is null "&strSQLTemp&STRimagefilenameb&" order by a.ImageFileNameB"

		'response.write strSQL
		'response.end

%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	
	set rs1=conn.execute(strSQL)
	If Not rs1.eof Then	

	if Cnt>0 then
%>
<div class="PageNext"></div>
<%	end if
	
	Cnt=Cnt+1
	StationNameBillBase=trim(rs1("MemberStation"))
	'--------------------------------------BILLBASEDCIRETURN------------------------------------
'先查有沒有車籍查尋的資料 沒有的話再用入案資料
	StationName=""	'到案處所
	IllegalMemID=""	'違規人證號
	IllegalMem=""	'違規人姓名
	IllegalAddress=""	'違規人地址
	OwnerName=""	'車主姓名
	OwnerAddress=""	'車主地址
	DciCarTypeID=""	'詳細車種代碼
	DciCarType=""	'詳細車種
	strDciA="select * from BillBaseDciReturn where (BillNo='"&trim(rs1("BillNo"))&"' or BillNo is Null)" &_
			" and CarNo='"&trim(rs1("CarNo"))&"'" &_
			" and ExchangeTypeID='A' and Status='S'"
	set rsDciA=conn.execute(strDciA)
	if not rsDciA.eof then

		if sys_City<>"台中市" then
			OwnerZipName=""
			DriverZipName=""
		else
			strOZip="select ZipName from Zip where ZipID='"&trim(rsDciA("OwnerZip"))&"'"
			set rsOZip=conn.execute(strOZip)
			if not rsOZip.eof then
				OwnerZipName=trim(rsOZip("ZipName"))
			end if
			rsOZip.close
			set rsOZip=nothing

			strDZip="select ZipName from Zip where ZipID='"&trim(rsDciA("DriverHomeZip"))&"'"
			set rsDZip=conn.execute(strDZip)
			if not rsDZip.eof then
				DriverZipName=trim(rsDZip("ZipName"))
			end if
			rsDZip.close
			set rsDZip=nothing
		end if

		StationNameDci=trim(rsDciA("DciReturnStation"))
			
		OwnerName=trim(rsDciA("Owner"))
		OwnerAddress=trim(rsDciA("OwnerZip"))&" "&trim(rsDciA("OwnerAddress"))
		DciCarTypeID=trim(rsDciA("DciReturnCarType"))
		if trim(rs1("BillTypeID"))="1" then
			IllegalMemID=trim(rsDciA("DriverID"))
			IllegalMem=trim(rsDciA("Driver"))
			IllegalAddress=trim(rsDciA("DriverHomeZip"))&" "&DriverZipName&trim(rsDciA("DriverHomeAddress"))
		else
			IllegalMemID=trim(rsDciA("OwnerID"))
			IllegalMem=trim(rsDciA("Owner"))
			IllegalAddress=trim(rsDciA("OwnerZip"))&" "&OwnerZipName&trim(rsDciA("OwnerAddress"))
		end if
	else
		strDciB="select a.* from BillBaseDciReturn a,DciReturnStatus b" &_
			" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
			" and (a.BillNo='"&trim(rs1("BillNo"))&"' or a.BillNo is Null)" &_
			" and a.CarNo='"&trim(rs1("CarNo"))&"'" &_
			" and b.DciReturnStatus=1 and ExchangeTypeID='W'"
		set rsDciB=conn.execute(strDciB)
		if not rsDciB.eof then

			if sys_City<>"台中市" then
				OwnerZipName=""
				DriverZipName=""
			else
				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciB("OwnerZip"))&"'"
				set rsOZip=conn.execute(strOZip)
				if not rsOZip.eof then
					OwnerZipName=trim(rsOZip("ZipName"))
				end if
				rsOZip.close
				set rsOZip=nothing

				strDZip="select ZipName from Zip where ZipID='"&trim(rsDciB("DriverHomeZip"))&"'"
				set rsDZip=conn.execute(strDZip)
				if not rsDZip.eof then
					DriverZipName=trim(rsDZip("ZipName"))
				end if
				rsDZip.close
				set rsDZip=nothing
			end if
			if trim(rs1("BillTypeID"))="2" then
				StationName=trim(rsDciB("DciReturnStation"))
			else
				StationName=trim(rs1("MemberStation"))
			end if
			OwnerName=trim(rsDciB("Owner"))
			OwnerAddress=trim(rsDciB("OwnerZip"))&" "&OwnerZipName&trim(rsDciB("OwnerAddress"))
			DciCarTypeID=trim(rsDciB("DciReturnCarType"))
			if trim(rs1("BillTypeID"))="1" then
				IllegalMemID=trim(rsDciB("DriverID"))
				IllegalMem=trim(rsDciB("Driver"))
				IllegalAddress=trim(rsDciB("DriverHomeZip"))&DriverZipName&" "&trim(rsDciB("DriverHomeAddress"))
			else
				IllegalMemID=trim(rsDciB("OwnerID"))
				IllegalMem=trim(rsDciB("Owner"))
				IllegalAddress=trim(rsDciB("OwnerZip"))&" "&OwnerZipName&trim(rsDciB("OwnerAddress"))
			end if
		end if
		rsDciB.close
		set rsDciB=nothing
	end if
	rsDciA.close
	set rsDciA=Nothing
	
	If sys_City="花蓮縣" Or sys_City="台東縣" Then
		If Not isnull(rs1("Owner")) Then
			IllegalMem=trim(rs1("Owner"))
			OwnerName=trim(rs1("Owner"))
		End If
		If Not isnull(rs1("OwnerAddress")) then
			IllegalAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
			OwnerAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
		End if
	End If

	strCarType="select Content from DciCode where TypeID=5 and ID='"&DciCarTypeID&"'"
	set rsCarType=conn.execute(strCarType)
	if not rsCarType.eof then
		DciCarType=trim(rsCarType("Content"))
	end if
	rsCarType.close
	set rsCarType=nothing

	CaseInDate=""	'入案日期
	CaseStatus=""	'入案狀態
	DciFileName=""	'入案檔名
	DciBatchNumber=""	'入案批號
	strCaseIn="select a.*,c.* from BillBaseDciReturn a,DciReturnStatus b,DciLog c" &_
			" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
			" and a.ExchangeTypeID=c.ExchangeTypeID and a.Status=c.DciReturnStatusID" &_
			" and a.BillNo=c.BillNo and a.CarNo=c.CarNo" &_
			" and a.BillNo='"&trim(rs1("BillNo"))&"' " &_
			" and a.CarNo='"&trim(rs1("CarNo"))&"' and a.ExchangeTypeID='W'" &_
			" order by c.ExchangeDate Desc"
	set rsCaseIn=conn.execute(strCaseIn)
	if not rsCaseIn.eof then
		CaseInDate=trim(rsCaseIn("DciCaseInDate"))
		if trim(rsCaseIn("STATUS"))<>"" and not isnull(rsCaseIn("STATUS")) then
			strStuts="select StatusContent from DciReturnStatus where DciActionID='W' and DciReturn='"&trim(rsCaseIn("STATUS"))&"'"
			set rsStuts=conn.execute(strStuts)
			if not rsStuts.eof then
				CaseStatus=trim(rsStuts("StatusContent"))
			end if
			rsStuts.close
			set rsStuts=nothing
		else
			CaseStatus="未處理"
		end if
		DciFileName=trim(rsCaseIn("FileName"))
		DciBatchNumber=trim(rsCaseIn("BatchNumber"))
	else
		CaseStatus="未上傳"
	end if
	rsCaseIn.close
	set rsCaseIn=nothing

'-----------------------------------BillMailHistory-------------------------------------
	StoreAndSendFlag=0	'是否做過寄存

	MailDate=""	'郵寄日期
	MailNumber=""	'郵寄序號
	MailStation=""	'寄存郵局
	GetFileName=""	'收受檔案
	GetBatchNumber=""	'收受批號
	GetStatus=""	'收受上傳狀態
	GetMailDate=""	'收受日期
	GetMailReason=""	'收受原因
	ReturnMailDate=""	'退回日期
	ReturnReason=""	'退件原因
	ReturnSendDate=""	'移送日期
	ReturnMailNumber=""	'退件郵寄序號
	ReturnSendMailDate=""	'退件郵寄日期
	StoreAndSendGovNumber=""	'寄存送達書號
	StoreAndSendEffectDate=""	'寄存送達日
	StoreAndSendEndDate=""	'寄存送達生效(完成)日
	OpenGovGovNumber=""	'公示送達書號
	OpenGovEffectDate=""	'公示送達生效日
	StoreAndSendDate=""	'二次送達日期
	StoreAndSendReason=""	'二次送達原因
	BillMailNo=""	'郵寄序號
	ReturnMailNo=""	'退件郵寄序號
	MailCheckNumber="" '郵局查詢號
	MailReturnCheckNumber="" '單退後投遞郵局查詢號
	SignMan=""	'簽收人
	'檢查是單退還是收受
	strCheck="select count(*) as cnt from Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7'"
	set rsCheck=conn.execute(strCheck)
	if not rsCheck.eof then
		if rsCheck("cnt")="0" then
			CheckFlag=0	'單退
		else
			CheckFlag=1	'收受
		end if
	end if
	rsCheck.close
	set rsCheck=nothing

	strMail="select * from StopBillMailHistory where BillNo='"&trim(rs1("ImageFileNameB"))&"' and CarNo='"&trim(rs1("CarNo"))&"'"
	set rsMail=conn.execute(strMail)
	if not rsMail.eof then
		if trim(rs1("BillTypeID"))="2" or (trim(rs1("BillTypeID"))="1" and trim(rs1("EquipMentID"))="1") then
			if trim(rsMail("MailDate"))<>"" and not isnull(rsMail("MailDate")) then
				MailDate=gArrDT(trim(rsMail("MailDate")))
			end if
		end if
		MailNumber=trim(rsMail("MailNumber"))
		if CheckFlag=0 then
			if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
				ReturnMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
			end if
			GetMailDate=""
		else
			if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
				GetMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
			end if
			ReturnMailDate=""
		end if
		'退件or收受原因
		if CheckFlag=0 then
			if trim(rsMail("RETURNRESONID"))<>"" and not isnull(rsMail("RETURNRESONID")) then
				strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("RETURNRESONID"))&"'"
				set rsRR=conn.execute(strReturnReason)
				if not rsRR.eof then
					ReturnReason=trim(rsRR("Content"))
				end if
				rsRR.close
				set rsRR=nothing
			end if
			GetMailReason=""
			GetFileName=""
			GetBatchNumber=""
			GetStatus="未上傳"
		else
			if trim(rsMail("RETURNRESONID"))<>"" and not isnull(rsMail("RETURNRESONID")) then
				strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("RETURNRESONID"))&"'"
				set rsRR=conn.execute(strReturnReason)
				if not rsRR.eof then
					GetMailReason=trim(rsRR("Content"))
				end if
				rsRR.close
				set rsRR=nothing
			end if
			ReturnReason=""
			if trim(rsMail("SignMan"))<>"" and not isnull(rsMail("SignMan")) then
				SignMan=trim(rsMail("SignMan"))
			end if
			strGet="select * from Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7' order by ExchangeDate desc"
			set rsGet=conn.execute(strGet)
			if not rsGet.eof then
				GetFileName=trim(rsGet("FileName"))
				GetBatchNumber=trim(rsGet("BatchNumber"))
				if trim(rsGet("DciReturnStatusID"))<>"" and not isnull(rsGet("DciReturnStatusID")) then
					strGStuts="select StatusContent from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsGet("DciReturnStatusID"))&"'"
					set rsGStuts=conn.execute(strGStuts)
					if not rsGStuts.eof then
						GetStatus=trim(rsGStuts("StatusContent"))
					end if
					rsGStuts.close
					set rsGStuts=nothing
				else
					GetStatus="未處理"
				end if
			end if
			rsGet.close
			set rsGet=nothing
		end if
		if trim(rsMail("MailStation"))<>"" and not isnull(rsMail("MailStation")) then
			MailStation=trim(rsMail("MailStation"))
		end if
		if trim(rsMail("SendOpenGovDocToStationDate"))<>"" and not isnull(rsMail("SendOpenGovDocToStationDate")) then
			ReturnSendDate=left(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-4)&"-"&mid(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-3,2)&"-"&mid(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-1,2)
		end if
		ReturnMailNumber=trim(rsMail("StoreAndSendMailNumber"))
		if trim(rsMail("StoreAndSendSendDate"))<>"" and not isnull(rsMail("StoreAndSendSendDate")) then
			ReturnSendMailDate=gArrDT(trim(rsMail("StoreAndSendSendDate")))
		end if
		if trim(rsMail("STOREANDSENDGOVNUMBER"))<>"" and not isnull(rsMail("STOREANDSENDGOVNUMBER")) then
			StoreAndSendGovNumber=trim(rsMail("STOREANDSENDGOVNUMBER"))
		end if
		if trim(rsMail("STOREANDSENDEFFECTDATE"))<>"" and not isnull(rsMail("STOREANDSENDEFFECTDATE")) then
			StoreAndSendEffectDate=gArrDT(trim(rsMail("STOREANDSENDEFFECTDATE")))
		end if
		if trim(rsMail("StoreAndSendMailDate"))<>"" and not isnull(rsMail("StoreAndSendMailDate")) then
			StoreAndSendEndDate=gArrDT(trim(rsMail("StoreAndSendMailDate")))
		end if
		if trim(rsMail("OPENGOVNUMBER"))<>"" and not isnull(rsMail("OPENGOVNUMBER")) then
			OpenGovGovNumber=trim(rsMail("OPENGOVNUMBER"))
		end if
		if trim(rsMail("OPENGOVDATE"))<>"" and not isnull(rsMail("OPENGOVDATE")) then
			OpenGovEffectDate=gArrDT(trim(rsMail("OPENGOVDATE")))
		end if
		if trim(rsMail("STOREANDSENDMAILRETURNDATE"))<>"" and not isnull(rsMail("STOREANDSENDMAILRETURNDATE")) then
			StoreAndSendDate=gArrDT(trim(rsMail("STOREANDSENDMAILRETURNDATE")))
		end if
		if trim(rsMail("STOREANDSENDRETURNRESONID"))<>"" and not isnull(rsMail("STOREANDSENDRETURNRESONID")) then
			strSReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("STOREANDSENDRETURNRESONID"))&"'"
			set rsSR=conn.execute(strSReason)
			if not rsSR.eof then
				StoreAndSendReason=trim(rsSR("Content"))
			end if
			rsSR.close
			set rsSR=nothing
		end if
		if trim(rsMail("MailSeqNo1"))<>"" and not isnull(rsMail("MailSeqNo1")) then
			BillMailNo=trim(rsMail("MailSeqNo1"))
		end if
		if trim(rsMail("MailSeqNo2"))<>"" and not isnull(rsMail("MailSeqNo2")) then
			ReturnMailNo=trim(rsMail("MailSeqNo2"))
		end if
		if trim(rsMail("ReturnResonID"))<>"" and not isnull(rsMail("ReturnResonID")) then
			if trim(rsMail("ReturnResonID"))="5" or trim(rsMail("ReturnResonID"))="6" or trim(rsMail("ReturnResonID"))="7" or trim(rsMail("ReturnResonID"))="T" then
				StoreAndSendFlag=1
			end if
		end if
		if trim(rsMail("MailChkNumber"))<>"" and not isnull(rsMail("MailChkNumber")) then
			MailCheckNumber=trim(rsMail("MailChkNumber"))
		end if
		if trim(rsMail("OpenGovReportNumber"))<>"" and not isnull(rsMail("OpenGovReportNumber")) then
			MailReturnCheckNumber=trim(rsMail("OpenGovReportNumber"))
		end if
	end if
	rsMail.close
	set rsMail=nothing



	'----------------------smith 增加催繳的顯示------------------------------
If trim(request("BillSn"))<>"" then
	strSTOPsql="select * from StopBillMailHistory where BillSn='" & trim(request("BillSn")) & "'"
Else
	strSTOPsql="select * from StopBillMailHistory where BillSn='" & trim(rs1("SN")) & "'"
End If 
  set rsOpenGov=conn.execute(strSTOPsql)
	if not rsOpenGov.eof then
		'判斷註記的原因種類
		if trim(rsOpenGov("ReturnResonID"))<>"" and not isnull(rsOpenGov("ReturnResonID")) then
			'寄存送達
			if trim(rsOpenGov("ReturnResonID"))="5" or trim(rsOpenGov("ReturnResonID"))="6" or trim(rsOpenGov("ReturnResonID"))="7" or trim(rsOpenGov("ReturnResonID"))="T" then
				UserMarkFlag=1
			end if
			'公示送達
			if trim(rsOpenGov("ReturnResonID"))="1" or trim(rsOpenGov("ReturnResonID"))="2" or trim(rsOpenGov("ReturnResonID"))="3" or trim(rsOpenGov("ReturnResonID"))="4" or trim(rsOpenGov("ReturnResonID"))="8" or trim(rsOpenGov("ReturnResonID"))="M" or trim(rsOpenGov("ReturnResonID"))="K" or trim(rsOpenGov("ReturnResonID"))="L" or trim(rsOpenGov("ReturnResonID"))="O" or trim(rsOpenGov("ReturnResonID"))="P" or trim(rsOpenGov("ReturnResonID"))="Q"  then
				UserMarkFlag=2
			end if		
			'收受註記
			if trim(rsOpenGov("ReturnResonID"))="A" or trim(rsOpenGov("ReturnResonID"))="B" or trim(rsOpenGov("ReturnResonID"))="C"  then
				UserMarkFlag=3
			end if						
		end if
		strStoreReturnDate=""
		strEffectDate=""
		strOpenGovReturnDate=""
		strOpenGovDate=""
		strUserGetDate=""
		ReturnReason=""
		SignMan=""
		  '取得原因和簽收人
			if trim(rsOpenGov("RETURNRESONID"))<>"" and not isnull(rsOpenGov("RETURNRESONID")) then
				strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsOpenGov("RETURNRESONID"))&"'"
				set rsRR=conn.execute(strReturnReason)
				if not rsRR.eof then
					ReturnReason=trim(rsRR("Content"))
				end if
				rsRR.close
				set rsRR=nothing
			end if
			if trim(rsOpenGov("SignMan"))<>"" and not isnull(rsOpenGov("SignMan")) then
				SignMan=trim(rsOpenGov("SignMan"))
			end if
					
		strUserMarkReason=""
		strOpenGovReason=""
		strStoreReason=""
		if UserMarkFlag=1 then 
			strStoreReturnDate=rsOpenGov("MailReturnDate")		
			strEffectDate=rsOpenGov("StoreAndSendEffectDate")		
			strStoreReason=ReturnReason	
		elseif UserMarkFlag=2 then
			strOpenGovReturnDate=rsOpenGov("MailReturnDate")
			strOpenGovDate=rsOpenGov("OpenGovDate")
			strOpenGovReason=ReturnReason
		elseif UserMarkFlag=3 then
			strUserGetDate=rsOpenGov("MailReturnDate")
			strUserMarkReason=ReturnReason
		end if
	end If
	
	If Not IsNull(rs1("imagefilenameb")) then
		if trim(rs1("BillTypeID"))="1" Then
			If Not IsNull(rs1("DriverAddress")) Then
				IllegalAddress=Trim(rs1("DriverZip"))&" "&Trim(rs1("DriverAddress"))
			End if
		else
			If Not IsNull(rs1("OwnerAddress")) Then
				IllegalAddress=Trim(rs1("OwnerZip"))&" "&Trim(rs1("OwnerAddress"))
			End if
		end if
		If Not IsNull(rs1("OwnerAddress")) Then
			OwnerAddress=Trim(rs1("OwnerZip"))&" "&Trim(rs1("OwnerAddress"))
		End if
	End If
	'-----------------------------------------------------------------------------------------------
	

%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td align="center">
				<span class="style6">停車管理催繳單</span>
			</td>
		</tr>
		<tr>
			<td><span class="style2">製表單位：</span><span class="style1"><%
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(session("Unit_ID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">操作人：</span><span class="style1"><%
			strMem="select ChName from MemberData where MemberID='"&trim(session("User_ID"))&"'"
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">製表時間：</span><span class="style3"><%=now%></span></td>
		</tr>
	</table>
	<hr>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">催繳單號：</span><span class="style1"><%
			if trim(rs1("imagefilenameb"))<>"" and not isnull(rs1("imagefilenameb")) then
				response.write trim(rs1("imagefilenameb"))
			end if
			%></span></td>
			<td width="27%"><%
			if trim(rs1("BillTypeID"))="2" then
				StationName=StationNameDci
			else
				StationName=StationNameBillBase
			end if
			strStation="select * from Station where DciStationID='"&StationName&"'"
			set rsStation=conn.execute(strStation)
			if not rsStation.eof then
				'response.write trim(rsStation("DCIStationName"))
			end if
			rsStation.close
			set rsStation=nothing
			%></td>
			<td width="23%"><span class="style2">催繳類別：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))="2" then
				response.write "停管催繳"
			else
				response.write "停管催繳"
			end if
			%></span></td>
			<td width="25%"><span class="style2">催繳單狀態：</span><span class="style1"><%
			if trim(rs1("RecordStateID"))="-1" then
				response.write "<font color=""red"">已刪除</font>"
			else
				response.write "正常"
			end if
			%></span></td>
		</tr>
		<tr>
			<%
			if CaseInDate<>"" and not isnull(CaseInDate) then
				'response.write left(CaseInDate,len(CaseInDate)-4)&"-"&mid(CaseInDate,len(CaseInDate)-3,2)&"-"&mid(CaseInDate,len(CaseInDate)-1,2)
			end if
			%>
			<td><span class="style2">是否郵寄：</span><span class="style1"><%
			if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write "是"
				else
					response.write "否"
				end if
			end if	
			%></span></td>
			<%
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
				set rsMem1=conn.execute(strMem1)
				if not rsMem1.eof then
					'response.write "("&trim(rsMem1("LoginID"))&")"
				end if
				rsMem1.close
				set rsMem1=nothing
			end if	
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "/&nbsp;"&trim(rs1("BillMem2"))
				strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
				set rsMem2=conn.execute(strMem2)
				if not rsMem2.eof then
					'response.write "("&trim(rsMem2("LoginID"))&")"
				end if
				rsMem2.close
				set rsMem2=nothing
			end if	
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "/&nbsp;"&trim(rs1("BillMem3"))
				strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
				set rsMem3=conn.execute(strMem3)
				if not rsMem3.eof then
					'response.write "("&trim(rsMem3("LoginID"))&")"
				end if
				rsMem3.close
				set rsMem3=nothing
			end if	
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "/&nbsp;"&trim(rs1("BillMem4"))
				strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
				set rsMem4=conn.execute(strMem4)
				if not rsMem4.eof then
					'response.write "("&trim(rsMem4("LoginID"))&")"
				end if
				rsMem4.close
				set rsMem4=nothing
			end if	
			%>
		</tr>
		<%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				if left(trim(rs1("Rule1")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end if
				end if
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					'response.write trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing

				if trim(rs1("BillTypeID"))="2" and trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					'response.write "&nbsp;"&trim(rs1("Rule4"))
				end if
			end if	
			%>
<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
<%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if left(trim(rs1("Rule2")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					'response.write trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%>
<%end if%>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
		<%
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if left(trim(rs1("Rule3")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					'response.write trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))<>"2" then%>
<%
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if left(trim(rs1("Rule4")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					'response.write trim(rs1("Rule4"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%>
<%end if%>
		<%'titan
		strSQL="select illegaldate,illegaladdressID,illegalAddress from billbase where CarNo='"&trim(rs1("CarNo"))&"' and ImagePathName is not null  "
		if not ifnull(rs1("ImagefilenameB")) then
			strIlg=" and imagefilenameB='"&trim(rs1("imagefilenameb"))&"'"
		else
			strIlg=" and imagefilenameB is null"
		end if
		
		set rsilg=conn.execute(strSQL&strIlg&" order by illegaldate")
		while not rsilg.eof%>
		<tr>
			<td><span class="style2">停車時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rsilg("IllegalDate")) then
				response.write gArrDT(trim(rsilg("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rsilg("IllegalDate")),2)&":"
				response.write Right("00"&minute(rsilg("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">停車路段：</span><span class="style1"><%
			response.write trim(rsilg("IllegalAddressID"))&" "&trim(rsilg("IllegalAddress"))
			%></span></td>			
		</tr><%
			rsilg.movenext
		wend
		rsilg.close
		%>
		<tr>
			<td><span class="style2">郵寄日期：</span><span class="style1"><%=MailDate%></span></td>
			<td><span class="style2">郵寄序號：</span><span class="style1"><%
			if sys_City<>"台南縣" and sys_City<>"台南市" then
				response.write MailNumber
			else
				response.write BillMailNo
			end if
			%></span></td>

		</tr>
		<tr>
		<%'if sys_City<>"台東縣" then%>
			<td><span class="style2">停車人證號：</span><span class="style1"><%=IllegalMemID%></span></td>
		<%'End If %>
			<td><span class="style2">停車人姓名：</span><span class="style1"><%=funcCheckFont(IllegalMem,20,1)%></span></td>
			<td colspan="3"><span class="style2">停車人住址：</span><span class="style1"><%=funcCheckFont(IllegalAddress,20,1)%></span></td>
		</tr>
		<tr>
			<td><span class="style2">車號：</span><span class="style1"><%=trim(rs1("CarNo"))%></span></td>
			<td><span class="style2">車主姓名：</span><span class="style1"><%=funcCheckFont(OwnerName,20,1)%></span></td>
			<td colspan="3"><span class="style2">車主住址：</span><span class="style1"><%=funcCheckFont(OwnerAddress,20,1)%></span></td>
		</tr>
		<tr>
			<td><span class="style2">填單日期：</span><span class="style1"><%
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			end if	
			%></span></td>
			<td><span class="style2">詳細車種：</span><span class="style1"><%=DciCarType%></span></td>
			<td colspan="3"><span class="style2">催繳單位：</span><span class="style1"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				response.write trim(rs1("BillUnitID"))&"&nbsp;"
				strBillUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsBillUnit=conn.execute(strBillUnit)
				if not rsBillUnit.eof then
					response.write trim(rsBillUnit("UnitName"))
				end if
				rsBillUnit.close
				set rsBillUnit=nothing
			end if	
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">繳費期限：</span><span class="style1"><%
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			end if	
			%></span></td>
			<td><span class="style2">註記日期：</span><span class="style1"><%
'			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
'				response.write gArrDT(trim(rs1("RecordDate")))
'			end if
			strSQL="select UserMarkDate from stopbillmailhistory where billsn="&rs1("sn")

			set rsmail=conn.execute(strSQL)
			If not rsmail.eof Then
				Response.Write gArrDT(trim(rsmail("usermarkdate")))
			End if
			rsmail.close
			%></span></td>
			<td><span class="style2">操作人員：</span><span class="style1"><%
			strRecMem="select ChName from MemberData where MemberID='"&trim(rs1("RecordMemberID"))&"'"
			set rsRecMem=conn.execute(strRecMem)
			if not rsRecMem.eof then
				response.write trim(rsRecMem("ChName"))
			end if
			rsRecMem.close
			set rsRecMem=nothing
			%></span></td>
			
		</tr>
	</table>
	<hr>
	<%

	strDCILog="select * from DciLog where BillSN="&trim(rs1("SN"))&" and exchangetypeid='A' order by ExchangeDate Desc"
	i=0
	set rsCity=conn.execute(strDCILog)
	If Not rsCity.Bof Then
		sBatchNumber=rsCity("BatchNumber")
		sFileName=rsCity("FileName")
	end if
	set rsCity=nothing
%>

	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td colspan="2"><span class="style2">車籍查詢 批號：</span><span class="style1"><%=sBatchNumber%></span></td>
			<td><span class="style2">車籍查詢 檔名：</span><span class="style1"><%=sFileName%></span></td>
		</tr>
			
		<tr>
			<td colspan="2"><span class="style2">簽收日期：</span><span class="style1"><%=strUserGetDate%></span></td>
			<td><span class="style2">簽收原因：</span><span class="style1"><%
			response.write strUserMarkReason
			%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">簽收人：</span><span class="style1"><%
			response.write SignMan
			%></span>
			</td>
		</tr>
		<tr>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">寄存送達 單退日：</span><span class="style1"><%=strStoreReturnDate%></span></td>			
			<!--<td><span class="style2">寄存送達 退件原因：</span><span class="style1">-->
			<!--<%=strStoreReason%>-->
			<!--</span></td> -->
			<td><span class="style2">寄存送達日：</span><span class="style1"><%=strEffectDate%></span></td>
		
		</tr>
		<tr>
			<td colspan="2"><span class="style2">郵局：</span><span class="style1"><%=MailStation%></span></td>		
		</tr>
		<tr>
			<td colspan="2"><span class="style2">公示送達 單退日：</span><span class="style1"><%=strOpenGovReturnDate%></span></td>		
			<td><span class="style2">公示送達 退件原因：</span><span class="style1"><%=strOpenGovReason%></span></td>	
		</tr>
		
		<tr>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">公示送達 公告日：</span><span class="style1"><%=strOpenGovDate%></span></td>		
			<td><span class="style2"></span></td>	
		</tr>
		
		<tr>
			<td colspan="3"><span class="style2">停管催繳號 & 催繳檔名：</span><span class="style1"><%=trim(rs1("Note"))%></span></td>
		</tr>
	
	</table>


<%	End if
	rs1.close
	Set rs1=Nothing
	
	rsArr1.MoveNext
	Wend
	rsArr1.close
	set rsArr1=nothing
%>
<Div id="Layer111" style="width:1041px; height:24px; ">
  <div align="center">
  <input type="button" value="列印" onclick="DP();">
  <br>
    (若無列印鈕，可按下滑鼠右鍵選擇列印功能，格式為A4橫印)
  </div>
</Div>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgFileName){
	urlstr='../ProsecutionImage/ProsecutionImageDetail.asp?FileName='+ImgFileName.replace(/\+/g,'@2@')+'&SN=1';
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
function DP(){
	window.focus();
<%if Cnt=1 then%>
	Layer112.style.visibility="hidden";
<%end if%>
	Layer111.style.visibility="hidden";
	window.print();
	window.close();
}
</script>
</html>
