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
	Server.ScriptTimeout = 65000
	Response.flush

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQLTemp=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp=strSQLTemp&" and a.BillNo='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("CarNo"))<>"" then
		strSQLTemp=strSQLTemp&" and a.CarNo like '%"&trim(request("CarNo"))&"%'"
	end if
	if trim(request("IllegalName"))<>"" then
		strSQLTemp=strSQLTemp&" and (b.Owner='"&trim(request("IllegalName"))&"' or b.Driver='"&trim(request("IllegalName"))&"')"
	end if
	if trim(request("IllegalID"))<>"" then
		strSQLTemp=strSQLTemp&" and (b.OwnerID='"&trim(request("IllegalID"))&"' or b.DriverID='"&trim(request("IllegalID"))&"' or a.DriverID='"&trim(request("IllegalID"))&"')"
	end if
	if trim(request("BillSn"))<>"" then
		strSQLTemp=strSQLTemp&" and a.SN='"&trim(request("BillSn"))&"'"
	end if
	strSQL="Select a.BillNo,a.Sn,a.CarNo,a.BillTypeID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.MemberStation,a.EquipMentID" &_
		",a.Recorddate,a.RecordMemberID,a.RecordStateID,a.IllegalDate,a.BillMemID1,a.BillMem1" &_
		",a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3" &_
		",a.BillMemID4,a.BillMem4,a.RuleVer,a.IllegalAddressID,a.IllegalAddress,a.BillFillDate,a.BillUnitID" &_
		",a.DealLineDate,a.Note,a.CarSimpleID from BillBase a,BillBaseDciReturn b" &_
		" where ((a.RecordStateID<>-1 and a.BillStatus='0')" &_
		" or a.BillStatus<>'0') and a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.ExChangeTypeID='W'" &_
		" and b.Status in ('Y','S','n') "&strSQLTemp&" order by a.Recorddate desc"

		'response.write strSQL
		'response.end

%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	Cnt=0
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("查無資料！");
	window.close();
</script>	
<%
	end if
	While Not rs1.Eof
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
	if not rsDciA.eof and trim(rs1("BillTypeID"))="2" then

		if sys_City<>"台中市" then
			OwnerZipName=""
			DriverZipName=""
		else
			if trim(rsDciA("NwnerZip"))<>"" and not isnull(rsDciA("NwnerZip")) then
				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciA("NwnerZip"))&"'"
				set rsOZip=conn.execute(strOZip)
				if not rsOZip.eof then
					OwnerZipName=trim(rsOZip("ZipName"))
				end if
				rsOZip.close
				set rsOZip=nothing
			else
				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciA("OwnerZip"))&"'"
				set rsOZip=conn.execute(strOZip)
				if not rsOZip.eof then
					OwnerZipName=trim(rsOZip("ZipName"))
				end if
				rsOZip.close
				set rsOZip=nothing
			end if

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
			IllegalAddress=trim(rsDciA("DriverHomeZip"))&" "&trim(rsDciA("DriverHomeAddress"))
		else
			if trim(rsDciA("Nwner"))<>"" and not isnull(rsDciA("Nwner")) then
				IllegalMemID=trim(rsDciA("NwnerID"))
				IllegalMem=trim(rsDciA("Nwner"))
				IllegalAddress=trim(rsDciA("NwnerZip"))&" "&trim(rsDciA("NwnerAddress"))
			else
				IllegalMemID=trim(rsDciA("OwnerID"))
				IllegalMem=trim(rsDciA("Owner"))
				IllegalAddress=trim(rsDciA("OwnerZip"))&" "&trim(rsDciA("OwnerAddress"))
			end if
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
	set rsDciA=nothing

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
			" and c.BillSn='"&trim(rs1("Sn"))&"' " &_
			" and a.ExchangeTypeID='W'" &_
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
		if sys_City<>"台中市" then
			CaseStatus="未上傳"
		else
			CaseStatus="&nbsp;"
		end if 
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
	StoreAndSendFinalMailDate=""	'送達證書郵寄日期
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

	strMail="select * from BillMailHistory where BillSn='"&trim(rs1("Sn"))&"'"
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
			if sys_City<>"台中市" then
				GetStatus="未上傳"
			else
				GetStatus="&nbsp;"
			end if 
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
		'送達證書郵寄日期
		if sys_City="基隆市" then
			if trim(rsMail("StoreAndSendFinalMailDate"))<>"" and not isnull(rsMail("StoreAndSendFinalMailDate")) then
				StoreAndSendFinalMailDate=gArrDT(trim(rsMail("StoreAndSendFinalMailDate")))
			end if
		end if
	end if
	rsMail.close
	set rsMail=nothing

'-----------------------------DciLog退件-----------------------------
	ReturnFileName=""	'退件上傳檔名
	ReturnBatchNumber=""	'退件批號
	ReturnStatus=""	'退件上傳狀態
	strReturn="select * from DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='3'" &_
		" order by ExchangeDate desc"
	set rsReturn=conn.execute(strReturn)
	if not rsReturn.eof then
		ReturnFileName=trim(rsReturn("FileName"))
		ReturnBatchNumber=trim(rsReturn("BatchNumber"))
		if trim(rsReturn("DciReturnStatusID"))<>"" and not isnull(rsReturn("DciReturnStatusID")) then
			strRStuts="select StatusContent from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsReturn("DciReturnStatusID"))&"'"
			set rsRStuts=conn.execute(strRStuts)
			if not rsRStuts.eof then
				ReturnStatus=trim(rsRStuts("StatusContent"))
			end if
			rsRStuts.close
			set rsRStuts=nothing
		else
			ReturnStatus="未處理"
		end if
	else
		if sys_City<>"台中市" then
			ReturnStatus="未上傳"
		else
			ReturnStatus="&nbsp;"
		end if 
	end if
	rsReturn.close
	set rsReturn=nothing

'-----------------------DciLog寄存--------------------------------
	StoreAndSendFileName=""	'寄存上傳檔名
	StoreAndSendBatchNumber=""	'寄存檔名
	StoreAndSendStatus=""	'寄存上傳狀態
	strSAndS="select * from DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='4'" &_
		" order by ExchangeDate desc"
	set rsSAndS=conn.execute(strSAndS)
	if not rsSAndS.eof then
		StoreAndSendFileName=trim(rsSAndS("FileName"))
		StoreAndSendBatchNumber=trim(rsSAndS("BatchNumber"))
		if trim(rsSAndS("DciReturnStatusID"))<>"" and not isnull(rsSAndS("DciReturnStatusID")) then
			strSStuts="select StatusContent from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsSAndS("DciReturnStatusID"))&"'"
			set rsSStuts=conn.execute(strSStuts)
			if not rsSStuts.eof then
				StoreAndSendStatus=trim(rsSStuts("StatusContent"))
			end if
			rsSStuts.close
			set rsSStuts=nothing
		else
			StoreAndSendStatus="未處理"
		end if
	else
		if sys_City<>"台中市" then
			StoreAndSendStatus="未上傳"
		else
			StoreAndSendStatus="&nbsp;"
		end if 
	end if
	rsSAndS.close
	set rsSAndS=nothing
'-----------------------DciLog公示--------------------------------
	OpenGovFileName=""	'公示上傳檔名
	OpenGovBatchNumber=""	'公示檔名
	OpenGovStatus=""	'公示上傳狀態
	strOpenGov="select * from DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='5'" &_
		" order by ExchangeDate desc"
	set rsOpenGov=conn.execute(strOpenGov)
	if not rsOpenGov.eof then
		OpenGovFileName=trim(rsOpenGov("FileName"))
		OpenGovBatchNumber=trim(rsOpenGov("BatchNumber"))
		if trim(rsOpenGov("DciReturnStatusID"))<>"" and not isnull(rsOpenGov("DciReturnStatusID")) then
			strOStuts="select StatusContent from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsOpenGov("DciReturnStatusID"))&"'"
			set rsOStuts=conn.execute(strOStuts)
			if not rsOStuts.eof then
				OpenGovStatus=trim(rsOStuts("StatusContent"))
			end if
			rsOStuts.close
			set rsOStuts=nothing
		else
			OpenGovStatus="未處理"
		end if
	else
		if sys_City<>"台中市" then
			OpenGovStatus="未上傳"
		else
			OpenGovStatus="&nbsp;"
		end if 
	end if
	rsOpenGov.close
	set rsOpenGov=nothing


%>

	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			end if
			%></span></td>
			<td width="27%">
			<span class="style2">車號：</span><span class="style1"><%=trim(rs1("CarNo"))%></span>
			</td>
			<%
'			<td width="27%"><span class="style2">到案處所：</span><span class="style1">
'			if trim(rs1("BillTypeID"))="2" then
'				StationName=StationNameDci
'			else
'				StationName=StationNameBillBase
'			end if
'			strStation="select * from Station where DciStationID='"&StationName&"'"
'			set rsStation=conn.execute(strStation)
'			if not rsStation.eof then
'				response.write trim(rsStation("DCIStationName"))
'			end if
'			rsStation.close
'			set rsStation=nothing
'			</span></td>
			%>
			<td width="23%"><span class="style2">告發類別：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))="2" then
				response.write "逕舉"
			else
				response.write "攔停"
			end if
			%></span></td>
			<td width="25%"><span class="style2">舉發單狀態：</span><span class="style1"><%
			if trim(rs1("RecordStateID"))="-1" then
				response.write "<font color=""red"">已刪除</font>"
			else
				response.write "正常"
			end if
			%></span></td>
		</tr>
		<tr>
			<%
'			<td><span class="style2">入案日期：</span><span class="style1">
'			if CaseInDate<>"" and not isnull(CaseInDate) then
'				response.write left(CaseInDate,len(CaseInDate)-4)&"-"&mid(CaseInDate,len(CaseInDate)-3,2)&"-"&mid(CaseInDate,len(CaseInDate)-1,2)
'			end if
'			</span></td>
			%>
			<td colspan="2"><span class="style2">違規時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">舉發單位：</span><span class="style1"><%
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
			<%
'			<td colspan="2"><span class="style2">舉發員警：</span><span class="style1">
'			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
'				response.write trim(rs1("BillMem1"))
'				strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
'				set rsMem1=conn.execute(strMem1)
'				if not rsMem1.eof then
'					response.write "("&trim(rsMem1("LoginID"))&")"
'				end if
'				rsMem1.close
'				set rsMem1=nothing
'			end if	
'			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
'				response.write "/&nbsp;"&trim(rs1("BillMem2"))
'				strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
'				set rsMem2=conn.execute(strMem2)
'				if not rsMem2.eof then
'					response.write "("&trim(rsMem2("LoginID"))&")"
'				end if
'				rsMem2.close
'				set rsMem2=nothing
'			end if	
'			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
'				response.write "/&nbsp;"&trim(rs1("BillMem3"))
'				strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
'				set rsMem3=conn.execute(strMem3)
'				if not rsMem3.eof then
'					response.write "("&trim(rsMem3("LoginID"))&")"
'				end if
'				rsMem3.close
'				set rsMem3=nothing
'			end if	
'			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
'				response.write "/&nbsp;"&trim(rs1("BillMem4"))
'				strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
'				set rsMem4=conn.execute(strMem4)
'				if not rsMem4.eof then
'					response.write "("&trim(rsMem4("LoginID"))&")"
'				end if
'				rsMem4.close
'				set rsMem4=nothing
'			end if	
'			</span></td>
			%>
		</tr>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				if left(trim(rs1("Rule1")),4)="2110" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" or trim(rs1("Rule1"))="4310104" then
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
					response.write trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing

				if trim(rs1("BillTypeID"))="2" and trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					response.write "&nbsp;"&trim(rs1("Rule4"))
				end if
			end if	
			%></span></td>
		</tr>
<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if left(trim(rs1("Rule2")),4)="2110" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" or trim(rs1("Rule2"))="4310104" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if left(trim(rs1("Rule3")),4)="2110" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" or trim(rs1("Rule3"))="4310104" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))<>"2" then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if left(trim(rs1("Rule4")),4)="2110" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" or trim(rs1("Rule4"))="4310104" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write trim(rs1("Rule4"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if	
			%></span></td>
		</tr>
<%end if%>
		<tr>
		<%
'			<td colspan="3"><span class="style2">違規路段：</span><span class="style1">
'			response.write trim(rs1("IllegalAddressID"))&" "&trim(rs1("IllegalAddress"))
'			</span></td>
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
			<td><span class="style2">詳細車種：</span><span class="style1"><%=DciCarType%></span></td>
		</tr>
		<!-- 
		<td><span class="style2">郵寄日期：</span><span class="style1"> --><%' response.write MailDate%><!-- </span></td> -->
			<!-- <td><span class="style2">郵寄序號：</span><span class="style1"> -->
			<%
'			if sys_City<>"台南縣" and sys_City<>"台南市" then
'				response.write MailNumber
'			else
'				response.write BillMailNo
'			end if
			%><!-- </span></td>
			-->
		<!-- 
		<td><span class="style2">違規人姓名：</span><span class="style1"> --><%' response.write IllegalMem%><!-- </span></td> -->
			
			<!-- <td colspan="3"><span class="style2">車主姓名：</span><span class="style1"> --><%' response.write OwnerName%><!-- </span></td>-->
			<!-- <td><span class="style2">填單日期：</span><span class="style1"> --><%
'			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
'				response.write gArrDT(trim(rs1("BillFillDate")))
'			end if	
			%><!-- </span></td> -->
			
			
			<!-- <td><span class="style2">到案日期：</span><span class="style1"> --><%
'			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
'				response.write gArrDT(trim(rs1("DealLineDate")))
'			end if	
			%><!-- </span></td> -->
			<!-- <td><span class="style2">建檔日期：</span><span class="style1"> --><%
'			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
'				response.write gArrDT(trim(rs1("RecordDate")))
'			end if	
			%><!-- </span></td> -->
			<!-- <td><span class="style2">操作人員：</span><span class="style1"> --><%
'			strRecMem="select ChName from MemberData where MemberID='"&trim(rs1("RecordMemberID"))&"'"
'			set rsRecMem=conn.execute(strRecMem)
'			if not rsRecMem.eof then
'				response.write trim(rsRecMem("ChName"))
'			end if
'			rsRecMem.close
'			set rsRecMem=nothing
			%><!-- </span></td> -->

			<!-- <td colspan="2"><span class="style2">代保管物：</span><span class="style1"> --><%
'			FastenerTmp=""
'			strFastener="select b.Content from BillFastenerDetail a,DciCode b where a.FastenerTypeID=b.ID and b.TypeID='6' and BillSN="&trim(rs1("SN"))
'			set rsFastener=conn.execute(strFastener)
'			If Not rsFastener.Bof Then rsFastener.MoveFirst 
'			While Not rsFastener.Eof
'				if FastenerTmp="" then
'					FastenerTmp=rsFastener("Content")
'				else
'					FastenerTmp=FastenerTmp&"、"&rsFastener("Content")
'				end if
'				rsFastener.MoveNext
'			Wend
'			rsFastener.close
'			set rsFastener=nothing
'			response.write FastenerTmp
			%><!-- </span></td> -->
			<!-- <td><span class="style2">移送日期：</span><span class="style1"> --><%' response.write MailDate%><!-- </span></td> -->

	</table>
	<hr>
<%	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>

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
