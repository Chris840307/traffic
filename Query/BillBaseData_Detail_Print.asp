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

	if trim(request("BillSnTmp"))<>"" then
'		strSQLTemp=strSQLTemp&" and a.SN='"&trim(request("BillSnTmp"))&"'"
'       2011/02/09 正修by jafe 原因 ，在快速查詢時會查出二筆，但是原本是用「 =」 會造成無效的數值，故改成 in
		strSQLTemp=strSQLTemp&" and a.SN in("&trim(request("BillSnTmp"))&")"
	end If
	
	If sys_City = "台中市" Or sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City="基隆市" then 
		strSQLAdd=strSQLAdd&",a.IsVideo"
	End if
	strSQL="Select a.SignType,a.BillNo,a.Sn,a.CarNo,a.BillTypeID,a.Rule1,a.Rule2,a.Rule3,a.Rule4" &_
		",a.MemberStation,a.EquipMentID,a.RuleSpeed,a.IllegalSpeed" &_
		",a.Recorddate,a.RecordMemberID,a.RecordStateID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2" &_
		",a.BillMem2,a.BillMemID3,a.BillMem3" &_
		",a.BillMemID4,a.BillMem4,a.RuleVer,a.IllegalAddressID,a.IllegalAddress,a.BillFillDate,a.BillUnitID" &_
		",a.driveraddress,a.driverzip,a.owner,a.ownerzip,a.owneraddress" &_
		",a.DealLineDate,a.Note,a.CarSimpleID,a.JurgeDay,a.TrafficAccidentNo,a.TrafficAccidentType"&strSQLAdd&" from BillBase a" &_
		" where ((a.RecordStateID<>-1 and a.BillStatus='0')" &_
		" or a.BillStatus<>'0') "&strSQLTemp

		'response.write strSQL
		'response.end


%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	Cnt=0
	BillSnTmp=""
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
	if BillSnTmp="" then
		BillSnTmp=trim(rs1("Sn"))
	else
		BillSnTmp=BillSnTmp&","&trim(rs1("Sn"))
	end if
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
	OwnerCID=""		'車主證號
	DciCarTypeID=""	'詳細車種代碼
	DciCarType=""	'詳細車種
'	strDciA="select * from BillBaseDciReturn where (BillNo='"&trim(rs1("BillNo"))&"' or BillNo is Null)" &_
'			" and CarNo='"&trim(rs1("CarNo"))&"'" &_
'			" and ExchangeTypeID='A' and Status='S'"
'	set rsDciA=conn.execute(strDciA)
'	if not rsDciA.eof and trim(rs1("BillTypeID"))="2" then
'
'		if sys_City<>"台中市" then
'			OwnerZipName=""
'			DriverZipName=""
'		else
'			if trim(rsDciA("NwnerZip"))<>"" and not isnull(rsDciA("NwnerZip")) then
'				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciA("NwnerZip"))&"'"
'				set rsOZip=conn.execute(strOZip)
'				if not rsOZip.eof then
'					OwnerZipName=trim(rsOZip("ZipName"))
'				end if
'				rsOZip.close
'				set rsOZip=nothing
'			else
'				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciA("OwnerZip"))&"'"
'				set rsOZip=conn.execute(strOZip)
'				if not rsOZip.eof then
'					OwnerZipName=trim(rsOZip("ZipName"))
'				end if
'				rsOZip.close
'				set rsOZip=nothing
'			end if
'
'			strDZip="select ZipName from Zip where ZipID='"&trim(rsDciA("DriverHomeZip"))&"'"
'			set rsDZip=conn.execute(strDZip)
'			if not rsDZip.eof then
'				DriverZipName=trim(rsDZip("ZipName"))
'			end if
'			rsDZip.close
'			set rsDZip=nothing
'		end if
'
'		StationNameDci=trim(rsDciA("DciReturnStation"))
'			
'		OwnerName=trim(rsDciA("Owner"))
'		OwnerAddress=trim(rsDciA("OwnerZip"))&" "&trim(rsDciA("OwnerAddress"))
'		DciCarTypeID=trim(rsDciA("DciReturnCarType"))
'		if trim(rs1("BillTypeID"))="1" then
'			IllegalMemID=trim(rsDciA("DriverID"))
'			IllegalMem=trim(rsDciA("Driver"))
'			IllegalAddress=trim(rsDciA("DriverHomeZip"))&" "&trim(rsDciA("DriverHomeAddress"))
'		else
'			if trim(rsDciA("Nwner"))<>"" and not isnull(rsDciA("Nwner")) then
'				IllegalMemID=trim(rsDciA("NwnerID"))
'				IllegalMem=trim(rsDciA("Nwner"))
'				IllegalAddress=trim(rsDciA("NwnerZip"))&" "&trim(rsDciA("NwnerAddress"))
'			else
'				IllegalMemID=trim(rsDciA("OwnerID"))
'				IllegalMem=trim(rsDciA("Owner"))
'				IllegalAddress=trim(rsDciA("OwnerZip"))&" "&trim(rsDciA("OwnerAddress"))
'			end if
'		end if
'	else
		strDciB="select a.* from BillBaseDciReturn a,DciReturnStatus b" &_
			" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
			" and (a.BillNo='"&trim(rs1("BillNo"))&"' or a.BillNo is Null)" &_
			" and a.CarNo='"&trim(rs1("CarNo"))&"'" &_
			" and b.DciReturnStatus=1 and ExchangeTypeID='W'"
		set rsDciB=conn.execute(strDciB)
		if not rsDciB.eof then

			'if sys_City<>"台中市" then
			'	OwnerZipName=""
			'	DriverZipName=""
			'else
				strOZip="select ZipName from Zip where ZipID='"&trim(rsDciB("OwnerZip"))&"'"
				set rsOZip=conn.execute(strOZip)
				if not rsOZip.eof Then
					If CDbl(Year(rs1("IllegalDate")))<2011 then
						OwnerZipName=ChangeOldCity(trim(rsDciB("OwnerZip")),trim(rsOZip("ZipName")))
					Else
						OwnerZipName=trim(rsOZip("ZipName"))
					End If 
				end if
				rsOZip.close
				set rsOZip=nothing

				strDZip="select ZipName from Zip where ZipID='"&trim(rsDciB("DriverHomeZip"))&"'"
				set rsDZip=conn.execute(strDZip)
				if not rsDZip.eof Then
					If CDbl(Year(rs1("IllegalDate")))<2011 then
						DriverZipName=ChangeOldCity(trim(rsDciB("DriverHomeZip")),trim(rsDZip("ZipName")))
					Else
						DriverZipName=trim(rsDZip("ZipName"))
					End If 
					
				end if
				rsDZip.close
				set rsDZip=nothing
			'end if
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
				if sys_City="台中市" then
					IllegalMemID=""
					IllegalMem=""
					IllegalAddress=""
				else
					IllegalMemID=trim(rsDciB("OwnerID"))
					IllegalMem=trim(rsDciB("Owner"))
					IllegalAddress=trim(rsDciB("OwnerZip"))&" "&OwnerZipName&trim(rsDciB("OwnerAddress"))
				end if
			end if
		else
			if (sys_City="高雄市" Or sys_City=ApconfigureCityName) and trim(rs1("BillTypeID"))="1" then
				strDciA1="select a.* from BillBaseDciReturn a,DciReturnStatus b" &_
				" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
				" and (a.BillNo='"&trim(rs1("BillNo"))&"' or a.BillNo is Null)" &_
				" and a.CarNo='"&trim(rs1("CarNo"))&"'" &_
				" and b.DciReturnStatus=1 and ExchangeTypeID='A'"
				set rsDciA1=conn.execute(strDciA1)
				if not rsDciA1.eof then
					strOZip1="select ZipName from Zip where ZipID='"&trim(rsDciA1("OwnerZip"))&"'"
					set rsOZip1=conn.execute(strOZip1)
					if not rsOZip1.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							OwnerZipName=ChangeOldCity(trim(rsDciA1("OwnerZip")),trim(rsOZip1("ZipName")))
						Else
							OwnerZipName=trim(rsOZip1("ZipName"))
						End If 						
					end if
					rsOZip1.close
					set rsOZip1=nothing

					strDZip1="select ZipName from Zip where ZipID='"&trim(rsDciA1("DriverHomeZip"))&"'"
					set rsDZip1=conn.execute(strDZip1)
					if not rsDZip1.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							DriverZipName=ChangeOldCity(trim(rsDciA1("DriverHomeZip")),trim(rsDZip1("ZipName")))
						Else
							DriverZipName=trim(rsDZip1("ZipName"))
						End If 							
					end if
					rsDZip1.close
					set rsDZip1=nothing

					OwnerName=trim(rsDciA1("Owner"))
					If Not IsNull(rsDciA1("OwnerAddress")) Then
						OwnerAddress=trim(rsDciA1("OwnerZip"))&" "&OwnerZipName&replace(replace(trim(rsDciA1("OwnerAddress")),"臺","台"),OwnerZipName,"")
					Else
						OwnerAddress=trim(rsDciA1("OwnerZip"))&" "&OwnerZipName&trim(rsDciA1("OwnerAddress"))
					End If 
					DciCarTypeID=trim(rsDciA1("DciReturnCarType"))
					IllegalMemID=trim(rsDciA1("DriverID"))
					IllegalMem=trim(rsDciA1("Driver"))
					If Not IsNull(rsDciA1("DriverHomeAddress")) Then
						IllegalAddress=trim(rsDciA1("DriverHomeZip"))&DriverZipName&" "&replace(replace(trim(rsDciA1("DriverHomeAddress")),"臺","台"),DriverZipName,"")
					Else
						IllegalAddress=trim(rsDciA1("DriverHomeZip"))&DriverZipName&" "&trim(rsDciA1("DriverHomeAddress"))
					End If 
					
				end if
				rsDciA1.close
				set rsDciA1=nothing
			end if
		end if
		rsDciB.close
		set rsDciB=nothing

		If sys_City="花蓮縣" Then
			If Not isnull(rs1("Owner")) Then
				IllegalMem=trim(rs1("Owner"))
				OwnerName=trim(rs1("Owner"))
			End If
			If Not isnull(rs1("OwnerAddress")) then
				IllegalAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
				OwnerAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
			End If 
		ElseIf sys_City="高雄市" Then '如果Billbase有寫以billbase為主
			If trim(rs1("BillTypeID"))="2" Then
				If Not isnull(rs1("Owner")) Then
					IllegalMem=trim(rs1("Owner"))
					OwnerName=trim(rs1("Owner"))
				End If
				If Not isnull(rs1("OwnerAddress")) Then

					IllegalAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
					OwnerAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
				End If
			End If 
		ElseIf sys_City="台中市" Then '如果Billbase有寫以billbase為主(逕舉不顯示違規人)
			If trim(rs1("BillTypeID"))="2" Then
				If Not isnull(rs1("Owner")) Then
					'IllegalMem=trim(rs1("Owner"))
					OwnerName=trim(rs1("Owner"))
				End If
				If Not isnull(rs1("OwnerAddress")) Then

					'IllegalAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
					OwnerAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
				End If
			End If 
		End If
'	end if
'	rsDciA.close
'	set rsDciA=nothing
	
	DciA_Name=""	'廠牌
	DciColor=""		'顏色
	DciDriverHomeAddress="" '車主戶籍地址
	DciIDstatus="" '行駕照狀態
	'if sys_City="台東縣" Or sys_City="高雄市" Or sys_City="高雄縣" then
		strDciA="select * from BillBaseDciReturn where (BillNo='"&trim(rs1("BillNo"))&"' or BillNo is Null)" &_
				" and CarNo='"&trim(rs1("CarNo"))&"'" &_
				" and ExchangeTypeID='A' and Status='S'"
		set rsDciA=conn.execute(strDciA)
		if not rsDciA.eof then
			OwnerCID=trim(rsDciA("OwnerID"))
			If trim(rsDciA("A_Name"))<>"" And Not IsNull(rsDciA("A_Name")) then
				DciA_Name=trim(rsDciA("A_Name"))
			End If
			if trim(rsDciA("DCIReturnCarColor"))<>"" and not isnull(rsDciA("DCIReturnCarColor")) then
				ColorLen=cint(Len(rsDciA("DCIReturnCarColor")))
				for Clen=1 to ColorLen
					colorID=mid(rsDciA("DCIReturnCarColor"),Clen,1)
					strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
					set rsColor=conn.execute(strColor)
					if not rsColor.eof then
						DciColor=DciColor & trim(rsColor("Content"))
					end if
					rsColor.close
					set rsColor=nothing
				next
			end If
			If trim(rsDciA("DriverHomeAddress"))<>"" And Not isnull(rsDciA("DriverHomeAddress")) then
				DciDriverHomeAddress=trim(rsDciA("DriverHomeZip"))&trim(rsDciA("DriverHomeAddress"))
			End If
			If trim(rsDciA("DciCounterID"))<>"" And Not isnull(rsDciA("DciCounterID")) then
				If trim(rsDciA("DciCounterID"))="x" Then
					DciIDstatus="駕照過期"
				ElseIf trim(rsDciA("DciCounterID"))="y" Then
					DciIDstatus="行照過期"
				ElseIf trim(rsDciA("DciCounterID"))="v" Then
					DciIDstatus="行駕照過期"
				End If 
			End If
		end if
		rsDciA.close
		set rsDciA=nothing
	'end if
	If sys_City="高雄市" Then '如果Billbase有寫以billbase為主
		If trim(rs1("BillTypeID"))="2" Then
			If Not isnull(rs1("driveraddress")) then
				DciDriverHomeAddress=trim(rs1("DriverZip"))&" "&trim(rs1("driveraddress"))
			End If
		End If 
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
	DciForfeit1=""	'罰金1
	DciForfeit2=""	'罰金2
	DciForfeit3=""	'罰金3
	strCaseIn="select a.*,c.* from BillBaseDciReturn a,DciReturnStatus b,DciLog c" &_
			" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
			" and a.ExchangeTypeID=c.ExchangeTypeID and a.Status=c.DciReturnStatusID" &_
			" and a.BillNo=c.BillNo and a.CarNo=c.CarNo" &_
			" and c.BillSn='"&trim(rs1("SN"))&"' " &_
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
		If Trim(rsCaseIn("Forfeit1"))<>"0" Then
			DciForfeit1=Trim(rsCaseIn("Forfeit1"))
		End If
		If Trim(rsCaseIn("Forfeit2"))<>"0" Then
			DciForfeit2=Trim(rsCaseIn("Forfeit2"))
		End If
		If Trim(rsCaseIn("Forfeit3"))<>"0" Then
			DciForfeit3=Trim(rsCaseIn("Forfeit3"))
		End if
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
	Storeandsendmailnumber=""  '台東寄存送達大宗號
	MailReturnDCIDate=""       '台東第一次單退日期
	StoreAndSendEffectDate=""	'寄存送達日
	StoreAndSendEndDate=""	'寄存送達生效(完成)日
	OpenGovGovNumber=""	'公示送達書號
	OpenGovEffectDate=""	'公示送達生效日
	OPenGovReasonID=""  '公式送送達原因
	StoreAndSendDate=""	'二次送達日期
	StoreAndSendReason=""	'二次送達原因
	BillMailNo=""	'郵寄序號
	ReturnMailNo=""	'退件郵寄序號
	MailCheckNumber="" '郵局查詢號
	MailReturnCheckNumber="" '單退後投遞郵局查詢號
	StoreAndSendFinalMailDate=""	'送達證書郵寄日期
	SignMan=""	'簽收人
	SignType=""
	
	'-------------------------------------------------------
	CancalSendDate=""   '撤銷送達日
	strCaseIn="select * from dcilog where " &_
						" BillSn=" & trim(rs1("SN")) & " and BillNo='"&trim(rs1("BillNo")) & "' and ExchangeTypeID='N' and ReturnMarkType='Y' and DCIRETURNSTATUSID='S'" 
	set rsCaseIn=conn.execute(strCaseIn)
	'response.write strCaseIn
	if not rsCaseIn.eof then
		CancalSendDate=gArrDT(trim(rsCaseIn("Exchangedate")))	
	end if
	rsCaseIn.close
	set rsCaseIn=nothing	
	'-------------------------------------------------------
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
		if trim(rs1("BillTypeID"))="2" or (trim(rs1("BillTypeID"))="1" and trim(rs1("EquipMentID"))="1") Then
			If sys_City="苗栗縣" Then
				If CaseInDate<>"" then
					MailDate=gArrDT(DateAdd("d",2,(left(CaseInDate,len(CaseInDate)-4)+1911)&"/"&mid(CaseInDate,len(CaseInDate)-3,2)&"/"&mid(CaseInDate,len(CaseInDate)-1,2)))
				End if
			else
				if trim(rsMail("MailDate"))<>"" and not isnull(rsMail("MailDate")) then
					MailDate=gArrDT(trim(rsMail("MailDate")))
				end If
			End if
		end If
		If sys_City="高雄市" Then
			MailNumber=trim(rsMail("MailChkNumber"))
		Else
			MailNumber=trim(rsMail("MailNumber"))
		End If 
		
		if CheckFlag=0 then
			if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
				ReturnMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
			end if
			if trim(rsMail("OpenGovMailReturnDate"))<>"" and not isnull(rsMail("OpenGovMailReturnDate")) then
				ReturnMailDate="&nbsp;"&gArrDT(trim(rsMail("OpenGovMailReturnDate")))
			end if
			GetMailDate=""
		else 
			'如果是收受. 這邊應該改成日期再 6/30前的要讀舊的欄位,支後讀這些欄未
			'台中市 6/30開始轉換 ,或是說如果signdate是空的就讀舊的. 						
			if trim(rsMail("SIGNDATE"))<>"" and not isnull(rsMail("SIGNDATE")) then
				GetMailDate=gArrDT(trim(rsMail("SIGNDATE")))
			else
				if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
					GetMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
				end if
			end if
			ReturnMailDate=""
		end if
		'退件or收受原因
		'smith
		'response.write "checkflag-->" & CheckFlag
		if CheckFlag=0 then
			'smith 20080626 暫時把收受注記誤寫入的部份排除掉
			if trim(rsMail("RETURNRESONID"))<>"" and not isnull(rsMail("RETURNRESONID")) and (rsMail("RETURNRESONID") <> "A") and (rsMail("RETURNRESONID") <> "B")and (rsMail("RETURNRESONID") <> "C")and (rsMail("RETURNRESONID") <> "D")  then
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
			'如果是收受. 這邊應該改成日期再 6/30前的要讀舊的欄位,支後讀這些欄未
			'台中市 6/30開始轉換 ,或是說如果SIGNRESONID是空的就讀舊的. 				
			if trim(rsMail("SIGNRESONID"))<>"" and not isnull(rsMail("SIGNRESONID")) then
				strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("SIGNRESONID"))&"'"
				set rsRR=conn.execute(strReturnReason)
				if not rsRR.eof then
					GetMailReason=trim(rsRR("Content"))
				end if
				rsRR.close
				set rsRR=nothing
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
		
 			if trim(rsMail("OPENGOVRESONID"))<>"" and not isnull(rsMail("OPENGOVRESONID")) and trim(ReturnReason) = "" then				
 				
					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("OPENGOVRESONID"))&"'"
					set rsRR=conn.execute(strReturnReason)
					if not rsRR.eof then	
						ReturnReason=trim(rsRR("Content"))			
					end if				
				rsRR.close			
				set rsRR=nothing			
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
		
		if trim(rsMail("Storeandsendmailnumber"))<>"" and not isnull(rsMail("Storeandsendmailnumber")) then
			Storeandsendmailnumber=trim(rsMail("Storeandsendmailnumber"))
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
			KSOpenGovEffDate=gArrDT(DateAdd("d",rsMail("OPENGOVDATE"),35))
		end if
		if trim(rsMail("OPENGOVRESONID"))<>"" and not isnull(rsMail("OPENGOVRESONID")) then
			strSReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMail("OPENGOVRESONID"))&"'"
			set rsSR=conn.execute(strSReason)
			if not rsSR.eof then
				OPENGOVRESONID=trim(rsSR("Content"))
			end if
			rsSR.close
			set rsSR=nothing
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
		'不管公示寄存都要寄第二次
		if sys_City="基隆市" then
			if trim(rsMail("UserMarkResonID"))<>"" and not isnull(rsMail("UserMarkResonID")) and trim(rsMail("UserMarkResonID"))<>"A" and trim(rsMail("UserMarkResonID"))<>"B" and trim(rsMail("UserMarkResonID"))<>"C" and trim(rsMail("UserMarkResonID"))<>"D" then
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
	ReturnIsClose=0 '單退是否結案
	strReturn="select * from DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='3'" &_
		" order by ExchangeDate desc"
	set rsReturn=conn.execute(strReturn)
	if not rsReturn.eof then
		ReturnFileName=trim(rsReturn("FileName"))
		ReturnBatchNumber=trim(rsReturn("BatchNumber"))
		MailReturnDCIDate=gArrDT(trim(rsReturn("ExchangeDate")))
		if trim(rsReturn("DciReturnStatusID"))="n" then
			ReturnIsClose=1
		end if
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

	'---------------------Smith 加入寄存期滿註記資料顯示	---------------------------------------------
	if sys_City="花蓮縣" then
		  MailStationReturnDate=""
			strQuery="select ReturnDate,StoreAndSendNumber from mailstationreturn where billno='"&trim(rs1("BillNo"))&"'"			
			set rsCity=conn.execute(strQuery)
			if not rsCity.eof then
				MailStationReturnDate=gArrDT(trim(rsCity("ReturnDate")))
				MailStationStoreAndSendNumber=Trim(rsCity("StoreAndSendNumber"))
			end if
			rsCity.close
			set rsCity=nothing		
	end if
	'----------------------------------------------------------------------------------------------------		


%>
<form name=myForm method="post">
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td align="center" colspan="2">
				<span class="style6">舉發違反交通管理事件通知單</span>
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
			<td align="right"><span class="style2">操作人：</span><span class="style1"><%
			strMem="select ChName from MemberData where MemberID='"&trim(session("User_ID"))&"'"
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			%></span>&nbsp; &nbsp; </td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">製表時間：</span><span class="style3"><%=now%></span></td>
		</tr>
	</table>
	<hr>
<%if trim(request("s_Detail"))="1" then%>
	
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			end if
			%></span></td>
			<td width="27%"><span class="style2">到案處所：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))<>"2" then
				StationName=StationNameBillBase
			end If
			If sys_City="台中市" Then 
				response.write getStationName_Date(StationName,trim(rs1("RecordDate")))
			Else
				strStation="select * from Station where DciStationID='"&StationName&"'"
				set rsStation=conn.execute(strStation)
				if not rsStation.eof then
					response.write trim(rsStation("DCIStationName"))
				end if
				rsStation.close
				set rsStation=nothing
			End If 
			
			%></span></td>
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
			
			'刪除原因
			if trim(rs1("RecordStateID"))="-1" Or sys_City="台中市" or trim(Session("Credit_ID"))="A000000000" then
				strDelRea="select b.Content from BillDeleteReason a,DciCode b where a.BillSn="&trim(rs1("Sn"))&" and b.TypeID=3 and a.DelReason=b.ID"
				set rsDelRea=conn.execute(strDelRea)
				if not rsDelRea.eof then
					response.write "<font color=""red"">." & trim(rsDelRea("Content")) & "</font>"
				else
					response.write "&nbsp;"
				end if
				rsDelRea.close
				set rsDelRea=nothing
			end if
			if trim(rs1("RecordStateID"))="-1" and (sys_City="高雄市" Or sys_City=ApconfigureCityName) then
				strDelTime="select * from log where typeid=352 and ActionContent like '%單號:"&trim(rs1("BillNo"))&"%' and ActionContent like '%車號:"&trim(rs1("CarNo"))&"%' and rownum<=1 order by ActionDate Desc"
				set rsDelTime=conn.execute(strDelTime)
				if not rsDelTime.eof then
					response.write "<font color=""red"">."&year(rsDelTime("ActionDate"))-1911&"/"&month(rsDelTime("ActionDate"))&"/"&day(rsDelTime("ActionDate"))&" "&hour(rsDelTime("ActionDate"))&":"&minute(rsDelTime("ActionDate"))&"</font>"
				end if
				rsDelTime.close
				set rsDelTime=nothing
			end if
			%>
			</span></td>
		</tr>
		<tr>
			<td><span class="style2">入案日期：</span><span class="style1"><%
			if CaseInDate<>"" and not isnull(CaseInDate) then
				response.write left(CaseInDate,len(CaseInDate)-4)&"-"&mid(CaseInDate,len(CaseInDate)-3,2)&"-"&mid(CaseInDate,len(CaseInDate)-1,2)
			end if
			%></span></td>
			<td><span class="style2">違規時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">舉發員警：</span><span class="style1"><%
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
				set rsMem1=conn.execute(strMem1)
				if not rsMem1.eof then
					response.write "("&trim(rsMem1("LoginID"))&")"
				end if
				rsMem1.close
				set rsMem1=nothing
			end if	
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "/&nbsp;"&trim(rs1("BillMem2"))
				strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
				set rsMem2=conn.execute(strMem2)
				if not rsMem2.eof then
					response.write "("&trim(rsMem2("LoginID"))&")"
				end if
				rsMem2.close
				set rsMem2=nothing
			end if	
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "/&nbsp;"&trim(rs1("BillMem3"))
				strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
				set rsMem3=conn.execute(strMem3)
				if not rsMem3.eof then
					response.write "("&trim(rsMem3("LoginID"))&")"
				end if
				rsMem3.close
				set rsMem3=nothing
			end if	
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "/&nbsp;"&trim(rs1("BillMem4"))
				strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
				set rsMem4=conn.execute(strMem4)
				if not rsMem4.eof then
					response.write "("&trim(rsMem4("LoginID"))&")"
				end if
				rsMem4.close
				set rsMem4=nothing
			end if	
			%></span></td>
		</tr>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				if (left(trim(rs1("Rule1")),2)="40" or (int(rs1("Rule1"))>4310200 and int(rs1("Rule1"))<4310209) or (int(rs1("Rule1"))>3310100 and int(rs1("Rule1"))<3310111)) and sys_City="基隆市" then
					response.write trim(rs1("Rule1"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
				else
					if left(trim(rs1("Rule1")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" or trim(rs1("Rule1"))="4310104" then
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
				end If
				If DciForfeit1<>"" And (sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City=ApconfigureCityName) Then
					response.write " &nbsp; 處新台幣 "&DciForfeit1&" 元"
				End if
			end if	
			%></span></td>
		</tr>
<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule2")),2)="40" or (int(rs1("Rule2"))>4310200 and int(rs1("Rule2"))<4310209) or (int(rs1("Rule2"))>3310100 and int(rs1("Rule2"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule2"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					if left(trim(rs1("Rule2")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" or trim(rs1("Rule2"))="4310104" then
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
			end If
			If DciForfeit2<>"" And (sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City=ApconfigureCityName) Then
				response.write " &nbsp; 處新台幣 "&DciForfeit2&" 元"
			End if
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule3")),2)="40" or (int(rs1("Rule3"))>4310200 and int(rs1("Rule3"))<4310209) or (int(rs1("Rule3"))>3310100 and int(rs1("Rule3"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule3"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					if left(trim(rs1("Rule3")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" or trim(rs1("Rule3"))="4310104" then
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
			end If
			If DciForfeit3<>"" And (sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City=ApconfigureCityName) Then
				response.write " &nbsp; 處新台幣 "&DciForfeit3&" 元"
			End if
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))<>"2" then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule4")),2)="40" or (int(rs1("Rule4"))>4310200 and int(rs1("Rule4"))<4310209) or (int(rs1("Rule4"))>3310100 and int(rs1("Rule4"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule4"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					if left(trim(rs1("Rule4")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" or trim(rs1("Rule4"))="4310104" then
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
			end If
			'If DciForfeit4<>"" And sys_City="高雄市" Then
			'	response.write " &nbsp; 處新台幣 "&DciForfeit4&" 元"
			'End if
			%></span></td>
		</tr>
<%end if%>
		<tr>
			<td colspan="<%
			If sys_City="台東縣" Then 
				response.write "2"
			Else
				response.write "3"
			End If 
			%>"><span class="style2">違規路段：</span><span class="style1"><%
			response.write trim(rs1("IllegalAddressID"))&" "&trim(rs1("IllegalAddress"))
			%></span></td>
<%	If sys_City="台東縣" Then
		If trim(rs1("Rule1"))="5620001" And not isnull(rs1("imagefilename")) Then 
%>
			<td><span class="style2">停車時間：</span><span class="style1"><%
			PFileArr=Split(trim(rs1("imagefilename")),"\")
			If UBound(PFileArr)>0 Then 
				PFile=Replace(PFileArr(1),".jpg","")
			End If 
			strPTime="select DealLineDate,IllegalDate from billbase " &_
				" where CarNo='"&Trim(rs1("CarNo"))&"' and billno is null " &_
				" and ImageFileNameB is not null and imagepathname='" & PFile & "'" &_
				" and recordstateid=0 "
			Set rsPTime=conn.execute(strPTime)
			If Not rsPTime.eof Then
				if trim(rsPTime("IllegalDate"))<>"" and not isnull(rsPTime("IllegalDate")) then
					response.write gArrDT(trim(rsPTime("IllegalDate")))&"&nbsp;"
					response.write Right("00"&hour(rsPTime("IllegalDate")),2)&":"
					response.write Right("00"&minute(rsPTime("IllegalDate")),2)
				end if	
			End If
			rsPTime.close
			Set rsPTime=Nothing 
			
			%></span></td>
<%		End If 
	End If 
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
		</tr>
		<tr>
			<td><span class="style2">郵寄日期：</span><span class="style1"><%=MailDate%></span></td>
			<td><span class="style2">郵寄序號：</span><span class="style1"><%
			if sys_City<>"台南縣" and sys_City<>"台南市" then
				response.write MailNumber
			else
				response.write BillMailNo
			end if
			%><%
			If MailNumber<>"" And Not IsNull(MailNumber) And sys_City="苗栗縣" Then
			%>
			<a href="mailhistory.asp?mailnum=<%=MailNumber%>" target="_blank" >郵寄歷程資料</a>
			<%
			End If 
			%></span></td>
			<td><span class="style2">簽收狀況：</span><span class="style1">
			<%
				'可參考google doc "攔停 簽收 狀況 "
				if trim(rs1("SignType"))<>"" and not isnull(rs1("SignType")) then
					if rs1("SignType")="A" then response.write "簽收"
					if rs1("SignType")="U" then 
						strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="2" then response.write "拒簽已收"
							if rsR2("SignStateID")="3" then response.write "已簽拒收"							
						else 
							response.write "拒簽收"
						end if
						rsR2.close
						set rsR2=nothing																
					end if				
				else
						strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="5" then response.write "補開單"
						end if
						rsR2.close
						set rsR2=nothing															
				end if
			%>			
			</span>
			
			</td>
			<!-- 20100107 jafe -->
			&nbsp;<span class="style2">序號:<span class="style1">
			<%
						strR2="select RecordDate,BillUnitID from Billbase where SN=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							RecordDate=Year(rsR2("RecordDate"))&"/"&Month(rsR2("RecordDate"))&"/"&Day(rsR2("RecordDate"))
							BillUnitID=rsR2("BillUnitID")
						end if
						rsR2.close
						set rsR2=nothing	

						strR2="select no from (select rownum  no,BillNo,sn from (select sn,BillNo from billbase where RecordStateID=0 and recordDate between to_date('"&RecordDate&" 0:0:0','yyyy/mm/dd hh24:mi:ss') and to_date('"&RecordDate&" 23:59:59','yyyy/mm/dd hh24:mi:ss') and  BillUnitID='"&BillUnitID&"' order by RecordDate)) where sn = "&trim(rs1("sn"))

						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							response.write rsR2("no")
						end if
						rsR2.close
						set rsR2=nothing


			%></span><span class="style2"></span>
			<!-- 20100107 jafe -->
			<%if sys_City="台東縣" then '台東縣要加車主證號%>
			<td><span class="style2">車主證號(查車)：</span><span class="style1"><%=OwnerCID%></span></td>
			<%end if%>
		</tr>
		<tr>
			<td><span class="style2">違規人證號：</span><span class="style1"><%=IllegalMemID%></span></td>
			<td><span class="style2">違規人姓名：</span><span class="style1"><%=funcCheckFont(IllegalMem,20,1)%></span></td>
			<td colspan="3"><span class="style2">違規人住址：</span><span class="style1"><%=funcCheckFont(IllegalAddress,20,1)%></span></td>
		</tr>
		<tr>
			<td><span class="style2">車號：</span><span class="style1"><font size="4"><%=trim(rs1("CarNo"))%></font></span></td>
			<td><span class="style2">車主姓名：</span><span class="style1"><%=funcCheckFont(OwnerName,20,1)%></span></td>
			<td colspan="3"><span class="style2">車主住址：</span><span class="style1"><%=funcCheckFont(OwnerAddress,20,1)%></span></td>
		</tr>
<%If sys_City="高雄市" Or sys_City="苗栗縣" then %>
		<tr>
			<td><span class="style2"></td>
			<td><span class="style2"></td>
			<td colspan="5"><span class="style2">(車籍)：</span><span class="style1"><%
			strNotify="select * from BillbaseDciReturn where CarNo='"&trim(rs1("CarNo"))&"' and Exchangetypeid='A'"
			Set rsNotify=conn.execute(strNotify)
			If Not rsNotify.eof Then
				
				response.write Trim(rsNotify("OwnerAddress"))
			End If
			rsNotify.close
			Set rsNotify=Nothing 
			%></span></td>
		</tr>
<%End If %>
<%If sys_City="花蓮縣" then %>
		<tr>
			<td colspan="5"><span class="style2">通訊住址：</span><span class="style1"><%
			strNotify="select * from BillbaseDciReturn where CarNo='"&trim(rs1("CarNo"))&"' and Exchangetypeid='A'"
			Set rsNotify=conn.execute(strNotify)
			If Not rsNotify.eof Then
				
				response.write Trim(rsNotify("OwnerNotifyAddress"))
			End If
			rsNotify.close
			Set rsNotify=Nothing 
			%></span></td>
		</tr>
<%End If %>
		<tr>
			<td><span class="style2">填單日期：</span><span class="style1"><%
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			end if	
			%></span></td>
			<td><span class="style2">詳細車種：</span><span class="style1"><%=DciCarType%></span></td>
			<td colspan="3"><span class="style2">舉發單位：</span><span class="style1"><%
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
			<td><span class="style2">到案日期：</span><span class="style1"><%
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			end if	
			%></span></td>
			<td><span class="style2">簡式車種：</span><span class="style1"><%
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車"
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "臨時車牌"
				end if
			end if	
			%></span></td>
			<td><span class="style2">建檔日期：</span><span class="style1"><%
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))
			end if	
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
			<%if (sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City=ApconfigureCityName) or sys_City="高雄縣" then%>	
				<tr>
					<td><span class="style2">廠牌：</span><span class="style1"><%=funcCheckFont(DciA_Name,20,1)%></span></td>
					<td><span class="style2">顏色：</span><span class="style1"><%=DciColor%></span></td>
					<td colspan="2"><span class="style2">車主戶籍地址：</span><span class="style1"><%=funcCheckFont(DciDriverHomeAddress,20,1)%></span></td>
				</tr>
			<%End if%>
			<%if (sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City=ApconfigureCityName) then%>	
				<tr>
				<td>
					<span class="style2"><%
					If sys_City="苗栗縣" Then
						response.write "違規事實："
					Else
						response.write "備註："
					End If 
					%></span><span class="style1"><%=trim(rs1("Note"))%></span>
				</td>
				</tr>
				<%if trim(rs1("BillTypeID"))="2" then

				strOther="select b.BillNO,c.DelReason from OtherBill a,BillBase b,BillDeleteReason c " &_
					" where a.NewBillSn="&trim(rs1("Sn")) &_
					" and a.OldBillSN=b.Sn and a.OldBillSN=c.Billsn" 
				set rsOther=conn.execute(strOther)
				if not rsOther.eof then
				%>
				<tr>
				<td colspan="3">
					<span class="style2"><font color="red">舉發前案：</font></span><span class="style1"><font color="red">
				<%
					response.write trim(rsOther("Billno"))

					strsql2="select * from DciCode where typeid=3 and id='"&trim(rsOther("DelReason"))&"'"
					set rs2=conn.execute(strsql2)
					if not rs2.eof then
						response.write "，"&rs2("Content")
					end if
					rs2.close
					set rs2=nothing
				%>
					</font></span>
				</td>
				</tr>
				<%
				end if
				rsOther.close
				set rsOther=nothing
				
				end if%>
			<% end if %>			
			<%if sys_City<>"台中市" then%>	
				<tr>
				<%if sys_City="台南市" then%>	
					<td ><span class="style2">輔助車種：</span><span class="style1"><%
					If Trim(rs1("CarAddID"))="1" Then
						response.write "1大貨"
					ElseIf Trim(rs1("CarAddID"))="2" Then
						response.write "2大客"
					ElseIf Trim(rs1("CarAddID"))="3" Then
						response.write "3砂石"
					ElseIf Trim(rs1("CarAddID"))="4" Then
						response.write "4土方"
					ElseIf Trim(rs1("CarAddID"))="5" Then
						response.write "5動力"
					ElseIf Trim(rs1("CarAddID"))="6" Then
						response.write "6貨櫃"
					ElseIf Trim(rs1("CarAddID"))="7" Then
						response.write "7大型重機"
					ElseIf Trim(rs1("CarAddID"))="8" Then
						response.write "8拖吊"
					ElseIf Trim(rs1("CarAddID"))="9" Then
						response.write "9(550cc)重機"
					ElseIf Trim(rs1("CarAddID"))="10" Then
						response.write "10計程車"
					ElseIf Trim(rs1("CarAddID"))="11" Then
						response.write "11危險物品"
					End If 
					%></span></td>
				<%End If %>
					<td colspan="2"><span class="style2">行駕照狀態：</span><span class="style1"><%=DciIDstatus%></span></td>
					<td ><span class="style2">民眾檢舉時間：</span><span class="style1"><%
				
					if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then
						response.write gArrDT(trim(rs1("JurgeDay")))
					end If
				
					%></span></td>
				</tr>
			<%End if%>
		<%
		strDSupd="select * from DCISTATUSUPDATE where Billsn="&Trim(rs1("Sn"))
		Set rsDSupd=conn.execute(strDSupd)
		If Not rsDSupd.eof Then
		%>
				<tr>
				<td colspan="3">
					<span class="style2">強制入案前狀態：</span><span class="style1"><%
				strDS1="select * from Dcireturnstatus where DciActionID='W' " &_
					" and DciReturn='"&Trim(rsDSupd("StatUS"))&"'"
				Set rsDS1=conn.execute(strDS1)
				If Not rsDS1.eof Then
					response.write rsDS1("StatusContent")
				End If
				rsDS1.close
				Set rsDS1=Nothing
				strDS2="select * from Dcireturnstatus where DciActionID='WE' " &_
					" and DciReturn='"&Trim(rsDSupd("DciErrorCarData"))&"'"
				Set rsDS2=conn.execute(strDS2)
				If Not rsDS2.eof Then
					response.write " "&rsDS2("StatusContent")
				End If
				rsDS2.close
				Set rsDS2=Nothing
				response.write " "&rsDSupd("RecordDate")
					%></span>
				</td>
				</tr>
		<%
		End If
		rsDSupd.close
		Set rsDSupd=nothing
		%>
	<%if sys_City="高雄縣" Or sys_City="苗栗縣" or (sys_City="高雄市" Or sys_City=ApconfigureCityName) then%>
		<tr>
			<td colspan="4">
			<span class="style2">代保管物：</span><span class="style1"><%
			strFast="select a.FASTENERTYPEID,b.Content from BILLFASTENERDETAIL a,DciCode b" &_
				" where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BillSn="&trim(rs1("SN"))
			set rsFast=conn.execute(strFast)
			If Not rsFast.Bof Then rsFast.MoveFirst 
			While Not rsFast.Eof
					response.write trim(rsFast("FASTENERTYPEID"))&trim(rsFast("Content"))&" "
				rsFast.MoveNext
			Wend
			rsFast.close
			set rsFast=nothing
			%></span>
			</td>
		</tr>
	<%end if%>
		<tr>
	<%if sys_City="台中市" then%>
			<td >告示單號：</span><span class="style1"><%
			strBR="select * from BillReportNo " &_
				" where BillSn="&trim(rs1("SN"))
			set rsBR=conn.execute(strBR)
			If Not rsBR.eof Then 
				response.write Trim(rsBR("ReportNo"))
			End If 
			rsBR.close
			set rsBR=nothing
			%></span>
			</td>
	<%End If %>
	<%If sys_City = "台中市" Or sys_City="高雄市" Or sys_City="苗栗縣" Or sys_City="基隆市" then %>
			<td >有無全程錄影：</span><span class="style1"><%
			If Trim(rs1("IsVideo"))="1" Then
				response.write "有"
			ElseIf Trim(rs1("IsVideo"))="0" Then
				response.write "無"
			End if
			%></span>
			</td>
	<%End If %>

		</tr>
	<%If sys_City="高雄市" then%>
		<tr>
			<td >交通事故案號：</span><span class="style1"><%
				response.write Trim(rs1("TrafficAccidentNo"))
			%></span>
			</td>
			<td >交通事故種類：</span><span class="style1"><%
				response.write Trim(rs1("TrafficAccidentType"))
			%></span>
			</td>
		</tr>
	<%End If %>

	</table>
	<hr>
<%end if%>
<%if trim(request("s_Mail"))="1" then%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
		
			<td width="35%"><span class="style2">入案檔名：</span><div id="Layer10" style="position:absolute ; width:90px; height:30px; z-index:0; border: 1px none #000000;">
			<span class="style1"><%=DciFileName%></span>
			</td>
			<td width="25%">
			<span class="style2">入案批號：</span><span class="style1"><%=DciBatchNumber%></span>
			</td>
			<td width="40%"><span class="style2">入案狀態：</span><span class="style1"><%
			if CaseStatus<>"" and not isnull(CaseStatus) then
				response.write CaseStatus
			end if
			%></span></td>
			
		</tr>
		<tr>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 上傳檔名："
			else
				response.write "簽收 上傳檔名："
			end if
			%></span><span class="style1"><%
			response.write GetFileName
			%></span>
			</td>
			<td>
			<span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存批號："
			else
				response.write "簽收批號："
			end if
			%></span><span class="style1"><%=GetBatchNumber%></span>
			</td>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 上傳狀態："
			else
				response.write "簽收 上傳狀態："
			end if
			%></span><span class="style1"><%
			response.write GetStatus
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 日期："
			else
				response.write "簽收 日期："
			end if
			%></span><span class="style1"><%
			response.write GetMailDate
			%></span></td>
			<td><span class="style2">簽收人：</span><span class="style1"><%
			response.write SignMan
			%></span></td>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 原因："
			else
				response.write "簽收 原因："
			end if
			%></span><span class="style1"><%
			response.write GetMailReason
			%></span>
			</td>			
		</tr>
		<tr>
		</tr>
		<tr>
			<td colspan="2"><span class="style2"><font color="clred">撤銷送達 日期：</font></span><span class="style1"><%
			response.write CancalSendDate
			%></span>
			</td>
			<td colspan="3"><span class="style2"><%
			if sys_City="花蓮縣" then
			%>查證文號：<%
			else
			%>寄存郵局：<%
			end if
			%></span><span class="style1"><%
			response.write MailStation
			%></span></td>
		</tr>		
<%
	if sys_City="台東縣" then
%>
		<tr>
			<td colspan="3"><span class="style2">附郵地址：</span><span class="style1"><%
			strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,dcierrorcardata,Nwner,NwnerZip,NwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
			set rsD=conn.execute(strSqlD)
			if not rsD.eof then
				if (trim(rsD("dcierrorcardata"))="D" and trim(rsD("OwnerZip"))<>"" and not isnull(rsD("OwnerZip"))) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							ZipName=ChangeOldCity(trim(rsD("OwnerZip")),trim(rsZip("ZipName")))
						Else
							ZipName=trim(rsZip("ZipName"))
						End If							
					end if
					rsZip.close
					set rsZip=nothing

					GetMailMem=trim(rsD("Owner"))
					GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
				elseif trim(rsD("DriverHomeZip"))<>"" and not isnull(rsD("DriverHomeZip")) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							ZipName=ChangeOldCity(trim(rsD("DriverHomeZip")),trim(rsZip("ZipName")))
						Else
							ZipName=trim(rsZip("ZipName"))
						End If						
					end if
					rsZip.close
					set rsZip=nothing
	
					GetMailMem=trim(rsD("Owner"))
					GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if isnull(rsD2("DriverHomeAddress")) then
							DriverHomeAddressLen=0
						else
							DriverHomeAddressLen=len(replace(trim(rsD2("DriverHomeAddress"))," ",""))
						end if
						if (trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))) and (DriverHomeAddressLen>3) then
							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("DriverHomeZip"))&trim(rsD2("DriverHomeAddress"))
						else
							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("OwnerZip"))&trim(rsD2("OwnerAddress"))
						end if
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof Then
							If CDbl(Year(rs1("IllegalDate")))<2011 then
								ZipName=ChangeOldCity(trim(rsD("OwnerZip")),trim(rsZip("ZipName")))
							Else
								ZipName=trim(rsZip("ZipName"))
							End If							
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					end if
					rsD2.close
					set rsD2=nothing
				end if
			end if
			rsD.close
			set rsD=nothing
			response.write funcCheckFont(GetMailAddress,20,1)
			%></span>
			</td>
		</tr>		
<%
	end if
%>
		<tr>
			<td><span class="style2">退件上傳檔名：</span>
			<span class="style1"><%=ReturnFileName%></span>
			</td>
			<td>
			<span class="style2">退件批號：</span><span class="style1"><%=ReturnBatchNumber%></span>
			</td>
			<td><span class="style2">退件上傳狀態：</span><span class="style1"><%=ReturnStatus%></span></td>
		</tr>
		<tr>
			<%if sys_City<>"台東縣" then %>
				<td colspan="2"><span class="style2">退件郵寄日期：</span><span class="style1"><%
				if sys_City="南投縣" then	'南投交通隊說單退結案不要顯示退件郵寄日981005
					if ReturnIsClose=0 then
						response.write ReturnSendMailDate
					end if
				else
					response.write ReturnSendMailDate
				end if
				%></span></td>
			<%else%>
				<td colspan="2"><span class="style2">退件郵寄日期：</span><span class="style1"><%=MailReturnDCIDate%></span></td>
			<%end if%>
			<td><span class="style2">退件郵寄序號：</span><span class="style1"><%
			if sys_City<>"台南縣" and sys_City<>"台南市" then
				response.write ReturnMailNumber
			else
				response.write ReturnMailNo
			end if
			%></span></td>
		</tr>
		<tr>

			<td colspan="2"><span class="style2">退回日期：</span><span class="style1"><%=ReturnMailDate%></span></td>
			<td><span class="style2">退件原因：</span><span class="style1"><%
			response.write ReturnReason
			if ReturnReason<>"" and sys_City="南投縣" then
				if instr(trim(rs1("Note")),"退回原因：")>0 then
					response.write "("&mid(trim(rs1("Note")),instr(trim(rs1("Note")),"退回原因：")+5,4)&")"
				end if
			end if
			%></span></td>			
		</tr>
		
		<tr>
			<td colspan="3"><span class="style2">二次郵寄地址：</span><span class="style1"><%
	if StoreAndSendFlag=1 or sys_City="彰化縣" then
		if trim(rs1("BillTypeID"))="1" then	
			'攔停抓戶籍地址
			strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
			set rsD=conn.execute(strSqlD)
			if not rsD.eof then
			if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣"  then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof Then
					If CDbl(Year(rs1("IllegalDate")))<2011 then
						ZipName=ChangeOldCity(trim(rsD("DriverHomeZip")),trim(rsZip("ZipName")))
					Else
						ZipName=trim(rsZip("ZipName"))
					End If						
				end if
				rsZip.close
				set rsZip=nothing
			end if
				GetMailMem=trim(rsD("Driver"))
				GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
			end if
			rsD.close
			set rsD=nothing
			response.write funcCheckFont(GetMailMem,20,1)&"--"&funcCheckFont(GetMailAddress,20,1)
		else
			if sys_City="台東縣" then
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeZip"))<>"" and not isnull(rsD("DriverHomeZip")) then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof Then
							If CDbl(Year(rs1("IllegalDate")))<2011 then
								ZipName=ChangeOldCity(trim(rsD("DriverHomeZip")),trim(rsZip("ZipName")))
							Else
								ZipName=trim(rsZip("ZipName"))
							End If									
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&trim(rsD2("DriverHomeAddress"))
							else
								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("OwnerZip"))&trim(rsD2("OwnerAddress"))
							end if
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD("OwnerZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If										
							end if
							rsZip.close
							set rsZip=nothing
			
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
						end if
						rsD2.close
						set rsD2=nothing
					end if
				end if
				rsD.close
				set rsD=nothing
			elseif sys_City="台中市" or (sys_City="高雄市" or sys_City=ApconfigureCityName) then
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD("DriverHomeZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If										
							end if
							rsZip.close
							set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
					else
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
					end if
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("DriverHomeZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If									
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("OwnerZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If								
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
					end if
					rsD2.close
					set rsD2=nothing
				end if
				rsD.close
				set rsD=nothing
			elseif sys_City="花蓮縣" or sys_City="基隆市" or sys_City="宜蘭縣" then
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
					'單退先抓W看有沒有做戶籍補正，沒有就抓owner
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
							if sys_City="宜蘭縣" then
								ZipName=""
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof Then
									If CDbl(Year(rs1("IllegalDate")))<2011 then
										ZipName=ChangeOldCity(trim(rsD2("DriverHomeZip")),trim(rsZip("ZipName")))
									Else
										ZipName=trim(rsZip("ZipName"))
									End If										
								end if
								rsZip.close
								set rsZip=nothing
							end if
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("OwnerZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If								
							end if
							rsZip.close
							set rsZip=nothing
			
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
				end if
				rsD2.close
				set rsD2=Nothing
			elseif sys_City="南投縣" then
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
					'單退先抓W看有沒有做戶籍補正，沒有的話再抓A,再沒有就抓owner
					if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
						'if sys_City="宜蘭縣" then
						'	ZipName=""
						'else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("DriverHomeZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If									
							end if
							rsZip.close
							set rsZip=nothing
						'end if
						
						if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
						end if
						GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress")),"臺","台"),ZipName,"")
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof Then
							If CDbl(Year(rs1("IllegalDate")))<2011 then
								ZipName=ChangeOldCity(trim(rsD2("OwnerZip")),trim(rsZip("ZipName")))
							Else
								ZipName=trim(rsZip("ZipName"))
							End If								
						end if
						rsZip.close
						set rsZip=nothing
			
						if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
						end if
						GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"") 
					end if
				end if
				rsD2.close
				set rsD2=nothing
			elseif sys_City<>"彰化縣" and sys_City<>"澎湖縣" and sys_City<>"台中市" and sys_City<>"台南市" and sys_City<>"台南縣" and sys_City<>"南投縣" and sys_City<>"高雄縣" then	'彰化澎湖單退要抓戶籍地址
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
				if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣" then
					ZipName=""
				else
					strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							ZipName=ChangeOldCity(trim(rsD("OwnerZip")),trim(rsZip("ZipName")))
						Else
							ZipName=trim(rsZip("ZipName"))
						End If							
					end if
					rsZip.close
					set rsZip=nothing
				end if
					GetMailMem=trim(rsD("Owner"))
					GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
				end if
				rsD.close
				set rsD=nothing
			else
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
					else
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
					end if
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("DriverHomeZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If									
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("OwnerZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If								
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
					end if
					rsD2.close
					set rsD2=nothing
				end if
				rsD.close
				set rsD=nothing
			end If
			If sys_City="高雄市" Then '如果Billbase有寫以billbase為主
				If trim(rs1("BillTypeID"))="2" Then
					If Not isnull(rs1("driveraddress")) then
						GetMailAddress=trim(rs1("DriverZip"))&" "&trim(rs1("driveraddress"))
					End If
				End If 
			End If
			response.write funcCheckFont(GetMailMem,20,1)&"--"&funcCheckFont(GetMailAddress,20,1)
		end if
	end if
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">寄存送達上傳檔名：</span>
			<span class="style1"><%=StoreAndSendFileName%></span>
			</td>
			<td>
			<span class="style2">寄存送達批號：</span>
			<span class="style1"><%=StoreAndSendBatchNumber%></span>
			</td>
			<td><span class="style2">寄存送達上傳狀態：</span><span class="style1"><%=StoreAndSendStatus%></span></td>
		</tr>
		<tr>
				<% if sys_City<>"台東縣" then %>
						<td colspan="2"><span class="style2">寄存送達書號：</span><span class="style1"><%=StoreAndSendGovNumber%></span></td>
					<% else %>
						<td colspan="2"><span class="style2">寄存送達掛號號碼：</span><span class="style1"><%=Storeandsendmailnumber%></span></td>	
					<% end if%>					
			<td><span class="style2">寄存送達日：</span><span class="style1"><%=StoreAndSendEffectDate%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">
				<% if sys_City="台東縣" then %>
					寄存送達生效
				<%elseif sys_City="高雄縣" then %>
					寄存送達期滿
				<%else%>
					寄存送達生效(完成)
				<% end if%>					
				日：</span><span class="style1"><%=StoreAndSendEndDate%></span></td>
			<td><span class="style2">寄存送達 
				<% if sys_City<>"台東縣" then %>
					退件原因：
				<% else %>
					狀態：
				<% end if%>
				</span><span class="style1"><%=StoreAndSendReason%></span></td>
		</tr>
		<tr>
			<%	
				'smith 加入寄存期滿退回日顯示 
				if sys_City<>"花蓮縣" then
			%>	
				<% if sys_City<>"台東縣" then %>
					<td colspan="3"><span class="style2">寄存送達 退回日期：</span><span class="style1"><%=StoreAndSendDate%></span></td>			
				<% end if %>					
			<% else %>
					<td colspan="2"><span class="style2">寄存送達 退回日期：</span><span class="style1"><%=StoreAndSendDate%></span></td>
					<td><span class="style2">寄存送達期滿 退回日期：</span><span class="style1"><%=MailStationReturnDate%></span></td>
				
			<% end if%>
			
		</tr>
		<tr>
			<td><span class="style2">公示送達上傳檔名：</span>
			<span class="style1"><%=OpenGovFileName%></span>
			</td>
			<td>
			<span class="style2">公示送達批號：</span><span class="style1"><%=OpenGovBatchNumber%></span>
			</td>
			<td><span class="style2">公示送達上傳狀態：</span><span class="style1"><%=OpenGovStatus%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">公示送達書號：</span><span class="style1"><%=OpenGovGovNumber%></span></td>
			<td><span class="style2"><%
			if sys_City="高雄縣" then
				response.write "公示送達公告日："
			else
				response.write "公示送達生效日："
			end if
			%></span><span class="style1"><%=OpenGovEffectDate%></span></td>
		</tr>

		<tr>
			
			<td colspan="2"><span class="style2">發文監理站日期：</span><span class="style1"><%=ReturnSendDate%></span></td>
			<td><span class="style2">公示送達原因：</span>
				<span class="style1">
					<% 							
						if  trim(OpenGovBatchNumber) <> "" then 
							response.write OPENGOVRESONID
						 end if
					%>
				</span>
			</td>
		</tr>
		<tr>
			<td colspan="<%
			if sys_City="高雄縣" then
				response.write "2"
			else
				response.write "3"
			end if
			%>"><span class="style2">備註：</span><span class="style1"><%=trim(rs1("Note"))%></span></td>
			<%if sys_City="高雄縣" then%>
			<td><span class="style2">公示送達生效日：</span>
				<span class="style1">
					<% 							
						if  trim(KSOpenGovEffDate) <> "" then 
							response.write KSOpenGovEffDate
						 end if
					%>
				</span>
			</td>
			<%end if%>
		</tr>
	<%if sys_City="基隆市" or sys_City="台中市" or sys_City="屏東縣" then%>
		<tr>
			<td colspan="3"><span class="style2">第一次投遞郵局查詢號：</span><span class="style1"><%=MailCheckNumber%> <%if sys_City="屏東縣" then response.write " <font size='2'> (98年11月後案件)</font>" end if%></span></td>
		</tr>
		<tr>
			<td colspan="3"><span class="style2">單退後投遞郵局查詢號：</span><span class="style1"><%=MailReturnCheckNumber%></span></td>
		</tr>
	<%end if%>
	<%if sys_City="基隆市" then%>
		<tr>
			<td colspan="3"><span class="style2">送達證書郵寄日期：</span><span class="style1"><%=StoreAndSendFinalMailDate%></span></td>
		</tr>
	<%end if%>
	</table>
<%end if%>
<%
if trim(request("s_Image"))="1" then
%>
	<div align="center">
<%
		'違規照片
		strImgKS="select * from BILLILLEGALIMAGE where billsn="&trim(rs1("SN"))
		set rsImgKS=conn.execute(strImgKS)
		if not rsImgKS.eof then
			response.write "<strong>違規影像</strong>　　"
			response.write "<br>"
			if trim(rsImgKS("ImageFileNameA"))<>"" then
		%>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameA"))%>" name="imgB1" width="600">
		<%
			end if
			if trim(rsImgKS("ImageFileNameB"))<>"" then
		%>		<br>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameB"))%>" name="imgB2" width="600">
		<%
			end if
			if trim(rsImgKS("ImageFileNameC"))<>"" then
		%>		<br>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameC"))%>" name="imgB3" width="600">
		<%
			end if
		end if
		rsImgKS.close
		set rsImgKS=Nothing
	If sys_City="高雄市" Then 
		'違規照片
		strImgKS="select * from BILLILLEGALIMAGETemp2 where billsn="&trim(rs1("SN"))
		set rsImgKS=conn.execute(strImgKS)
		if not rsImgKS.eof then
			response.write "<strong>民眾檢舉違規影像</strong>　　"
			response.write "<br>"
			
PicturePath="/ReportCaseImage"
			Vedio1=""
			Picture1=""

			If trim(rsImgKS("ImageFileNameA"))<>"" Then
				ImageFileNameAArray=Split(Trim(rsImgKS("ImageFileNameA")),"/")
				ImageFileNameATemp=ImageFileNameAArray(UBound(ImageFileNameAArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameA")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameA")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameA")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameA")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameA")),3))="GIF" Then
					IsPicture1="1"
				Else
					IsPicture1="0"
				End If 
			
				'bPicWebPath= Trim(rsImgKS("ImageFileName"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameA")),4))="HTTP" then
					bPicWebPath=Trim(rsImgKS("ImageFileNameA"))
					
				Else
					bPicWebPath="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameA"))
				End If 
				If IsPicture1="1" Then
					Picture1="<img src="""&bPicWebPath&""" border=1 width=""650"" id=""img1"" onclick=""OpenPic('"&bPicWebPath&"')"">"
				Else
					Vedio1="<a href="""&bPicWebPath&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameATemp&"</a>"
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameB"))<>"" Then
				ImageFileNameBArray=Split(Trim(rsImgKS("ImageFileNameB")),"/")
				ImageFileNameBTemp=ImageFileNameBArray(UBound(ImageFileNameBArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameB")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameB")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameB")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameB")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameB")),3))="GIF" Then
					IsPicture2="1"
				Else
					IsPicture2="0"
				End If 

				'bPicWebPath2= Trim(rsImgKS("ImageFileNameB"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameB")),4))="HTTP" then
					bPicWebPath2=Trim(rsImgKS("ImageFileNameB"))
				Else
					bPicWebPath2="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameB"))
				End If 
				If IsPicture2="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath2&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath2&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameBTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath2&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameBTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameC"))<>"" Then
				ImageFileNameCArray=Split(Trim(rsImgKS("ImageFileNameC")),"/")
				ImageFileNameCTemp=ImageFileNameCArray(UBound(ImageFileNameCArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameC")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameC")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameC")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameC")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameC")),3))="GIF" Then
					IsPicture3="1"
				Else
					IsPicture3="0"
				End If 

				'bPicWebPath3= Trim(rsImgKS("ImageFileNameC"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameC")),4))="HTTP" then
					bPicWebPath3=Trim(rsImgKS("ImageFileNameC"))
				Else
					bPicWebPath3="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameC"))
				End If 
				If IsPicture3="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath3&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath3&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameCTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath3&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameCTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameD"))<>"" Then
				ImageFileNameDArray=Split(Trim(rsImgKS("ImageFileNameD")),"/")
				ImageFileNameDTemp=ImageFileNameDArray(UBound(ImageFileNameDArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameD")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameD")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameD")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameD")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameD")),3))="GIF" Then
					IsPicture4="1"
				Else
					IsPicture4="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameD"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameD")),4))="HTTP" then
					bPicWebPath4=Trim(rsImgKS("ImageFileNameD"))
				Else
					bPicWebPath4="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameD"))
				End If 
				If IsPicture4="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath4&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath4&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameDTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath4&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameDTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameE"))<>"" Then
				ImageFileNameEArray=Split(Trim(rsImgKS("ImageFileNameE")),"/")
				ImageFileNameETemp=ImageFileNameEArray(UBound(ImageFileNameEArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameE")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameE")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameE")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameE")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameE")),3))="GIF" Then
					IsPicture5="1"
				Else
					IsPicture5="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameE"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameE")),4))="HTTP" then
					bPicWebPath5=Trim(rsImgKS("ImageFileNameE"))
				Else
					bPicWebPath5="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameE"))
				End If 
				If IsPicture5="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath5&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath5&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameETemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath5&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameETemp&"</a>"
					End If 
				End If 
			End If 
			
			If trim(rsImgKS("ImageFileNameF"))<>"" Then
				ImageFileNameFArray=Split(Trim(rsImgKS("ImageFileNameF")),"/")
				ImageFileNameFTemp=ImageFileNameFArray(UBound(ImageFileNameFArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameF")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameF")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameF")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameF")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameF")),3))="GIF" Then
					IsPicture6="1"
				Else
					IsPicture6="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameF"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameF")),4))="HTTP" then
					bPicWebPath6=Trim(rsImgKS("ImageFileNameF"))
				Else
					bPicWebPath6="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameF"))
				End If 
				If IsPicture6="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath6&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath6&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameFTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath6&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameFTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameG"))<>"" Then
				ImageFileNameGArray=Split(Trim(rsImgKS("ImageFileNameG")),"/")
				ImageFileNameGTemp=ImageFileNameGArray(UBound(ImageFileNameGArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameG")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameG")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameG")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameG")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameG")),3))="GIF" Then
					IsPicture7="1"
				Else
					IsPicture7="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameF"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameG")),4))="HTTP" then
					bPicWebPath7=Trim(rsImgKS("ImageFileNameG"))
				Else
					bPicWebPath7="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameG"))
				End If 
				If IsPicture7="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath7&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath7&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameGTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath7&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameGTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameH"))<>"" Then
				ImageFileNameHArray=Split(Trim(rsImgKS("ImageFileNameH")),"/")
				ImageFileNameHTemp=ImageFileNameHArray(UBound(ImageFileNameHArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameH")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameH")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameH")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameH")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameH")),3))="GIF" Then
					IsPicture8="1"
				Else
					IsPicture8="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameF"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameH")),4))="HTTP" then
					bPicWebPath8=Trim(rsImgKS("ImageFileNameH"))
				Else
					bPicWebPath8="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameH"))
				End If 
				If IsPicture8="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath8&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath8&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameHTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath8&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameHTemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameI"))<>"" Then
				ImageFileNameIArray=Split(Trim(rsImgKS("ImageFileNameI")),"/")
				ImageFileNameITemp=ImageFileNameIArray(UBound(ImageFileNameIArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameI")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameI")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameI")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameI")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameI")),3))="GIF" Then
					IsPicture9="1"
				Else
					IsPicture9="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameF"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameI")),4))="HTTP" then
					bPicWebPath9=Trim(rsImgKS("ImageFileNameI"))
				Else
					bPicWebPath9="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameI"))
				End If 
				If IsPicture9="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath9&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath9&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameITemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath9&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameITemp&"</a>"
					End If 
				End If 
			End If 

			If trim(rsImgKS("ImageFileNameJ"))<>"" Then
				ImageFileNameJArray=Split(Trim(rsImgKS("ImageFileNameJ")),"/")
				ImageFileNameJTemp=ImageFileNameJArray(UBound(ImageFileNameJArray))

				If UCase(Right(Trim(rsImgKS("ImageFileNameJ")),3))="BMP" Or UCase(Right(Trim(rsImgKS("ImageFileNameJ")),3))="PNG" Or UCase(Right(Trim(rsImgKS("ImageFileNameJ")),3))="JPG" Or UCase(Right(Trim(rsImgKS("ImageFileNameJ")),4))="JPEG" Or UCase(Right(Trim(rsImgKS("ImageFileNameJ")),3))="GIF" Then
					IsPicture10="1"
				Else
					IsPicture10="0"
				End If 

				'bPicWebPath4= Trim(rsImgKS("ImageFileNameF"))
				if UCase(Left(Trim(rsImgKS("ImageFileNameJ")),4))="HTTP" then
					bPicWebPath10=Trim(rsImgKS("ImageFileNameJ"))
				Else
					bPicWebPath10="/Imgfix/ReportCase/" & Trim(rsImgKS("ImageFileNameJ"))
				End If 
				If IsPicture10="1" Then
					Picture1=Picture1&"<br/><img src="""&bPicWebPath10&""" border=1 width=""650"" id=""img1"">"
				Else
					If Vedio1="" Then
						Vedio1="<a href="""&bPicWebPath10&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameJTemp&"</a>"
					Else
						Vedio1=Vedio1&" 、 <a href="""&bPicWebPath10&""" target=""_blank"" style=""font-size: 18px;"">"&ImageFileNameJTemp&"</a>"
					End If 
				End If 
			End If 


'			If Vedio1<>"" then
'			response.write "動態錄影檔  " & Vedio1 & "<br/>"
'			End if
			response.write "<br>"
			response.write Picture1



		end if
		rsImgKS.close
		set rsImgKS=Nothing
	End If 
%>
	</div>
<%
end if
if trim(request("s_Send"))="1" then
%>
	<div align="center">
<%
		'送達証書
		strScan="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID in (0,1,4) and Recordstateid=0 order by RecordDate"
		set rsScan=conn.execute(strScan)
		while Not rsScan.eof
		%>
			<br> <strong>送達証書&nbsp;<%
			'掃描日期
			response.write year(rsScan("RecordDate"))&"/"&month(rsScan("RecordDate"))&"/"&day(rsScan("RecordDate"))&" "&hour(rsScan("RecordDate"))&":"&minute(rsScan("RecordDate"))
			'掃瞄人
			strSMem="select Chname from Memberdata where memberid="&trim(rsScan("RecordMemberiD"))
			set rsSMem=conn.execute(strSMem)
			if not rsSMem.eof then
				response.write "&nbsp;"&rsSMem("Chname")
			end if
			rsSMem.close
			set rsSMem=nothing
			%></strong><br><img src='<%=replace(trim(rsScan("FileName")),"/img/","/scanimg/")%>' name='imgB1' width='700'>
		<%
			rsScan.movenext
		wend
		rsScan.close
		set rsScan=nothing
%>
	</div>
<%
end if
if trim(request("s_Gov"))="1" then
%>
	<div align="center">
<%
		'公告公文
		strScan2="select * from BillAttatchImage where BillNo='"&trim(OpenGovBatchNumber)&"' and TypeID=4 and Recordstateid=0"
		set rsScan2=conn.execute(strScan2)
		while Not rsScan2.eof
		%>
			<div class='PageNext'>&nbsp;</div> <strong>公告公文</strong><br><img src='<%=replace(trim(rsScan2("FileName")),"/img/","/scanimg/")%>' name='imgB1' width='700' >
		<%
		rsScan2.movenext
		wend
		rsScan2.close
		set rsScan2=nothing
%>
	</div>
<%
end if
%>

<%	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>

<%
conn.close
set conn=nothing
%>
</form>
</body>
<script type="text/javascript" src="../js/tjpzoom.js"></script>
<script language="JavaScript">
	window.focus();
	window.print();
</script>
</html>
