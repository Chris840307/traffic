<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>交寄大宗函件</title>
<script type="text/javascript" src="../js/Print.js"></script>
<%
Server.ScriptTimeout = 6800
Response.flush
'權限
'AuthorityCheck(234)
%>
<style type="text/css">
<!--

.style35 {
	font-size: 10pt;
	font-family: "標楷體";
}
.style33 {
	font-size: 8pt;
	font-family: "新細明體";
}
.style5 {
	font-size: 8pt;
	font-family: "標楷體";}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 12pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
</head>

<body>

<%
strwhere=request("SQLstr")

'郵資
theMailMoney=trim(request("MailMoneyValue"))
'使用者單位資料
UnitName=""
UnitAddress=""
UnitTel=""
strUnitName="select Value from ApConfigure where ID=40"
set rsUnitName=conn.execute(strUnitName)
if not rsUnitName.eof then
	TitleUnitName=trim(rsUnitName("value"))
end if
rsUnitName.close
set rsUnitName=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

ExchangeTypeFlag="W"
stopBatchnumber=""
DealLineDateTmp=""
RecordMemberIDTemp=""
strExchangeType="select a.ExchangeTypeID,f.BillUnitID,a.Batchnumber,f.DealLineDate,f.RecordMemberID from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 "&strwhere
set rsEType=conn.execute(strExchangeType)
if not rsEType.eof then
	if trim(rsEType("ExchangeTypeID"))="N" then
		ExchangeTypeFlag="N"
	else
		ExchangeTypeFlag="W"
	end if
	BillUnitIDtmp=trim(rsEType("BillUnitID"))
	stopBatchnumber=left(trim(rsEType("Batchnumber")),2)
	DealLineDateTmp=Year(rsEType("DealLineDate"))-1911&"/"&month(rsEType("DealLineDate"))&"/"&day(rsEType("DealLineDate"))
	RecordMemberIDTemp=trim(rsEType("RecordMemberID"))
else
	ExchangeTypeFlag="W"
	BillUnitIDtmp=""
end if
rsEType.close
set rsEType=nothing
'台中市停管
if sys_City="台中市" and stopBatchnumber="WT" then
	strwhere=strwhere&" and (f.Note like '2%')"
end if
if sys_City="台中市" then 
	if BillUnitIDtmp="" then
		strSendMailUnit="select b.UnitName,b.Address,b.Tel from Apconfigure a,UnitInfo b " &_
				" where a.ID=49 and a.Value=b.UnitID"
		set rsSendMailUnit=conn.execute(strSendMailUnit)
		if not rsSendMailUnit.eof then
			
			if sys_City<>"花蓮縣" and sys_City<>"台中市" then 
				UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
			else
				UnitName=trim(rsSendMailUnit("UnitName"))
			end if
			UnitAddress=trim(rsSendMailUnit("Address"))
			UnitTel=trim(rsSendMailUnit("Tel"))
		end if
		rsSendMailUnit.close
		set rsSendMailUnit=nothing
	else
		'檢查舉發單位showorder
		strShow="select * from UnitInfo where UnitID='"&BillUnitIDtmp&"'"
		set rsShow=conn.execute(strShow)
		if not rsShow.eof then
			'showorder=0 or 1,寄件人就是舉發單位
			if trim(rsShow("ShowOrder"))="0" or trim(rsShow("ShowOrder"))="1" or trim(rsShow("UnitID"))="046A" or trim(rsShow("UnitID"))="0469" then
				UnitName=trim(rsShow("UnitName"))
				UnitAddress=trim(rsShow("Address"))
				UnitTel=trim(rsShow("Tel"))
			'showorder=2,寄件人是上層單位
			elseif trim(rsShow("ShowOrder"))="2" then
				strUnitType="select * from UnitInfo where UnitID='"&trim(rsShow("UnitTypeID"))&"'"
				set rsUnitType=conn.execute(strUnitType)
				if not rsUnitType.eof then
					UnitName=trim(rsUnitType("UnitName"))
					UnitAddress=trim(rsUnitType("Address"))
					UnitTel=trim(rsUnitType("Tel"))
				end if
				rsUnitType.close
				set rsUnitType=nothing
			end if
		else
			UnitName=""
			UnitAddress=""
			UnitTel=""
		end if
		rsShow.close
		set rsShow=nothing
	end If
elseif sys_City="屏東縣" And BillUnitIDtmp="9800" then 
	strSendMailUnit="select UnitName,Address,Tel from UnitInfo " &_
			" where UnitID='" & BillUnitIDtmp & "'"
	set rsSendMailUnit=conn.execute(strSendMailUnit)
	if not rsSendMailUnit.eof then
		
		UnitName=replace(rsSendMailUnit("UnitName"),"屏東縣政府警察局","")

		UnitAddress=trim(rsSendMailUnit("Address"))
		UnitTel=trim(rsSendMailUnit("Tel"))
	end if
	rsSendMailUnit.close
	set rsSendMailUnit=nothing
else
	strSendMailUnit="select b.UnitName,b.Address,b.Tel from MemberData a,UnitInfo b " &_
			" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
	set rsSendMailUnit=conn.execute(strSendMailUnit)
	if not rsSendMailUnit.eof then
		
		if sys_City="花蓮縣" then 
			UnitName=trim(rsSendMailUnit("UnitName"))
		elseif sys_City="屏東縣" then 
			UnitName=TitleUnitName&replace(rsSendMailUnit("UnitName"),"屏東縣政府警察局","")
		else
			UnitName=TitleUnitName&replace(trim(rsSendMailUnit("UnitName")),TitleUnitName,"")
		end if
		UnitAddress=trim(rsSendMailUnit("Address"))
		UnitTel=trim(rsSendMailUnit("Tel"))
	end if
	rsSendMailUnit.close
	set rsSendMailUnit=nothing
end if

If sys_City="苗栗縣" Then
	ChangeRow=0
	BatchNumberTmp=""
	strB="select distinct(a.BatchNumber) " &_
	" from DCILog a" &_
	",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
	'response.write strB
	set rsB=conn.execute(strB)
	While Not rsB.Eof
		strBDel="Delete from batchnumberjob where batchNumber='"&Trim(rsB("Batchnumber"))&"' and PrintTypeID=0"
		conn.execute strBDel

		strBIns="Insert into batchnumberjob values('"&Trim(rsB("Batchnumber"))&"',"&Trim(session("User_ID"))&",2,sysdate)"
		conn.execute strBIns

		ChangeRow=ChangeRow+1
		If BatchNumberTmp="" Then
			BatchNumberTmp=Trim(rsB("Batchnumber"))
		Else	
			BatchNumberTmp=BatchNumberTmp&","&Trim(rsB("Batchnumber"))
			If ChangeRow=12 Then
				BatchNumberTmp=BatchNumberTmp&"<br>"
			End If 
		End If 
	rsB.MoveNext
	Wend
	rsB.close
	Set rsB=Nothing 
End If 
	'每頁幾筆
	PageCaseCnt=60

	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N')" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) and a.BillSn=g.BillSN "&strwhere&" order by g.MailNumber,f.RecordDate"
	end if


set rs1=conn.execute(strSQL)

	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
	",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere

set rsCnt=conn.execute(strCnt)
if not rsCnt.eof then
	if trim(rsCnt("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rsCnt("cnt"))/PageCaseCnt+0.9999999)
	end if
end if
rsCnt.close
set rsCnt=nothing

MDate=""
if ExchangeTypeFlag="N" then
	strMailDate="select g.STOREANDSENDSENDDATE as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
else
	strMailDate="select g.MailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
end if
	'response.write strMailDate
	set rsMailDate=conn.execute(strMailDate)
	if not rsMailDate.eof then
		MDate=trim(rsMailDate("MDate"))
	end if
	rsMailDate.close
	set rsMailDate=nothing
	if MDate="" or isnull(MDate) then
		MDate=now
	end if

CaseSN=0
mailSNTmp=0

If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	BillFillDateTmp=""
	if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
		BillFillDateTmp=trim(rs1("BillFillDate"))
	end if
	strList=""
	mailSN=0
	pageNum=fix(CaseSN/PageCaseCnt)+1
	for i=1 to PageCaseCnt
		if rs1.eof then exit for
		ZipName=""
		sysBillTypeID=trim(rs1("BillTypeID"))
		MailBatchNumber=trim(rs1("BatchNumber"))
		mailSN=mailSN+1
		CaseSN=CaseSN+1
		strList=strList&"<tr>"		
		'順序號碼
		strList=strList&"<td align=""center"" class=""style33"">"&CaseSN&"</td>"
			
		'掛號號碼
		theMailNumber=""
		'移送監理站日期
		theSendDocDate=""
		strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from BillMailHistory where BillSN="&trim(rs1("BillSN"))
		set rsH=conn.execute(strSqlH)
		if not rsH.eof then
			if sys_City="台中市" or sys_City="雲林縣" then
				if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
					theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
				end if
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=right("00000000" & trim(rsH("MailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=right("00000000" & trim(rsH("StoreAndSendMailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end if
'			elseif sys_City="南投縣" and trim(Session("Unit_ID"))="05BA" then
'				if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
'					theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
'				end if
'				theMailNumber="&nbsp;"
			elseif sys_City="南投縣" then
				if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
					theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
				end if
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=left(right("000000000000000000" & trim(rsH("MailNumber")),14),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=left(right("000000000000000000" & trim(rsH("StoreAndSendMailNumber")),14),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end if
			else
				if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
					theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
				end if
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=trim(rsH("MailNumber"))&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=trim(rsH("StoreAndSendMailNumber"))&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end if
			end if
		else
			theMailNumber="&nbsp;"
		end if
		rsH.close
		set rsH=Nothing
		
		strList=strList&"<td align=""center"" class=""style33"">"&theMailNumber&"</td>"

		GetMailMem=""
		GetMailAddress=""
		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			
				if ExchangeTypeFlag="N" then
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress"))&"","臺","台"),ZipName,"")
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=Nothing
					If sys_City="苗栗縣" Then 
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							GetMailMem=trim(rsD2("Owner"))
						end if
						rsD2.close
						set rsD2=nothing
					End If 
					If sys_City="高雄市" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("DriverAddress")) Then
								GetMailAddress=trim(rs1("DriverZip"))&" "&trim(rs1("DriverAddress"))
							End If
						End If 
					End If
				Else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
							Set rsD3=conn.execute(strSqlD)
							If Not rsD3.eof Then
								If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")&"(戶)"
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
							Else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
							rsD3.close
							Set rsD3=Nothing 
						End if
					end if
					rsD2.close
					set rsD2=Nothing
					If sys_City="高雄市" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								GetMailAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
							End If
						End If 
					End If
				end If
			
		else	'攔停抓Driver
			if sys_City="高雄縣" then
				strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,Rule1 from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					RuleTarget=""
					strRule="select Target from Law where ItemID='"&trim(rsD("Rule1"))&"'"
					set rsRule=conn.execute(strRule)
					if not rsRule.eof then
						RuleTarget=trim(rsRule("Target"))
					end if
					rsRule.close
					set rsRule=nothing
					if RuleTarget="V" then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					else
						'沒Driver就抓Owner
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
							if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣" or sys_City="台南市" then
								ZipName=""
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							end if
								GetMailMem=trim(rsD("Driver"))
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
						'else
						'	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣" or sys_City="台南市" then
						'		ZipName=""
						'	else
						'		strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						'		set rsZip=conn.execute(strZip)
						'		if not rsZip.eof then
						'			ZipName=trim(rsZip("ZipName"))
						'		end if
						'		rsZip.close
						'		set rsZip=nothing
						'	end if
						'	if sys_City="台南市" then
						'		GetMailMem=trim(rsD("Owner"))
						'	else
						'		GetMailMem=trim(rsD("Driver"))
						'	end if
						'		GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
						end if
					end if
				end if
				rsD.close
				set rsD=nothing
			else
				strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					'沒Driver就抓Owner
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						if sys_City="基隆市" or sys_City="金門縣" or sys_City="嘉義縣" or sys_City="台南市" then
							ZipName=""
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						end if
							if isnull(rsD("Driver")) or trim(rsD("Driver"))="" Then
								If not isnull(rsD("Owner")) and trim(rsD("Owner"))<>"" Then
									GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
								End if
							else
								GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
					else
						if sys_City="基隆市" or sys_City="金門縣" or sys_City="嘉義縣" or sys_City="台南市" then
							ZipName=""
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						end if
						if sys_City="台南市" or sys_City="台中市" or sys_City="高雄市" Or sys_City=ApconfigureCityName then
							if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							end if
						elseif sys_City="宜蘭縣" or sys_City="澎湖縣" or sys_City="南投縣" or sys_City="台東縣" then
							if not isnull(rsD("Driver")) and trim(rsD("Driver"))<>"" then
								GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
							elseif not isnull(rsD("Owner")) and trim(rsD("Owner"))<>"" then
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							else
								GetMailMem="&nbsp;"
							end if
						else
							if isnull(rsD("Driver")) or trim(rsD("Driver"))="" then
								If not isnull(rsD("Owner")) and trim(rsD("Owner"))<>"" Then
									GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
								End if
							else
								GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
							end if
						end if
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					end if
				end if
				rsD.close
				set rsD=nothing
			end if
		end if
		'收件人姓名
		strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailMem,12,1)&"</td>"
			
		'收件地址
		strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailAddress,12,1)&"</td>"
		
		'備考=單號
		strList=strList&"<td align=""left"" class=""style33"">"&trim(rs1("BillNO"))&"</td>"
		strList=strList&"</tr>"
		rs1.MoveNext
	next
	if mailSN<PageCaseCnt then
		
			mailSNTmp=CaseSN
		
		for Sp=1 to PageCaseCnt-mailSN
			mailSNTmp=mailSNTmp+1
			strList=strList&"<tr>"
			'順序號碼
			strList=strList&"<td align=""center"" class=""style33"">"&mailSNTmp&"</td>"
			strList=strList&"<td align=""center"" class=""style33"">&nbsp;</td>"
			strList=strList&"<td align=""center"" class=""style33"">&nbsp;</td>"
			strList=strList&"<td align=""center"" class=""style33"">&nbsp;</td>"
			strList=strList&"<td align=""center"" class=""style33"">&nbsp;</td>"
			strList=strList&"</tr>"
		next
	end if

%>
<table width="710" align="center"  border="0" cellpadding="0">
<tr>

<td>
	<table width="100%" align="center" cellpadding="0" border="0">
		<tr>
			<td colspan="3" ><div align="center"><span class="style6">中 華 民 國 郵 政</span></div></td> 
		</tr>
		<tr>
			<td colspan="3" ><div align="right"><span class="style5"><%
			response.write InstrAdd(BatchNumberTmp,130)
			response.write "-"
			response.write pageNum
			%></span></div></td> 
		</tr>
		<tr>			
			<td width="37%" ><span class="style5"><%
			response.write Year(now)-1911
			%> 年 <%
			response.write Right("00"&month(now),2)
			%> 月 <%
			response.write Right("00"&day(now),2)
			%> 日 </span></td>
			<td width="26%" align="center"><span class="style5">交寄大宗掛號函件執聯</span></td>
			<td width="37%" align="right" ><span class="style5"><%

			%></span></td>
		</tr>
		<tr>			
			<td colspan="2"><span class="style5">交寄人名稱:<%
			If sys_City="苗栗縣" And RecordMemberIDTemp="3552" Then
				UnitName="苗栗市公所"
			End If 
			response.write UnitName
			%></span></td>
			<td align="right"><span class="style5">詳細地址:<%
			response.write UnitAddress
			%></span></td>

		</tr>
	</table>

</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="0">
	<tr>
	<td width="6%" ><div align="center"><span class="style5">序號</span></div></td>
	<td width="9%" ><div align="center"><span class="style5">掛號碼</span></div></td>
	<td width="22%" class="style5"><div align="center">收件人姓名</div></td>
	<td width="48%" class="style5"><div align="center">寄達地名(或地址)</div></td>
	<td width="8%" ><div align="center"><span class="style5">備註</span></div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
</table>

<%		
	
Wend
rs1.close
set rs1=nothing
%>			
</body>

<script language="javascript">
window.print();

</script>
</html>
