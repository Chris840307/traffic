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
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"嘉義縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>交寄大宗函件</title>
<script type="text/javascript" src="../js/Print.js"></script>
<%if sys_City="新北市" then %>
<script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../js/jquery-barcode-2.0.2.min.js"></script>
<%End If %>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
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
<%if sys_City="台東縣" then%>
	font-size: 8pt;
<%else%>
	font-size: 9pt;
<%end if%>
	line-height:10pt;
	font-family: "標楷體";
}
.style5 {
	font-size: 10pt;
	font-family: "標楷體";}
.style7 {
<%if sys_City="台東縣" then%>
	font-size: 9pt;
<%else%>
	font-size: 10pt;
<%end if%>
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
<%if sys_City="台東縣" then%>
	font-size: 10px;
<%else%>
	font-size: 10px;
<%end if%>
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
<%end if%>
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
strExchangeType="select a.ExchangeTypeID,f.BillUnitID,a.Batchnumber,f.DealLineDate from DciLog a,BillBase f where a.BillSN=f.SN "&_
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
else
	ExchangeTypeFlag="W"
	BillUnitIDtmp=""
end if
rsEType.close
set rsEType=nothing

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

if sys_City="台東縣" then
	PageCaseCnt=20
else
	PageCaseCnt=20
end if

if sys_City="南投縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and a.DciReturnStatusID in ('S','h') and e.ExchangeTypeID='W'" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8 and (f.EquipmentID<>'-1' or f.EquipmentID is null)) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
else
	
	strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
	" from DCILog a" &_
	",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
end If

set rs1=conn.execute(strSQL)

if sys_City="南投縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and a.DciReturnStatusID in ('S','N','h') and e.ExchangeTypeID='W'" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"

	else
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8 and (f.EquipmentID<>'-1' or f.EquipmentID is null)) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
else
	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
	",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
end if
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
		if sys_City="花蓮縣"  then
			strList=strList&"<tr height=""23"">"
		else
			strList=strList&"<tr>"		
		end if
		'順序號碼
		if sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕" then 
			strList=strList&"<td align=""center"">"&CaseSN&"</td>"
		elseif sys_City<>"雲林縣" and sys_City<>"台南縣" and sys_City<>"台南市" And sys_City<>ApconfigureCityName then
			strList=strList&"<td align=""center"">"&mailSN&"</td>"
		else
			strList=strList&"<td align=""center"">"&CaseSN&"</td>"
			if sys_City="台南縣" or sys_City="台南市" then
				if ExchangeTypeFlag="N" then
					strUpd="Update BillMailHistory set MailSeqNo2="&CaseSN&" where BillSN="&trim(rs1("BillSN"))
					conn.execute strUpd
				else
					strUpd="Update BillMailHistory set MailSeqNo1="&CaseSN&" where BillSN="&trim(rs1("BillSN"))
					conn.execute strUpd
				end if
			end if
		end if
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
		set rsH=nothing
		if ExchangeTypeFlag="N" and sys_City="台東縣" then
			strList=strList&"<td align=""center"">&nbsp;</td>"
		else
			strList=strList&"<td align=""center"">"&theMailNumber&"</td>"
		end if
		GetMailMem=""
		GetMailAddress=""
		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			if sys_City="南投縣" then
				if ExchangeTypeFlag="N" then
					'strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='N' and Status in ('Y','n','S') and DriverHomeAddress is not null"

					strSqlD="select * from BillbaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangetypeID='W'"
					set rsD=conn.execute(strSqlD)

					if not rsD.eof then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						If Not IsNull(rsD("DriverHomeAddress")) then
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
						End If 
					end if 
					rsD.close

					If ifnull(GetMailAddress) Then
						strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn " &_
						" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S'" &_
						" and Carno in (select carno from dcilog where BillSN="&trim(rs1("BillSN")) &_
						" and ExchangetypeID='A' and dcireturnstatusid='S')"
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
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
							else
								GetMailMem=trim(rsD("Owner"))
								GetMailAddress="(車)"&trim(rsD("OwnerZip"))&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
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
									GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress")),"臺","台"),ZipName,"")
								else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing

									GetMailMem=trim(rsD2("Owner"))
									GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
								end if
							end if
							rsD2.close
							set rsD2=nothing
						end if
						rsD.close
						set rsD=nothing
					End if
				else
					'入案先抓住就地,再抓查車driver,再抓入案車籍地
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
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn " &_
							" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S'" &_
							" and Carno in (select carno from dcilog where BillSN="&trim(rs1("BillSN")) &_
							" and ExchangetypeID='A' and dcireturnstatusid='S')"
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
					 '如果Billbase有寫以billbase為主
						'If trim(rs1("BillTypeID"))="2" Then
						'	If Not isnull(rs1("Owner")) Then
						'		GetMailMem=trim(rs1("Owner"))
						'	End If
						'	If Not isnull(rs1("OwnerAddress")) Then
						'		GetMailAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
						'	End If
						'End If 
				end If
			end if
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
								GetMailMem="&nbsp;"
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
		'收件人姓名
		if sys_City="花蓮縣"  then
			strList=strList&"<td align=""center"" width=""100"">&nbsp;</td>"
			strList=strList&"<td align=""left"" width=""100""class=""style35"">"&funcCheckFont(GetMailMem,14,1)&"</td>"
		else
			strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailMem,14,1)&"</td>"
		end if
			
		'收件地址
		if sys_City="花蓮縣"  then
			strList=strList&"<td align=""left"" class=""style35"" width=""300"">"&funcCheckFont(GetMailAddress,14,1)&"</td>"
		else
			strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailAddress,14,1)&"</td>"
		end if
		
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"" width=""20"">"&theMailMoneyTmp&"</td>"
		'備考=單號
		strList=strList&"<td align=""left"">"&trim(rs1("BillNO"))&"</td>"
		strList=strList&"</tr>"
		rs1.MoveNext
	next
	if mailSN<PageCaseCnt then
		if sys_City<>"雲林縣" and sys_City<>"台南縣" and sys_City<>"台南市" then
			mailSNTmp=mailSN
		else
			mailSNTmp=CaseSN
		end if
		for Sp=1 to PageCaseCnt-mailSN
			mailSNTmp=mailSNTmp+1
			if sys_City="花蓮縣"  then
				strList=strList&"<tr height=""23"">"
			else
				strList=strList&"<tr>"
			end if
			'順序號碼
			if sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕" then 
				strList=strList&"<td align=""center"">&nbsp;</td>"
			else
				strList=strList&"<td align=""center"">"&mailSNTmp&"</td>"
			end if
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"</tr>"
		next
	end if

if (sys_City="南投縣" And Trim(session("Unit_ID"))<>"05A7") or sys_City="雲林縣" or sys_City="宜蘭縣" then 
	ReportCount=3
elseif sys_City="花蓮縣" or sys_City="嘉義縣" then 
	ReportCount=1
else
	ReportCount=2
end if
if sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕" then 
	ReportCount=1
end if
if sys_City="宜蘭縣" and trim(Session("Unit_ID"))="TQ00" then 
	If sysBillTypeID=2 And ExchangeTypeFlag="W" Then
		theSendDocDate=Year(date)-1911 & Right("00"&Month(date),2) & Right("00"&day(date),2)
	End If 
end if

%>
<%if sys_City="新北市" then %>

<script type="text/javascript">
      $(function(){
	<% for Bi=1 to ReportCount
			BarCodeName="bcTarget"&pageNum&Bi
	%>
			$("#<%=BarCodeName%>").barcode("<%=MailBatchNumber%>", "code128",{barWidth:1, barHeight:30,fontSize:12,showHRI:true,bgColor:"#FFFFFF"});
	<%next%>
      });
</script>
<%End if%>
<table width="710" align="center"  border="0">
<tr>

<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"花蓮縣" and sys_City<>"嘉義縣" and sys_City<>"台東縣" then %>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>

	<tr>
<%if sys_City<>"花蓮縣" then %>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		
		</span></td>

		<td rowspan="3" width="39%" align="center"><span class="style7">

		<table width="100%">
	
		<tr>
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"1"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End If 
			%></div></td> 
		</tr>
		<%If sys_City="台東縣" Then %>
		<div id="num30" style="position:absolute; left:1;top:50;font-size: 36pt;line-height: 50pt;">
			<font face="標楷體"><b><%=RIGHT("000" &pageNum,3)%></b></font>
		<div>
		<%end if%>

		<tr>			
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件執據</td>
		</tr>

		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
<%end if%>
		</table>

	<%if sys_City<>"花蓮縣" then %>	
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	<%end if%>

	</tr>

	<tr>
		<td height="40" valign="top"><span class="style7">

<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
		 <br>
<%end if%>		
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日

<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
		中華民國 <%
		response.write year(MDate)-1911
		%>年 <%
		response.write right("00"&month(MDate),2)
		%>月 <%
		response.write right("00"&day(MDate),2)
		%>日

<%end if%>

		<br>
<%if sys_City="台南市" then %>	
		填單日期 <%
			if BillFillDateTmp<>"" then
				response.write year(BillFillDateTmp)-1911&"年 "
			end if
			if BillFillDateTmp<>"" then
				response.write month(BillFillDateTmp)&"月 "
			end if
			if BillFillDateTmp<>"" then
				response.write day(BillFillDateTmp)&"日"
			end if
		%>
<%elseif sys_City<>"澎湖縣" then %>	
		移送監理站日期 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write left(theSendDocDate,2)
				elseif len(theSendDocDate)=7 then
					response.write left(theSendDocDate,3)
				end if
			end if
		%>年 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,3,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,4,2)
				end if
			end if
		%>月 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,5,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,6,2)
				end if
			end if
		%>日
		<br>
<%end if%>
		<%
	if sys_City="南投縣" or sys_City="基隆市" or sys_City="台東縣" or sys_City="台中市"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>

		</span>

		</td>

	</tr>
<%if sys_City<>"花蓮縣" then %>	
	<tr>
		<td><span class="style7">
		寄件人 <%
		response.write UnitName
		%>
		</span></td>
	</tr>

	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>

<%else%>
	<tr><td><span class="style7">  <% response.write UnitName %> </span> </td>
	    <td> <span class="style7"><%response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日</span> 
	  
	   <td>
		<td width="34%"><span class="style7">
		<%=pageNum%> of <%=pagecnt%>
		</span></td>	
	</tr>	
	<tr>
	</tr>
<%end if%>
	</table>

</td>
</tr>
<tr>
<td>
    <%if sys_City<>"花蓮縣" then%>	
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	
    <%else%>
	<table align="center" width="100%" border="0" cellspacing="0" cellpadding="3">
	
    <%end if%>
   <tr>
    <%if sys_City<>"花蓮縣" then%>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
   
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>

	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>

	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
<%end if%>
	</tr>
	<tr>
<%if sys_City<>"花蓮縣" then%>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
<%end if%>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>

<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
<%if sys_City<>"花蓮縣" then%>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
<%end if%>

	<td width="34%" class="style5" valign="Top">
<%if sys_City<>"花蓮縣" then%>
	  <p>限時掛號<br>
<%else%>
	<br>
<%end if%>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤
<%if sys_City<>"花蓮縣" then%>
		<br>
	    快捷郵件<br>
		<%if sys_City<>"台東縣" then%>
		<br>
		<%end if%>
<%else%>
 ( 
<%end if%>	    
	    
	   郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元 
	  <%if sys_City<>"花蓮縣" then%>
		</p><p align="right"><%
		If sys_City="台中市" then
			response.write Trim(DealLineDateTmp)&"  "
		End If 
		%>______________<br>經辦員簽署&nbsp; </p>
	  <%else%>
		)	
	  <%end if%>
	  </td>
	</tr>
	</table>
</td>
</tr>

</table>


<%if ReportCount>1 then %>
<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" then%>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>
	<tr>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>
		<td rowspan="3" width="39%" align="center"><span class="style7">
		<table width="100%">
		<tr>
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"2"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End If 
			%></div></td> 
		</tr>
		<%If sys_City="台東縣" Then %>
		<div id="num30" style="position:absolute; left:70;top:50;font-size: 36pt;line-height: 50pt;">
			<font face="標楷體"><b><%=RIGHT("000" &pageNum,3)%></b></font>
		<div>
		<%end if%>
		<tr>
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件存根</td>
		</tr>
		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
		</table>
		
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	</tr>
	<tr>
		<td height="40" valign="top"><span class="style7">
<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
		 <br>
<%end if%>		
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日
<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
		中華民國 <%
		response.write year(MDate)-1911
		%>年 <%
		response.write right("00"&month(MDate),2)
		%>月 <%
		response.write right("00"&day(MDate),2)
		%>日
<%end if%>
		<br>
<%if sys_City="台南市" then %>	
		填單日期 <%
			if BillFillDateTmp<>"" then
				response.write year(BillFillDateTmp)-1911&"年 "
			end if
			if BillFillDateTmp<>"" then
				response.write month(BillFillDateTmp)&"月 "
			end if
			if BillFillDateTmp<>"" then
				response.write day(BillFillDateTmp)&"日"
			end if
		%>
<%elseif sys_City<>"澎湖縣" then %>	
		移送監理站日期 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write left(theSendDocDate,2)
				elseif len(theSendDocDate)=7 then
					response.write left(theSendDocDate,3)
				end if
			end if
		%>年 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,3,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,4,2)
				end if
			end if
		%>月 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,5,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,6,2)
				end if
			end if
		%>日
		<br>
<%end if%>
		<%
	if sys_City="南投縣"  or sys_City="基隆市" or sys_City="台東縣" or sys_City="台中市"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人 <%=UnitName%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	<tr>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
	</tr>
	<tr>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
	<td width="34%" class="style5" valign="Top">
	  <p>限時掛號<br>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤<br>
	    快捷郵件<br>
	    
	    <%if sys_City<>"台東縣" then%>
		<br>
		<%end if%>
	    郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元	  </p>
	  <p align="right"><%
		If sys_City="台中市" then
			response.write Trim(DealLineDateTmp)&"  "
		End If 
		%>______________<br>經辦員簽署&nbsp; </p>
	  </td>
	</tr>
	</table>
</td>
</tr>
</table>
<%end if%>
<%if ReportCount=3 then %>

<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" then%>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>
	<tr>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>
		<td rowspan="3" width="39%" align="center"><span class="style7">
		<table width="100%">
		<tr>
			<td colspan="3" height="28"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"3"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End if
			%></div></td> 
		</tr>
		<tr>
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件存根</td>
		</tr>
		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
		</table>
		
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	</tr>
	<tr>
		<td height="40" valign="top"><span class="style7">
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日
<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
		中華民國 <%
		response.write year(MDate)-1911
		%>年 <%
		response.write right("00"&month(MDate),2)
		%>月 <%
		response.write right("00"&day(MDate),2)
		%>日
<%end if%>
		<br>
<%if sys_City="台南市" then %>	
		填單日期 <%
			if BillFillDateTmp<>"" then
				response.write year(BillFillDateTmp)-1911&"年 "
			end if
			if BillFillDateTmp<>"" then
				response.write month(BillFillDateTmp)&"月 "
			end if
			if BillFillDateTmp<>"" then
				response.write day(BillFillDateTmp)&"日"
			end if
		%>
<%elseif sys_City<>"澎湖縣" then %>	
		移送監理站日期 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write left(theSendDocDate,2)
				elseif len(theSendDocDate)=7 then
					response.write left(theSendDocDate,3)
				end if
			end if
		%>年 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,3,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,4,2)
				end if
			end if
		%>月 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,5,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,6,2)
				end if
			end if
		%>日
<%end if%>
		<br>
		<%
	if sys_City="南投縣"  or sys_City="基隆市" or sys_City="台東縣"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人 <%=UnitName%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	<tr>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
	</tr>
	<tr>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
	<td width="34%" class="style5" valign="Top">
	  <p>限時掛號<br>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤<br>
	    快捷郵件<br>
	    
	    <br>
	    郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元	  </p>
	  <p align="right">______________<br>經辦員簽署&nbsp; </p>
	  </td>
	</tr>
	</table>
</td>
</tr>
</table>
<%end if%>
<%		
	
Wend
rs1.close
set rs1=nothing
%>			
</body>

<script language="javascript">
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" or sys_City="花蓮縣" then%>
window.print();
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
</html>
