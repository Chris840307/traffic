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
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)
%>
<style type="text/css">
<!--

.style35 {
	font-size: 8pt;
	line-height:11px;
}
.style33 {
	font-size: 9pt;
	font-family: "標楷體";
}
.style5 {
	font-size: 10pt;
	font-family: "標楷體";}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 10pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
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
strExchangeType="select a.ExchangeTypeID,f.BillUnitID,a.Batchnumber from DciLog a,BillBase f where a.BillSN=f.SN "&_
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
else
	ExchangeTypeFlag="W"
	BillUnitIDtmp=""
end if
rsEType.close
set rsEType=Nothing

MailBatchNumber=""
strBatch="select distinct(a.Batchnumber) from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 "&strwhere
set rsBatch=conn.execute(strBatch)
If Not rsBatch.Bof Then rsBatch.MoveFirst 
While Not rsBatch.Eof
	If MailBatchNumber="" Then
		MailBatchNumber=trim(rsBatch("BatchNumber"))
	Else
		MailBatchNumber=MailBatchNumber&","&trim(rsBatch("BatchNumber"))
	End If 

	rsBatch.MoveNext
Wend
rsBatch.close
set rsBatch=Nothing

'台中市停管
if sys_City="台中市" and stopBatchnumber="WT" then
	strwhere=strwhere&" and (f.Note like '2%')"
end if
if sys_City="台中市" or sys_City="高雄市" then 
	if BillUnitIDtmp="" then
		strSendMailUnit="select b.UnitName,b.Address,b.Tel from Apconfigure a,UnitInfo b " &_
				" where a.ID=49 and a.Value=b.UnitID"
		set rsSendMailUnit=conn.execute(strSendMailUnit)
		if not rsSendMailUnit.eof then
			
			if sys_City<>"花蓮縣" and sys_City<>"台中市" and sys_City<>"高雄市" then 
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
			if trim(rsShow("ShowOrder"))="0" or trim(rsShow("ShowOrder"))="1" or trim(rsShow("UnitID"))="046A" or trim(rsShow("UnitID"))="0463" or trim(rsShow("UnitID"))="0464" or trim(rsShow("UnitID"))="0465" or trim(rsShow("UnitID"))="0469" or trim(rsShow("UnitID"))="0561" then
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
	if sys_City="台中市" Then
		If UnitName="交通警察大隊直屬第一分隊" Or UnitName="交通警察大隊直屬第三分隊" Then
			UnitName="交通警察大隊第一中隊"
			UnitTel="(04)23274655"
			UnitAddress="407台中市西屯區大隆路192號"
		ElseIf UnitName="交通警察大隊直屬第二分隊" then
			UnitName="交通警察大隊第二中隊"
		End If 
	End if
else
	strSendMailUnit="select b.UnitName,b.Address,b.Tel from MemberData a,UnitInfo b " &_
			" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
	set rsSendMailUnit=conn.execute(strSendMailUnit)
	if not rsSendMailUnit.eof then
		
		if sys_City<>"高雄縣" and sys_City<>"台中市" and sys_City<>"高雄市" then 
			UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
		else
			UnitName=trim(rsSendMailUnit("UnitName"))
		end if
		UnitAddress=trim(rsSendMailUnit("Address"))
		UnitTel=trim(rsSendMailUnit("Tel"))
	end if
	rsSendMailUnit.close
	set rsSendMailUnit=Nothing
	
end if
if sys_City="台中市" or sys_City="高雄市" then 
	if ExchangeTypeFlag="N" Then
		If sys_City="台中市" then
			strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.DealLineDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and (e.Status in ('S','N','h') or (e.Status='n' and e.billcloseid='j')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
		Else
			strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.DealLineDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N','h'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
		End If 
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.DealLineDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" and NVL(f.EquiPmentID,1)<>-1 order by f.RecordMemberID,f.RecordDate"
	end if
end if
set rs1=conn.execute(strSQL)
if sys_City="台中市" or sys_City="高雄市" then 
	if ExchangeTypeFlag="N" Then
		If sys_City="台中市" Then
			strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and (e.Status in ('S','N','h') or (e.Status='n' and e.billcloseid='j')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
		Else
			strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N','h'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
		End If 
		
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) and NVL(f.EquiPmentID,1)<>-1 "&strwhere	
	end if
end if
set rsCnt=conn.execute(strCnt)
if not rsCnt.eof then
	if trim(rsCnt("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rsCnt("cnt"))/20+0.9999999)
	end if
end if
rsCnt.close
set rsCnt=nothing
'response.write strSQL

MDate=""
if  sys_City<>"高雄市" then
	if ExchangeTypeFlag="N" then
		strMailDate="select g.StoreAndSendMailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
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
end if
	if MDate="" or isnull(MDate) then
		MDate=now
	end if

CaseSN=0
mailSNTmp=0
DealLineDateTmp=""
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	BillFillDateTmp=""
	if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
		BillFillDateTmp=trim(rs1("BillFillDate"))
	end if
	if DealLineDateTmp="" then
		if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
			DealLineDateTmp=year(rs1("DealLineDate"))-1911&"/"&Month(rs1("DealLineDate"))&"/"&day(rs1("DealLineDate"))
		end if
	end if
	strList=""
	mailSN=0
		strList=strList&"<br>"
		strList=strList&"<br>"
		strList=strList&"<br>"
	
	pageNum=fix(CaseSN/20)+1
	for i=1 to 20
		if rs1.eof then exit for
		'MailBatchNumber=trim(rs1("BatchNumber"))
		mailSN=mailSN+1
		CaseSN=CaseSN+1
		if  sys_City="台中市" or sys_City="高雄市" then
			strList=strList&"<tr height=""26"">"
		else
			strList=strList&"<tr>"		
		end if
		'順序號碼
		strList=strList&"<td align=""center""></td>"

		'掛號號碼
		theMailNumber=""
		'移送監理站日期
		theSendDocDate=""
		strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from BillMailHistory where BillSN="&trim(rs1("BillSN"))
		set rsH=conn.execute(strSqlH)
		if not rsH.eof then
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
		else
			theMailNumber="&nbsp;"
		end if
		rsH.close
		set rsH=nothing
		strList=strList&"<td align=""center"" width=""45""></td>"
		strList=strList&"<td align=""center"" width=""80"">"&theMailNumber&"</td>"
		GetMailAddress=""
		ZipName=""
		GetMailMem=""
		GetMailAddress=""
		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			if sys_City="台中市"  then	'台中入案不要抓車籍查詢
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
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
						ElseIf trim(rsD("OwnerAddress"))<>"" then
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
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&" ","臺","台"),ZipName,"")
					end if
					rsD2.close
					set rsD2=Nothing
					
					If GetMailMem="" Then
						GetMailMem=trim(rs1("Owner"))
					End If
					If GetMailZip="" Then
						GetMailZip=trim(rs1("OwnerZip"))
						strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
					End If
					If GetMailAddress="" Then
						GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress")&"")&" ","臺","台"),ZipName,"")
					End If

				end If
			ElseIf sys_City="高雄市" Then
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
					set rsD=Nothing
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
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress")),"臺","台"),ZipName,"")&"(戶)"
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
								End If
							Else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
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
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress")),"臺","台"),ZipName,"")

							End If
						End If 
					End If
				end If
			end if
		else	'攔停抓Driver
			strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
			set rsD=conn.execute(strSqlD)
			if not rsD.eof then
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
						if trim(rsD("Driver"))<>"" and not isnull(rsD("Driver")) then
							GetMailMem=trim(rsD("Driver"))
						else
							GetMailMem=trim(rsD("Owner"))
						end if
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
				else
					if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣" or sys_City="台南市" then
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
					if sys_City="台南市" or sys_City="台中市" or sys_City="高雄市" then
						GetMailMem=trim(rsD("Owner"))
					else
						GetMailMem=trim(rsD("Driver"))
					end if
						GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&" ","臺","台"),ZipName,"")
				end if
			end if
			rsD.close
			set rsD=nothing
		end if
		'收件人姓名
		'If sys_City="高雄市" Then
			GetMailMem=funcCheckFont(GetMailMem,15,1)
			GetMailAddress=funcCheckFont(GetMailAddress,15,1)
		'end if
		
		if GetMailMem="" then
			strList=strList&"<td align=""left"" width=""70""><span class=""style35"">"&GetMailMem&"</span></td>"
		else
			if len(GetMailMem)>4 and instr(GetMailMem,"<img")=0 then
				strList=strList&"<td align=""left"" width=""70""><span class=""style35"">"&left(GetMailMem,12)&"</span></td>"
			else
				strList=strList&"<td align=""left"" width=""70""><span class=""style8"">"&GetMailMem&"</span></td>"
			end if
		end if
			
		'收件地址
'		if GetMailAddress="" then
			strList=strList&"<td align=""left"" width=""390""><span class=""style8"">"&GetMailAddress&"</span></td>"
'		else
'			if len(GetMailAddress)>20 then
'				strList=strList&"<td align=""left"" width=""280""><span class=""style35"">"&GetMailAddress&"</span></td>"
'			else
'				strList=strList&"<td align=""left"" width=""280""><span class=""style8"">"&GetMailAddress&"</span></td>"
'			end if
'		end if

		'strList=strList&"<td align=""center"" width=""120""></td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"" width=""50"">"&theMailMoneyTmp&"</td>"
		'備考=單號
		strList=strList&"<td align=""left"">"&trim(rs1("BillNO"))&"</td>"
		strList=strList&"</tr>"
		rs1.MoveNext
	next
	if mailSN<20 then
		if sys_City<>"雲林縣" and sys_City<>"台南縣" and sys_City<>"台南市" then
			mailSNTmp=mailSN
		else
			mailSNTmp=CaseSN
		end if
		for Sp=1 to 20-mailSN
			mailSNTmp=mailSNTmp+1
			if sys_City="高雄縣" or sys_City="台中市" or sys_City="高雄市" then
				strList=strList&"<tr height=""26"">"
			else
				strList=strList&"<tr>"
			end if
			'順序號碼
			'strList=strList&"<td align=""center"">&nbsp;</td>"
			'strList=strList&"<td align=""center"">&nbsp;</td>"
			'strList=strList&"<td align=""center"">&nbsp;</td>"
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


%>

<!-- smith for 高雄縣-->

<br>
<br>
<br>
<br>




<table width="780" border="0">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
	<tr>
		<td height="<%
		if sys_City="高雄市" then
			response.write "10"
		else
			response.write "21"
		end if
		%>"></td>
	</tr>
	<tr>
		<td width="7%"></td>
		<td><span class="style7">
		</span></td>
		<td></td>
		<td width="25%" align="right"><span class="style7">
		<%
		if ExchangeTypeFlag="N" then
			response.write "單退"
		end if
		%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</span></td>	
	</tr>
	<tr>
		<td width="7%"></td>
		<td valign="top"><span class="style7">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%response.write year(MDate)-1911
		%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<%
		response.write right("00"&month(MDate),2)
		%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%
		response.write right("00"&day(MDate),2)
		%>
		</span></td>
		<td><%response.write MailBatchNumber%></td>
		<td width="25%" align="right" valign="top"><span class="style7">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>	
	</tr>
	<!--smith 高雄縣-->
	<tr>
		<td width="7%"></td>
		<td ><span >&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<% response.write UnitName %> </span> </td>
	    <td> &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<% response.write UnitAddress %>  </td>
	  
	
		<td></td>	
	</tr>	
	<tr>
	</tr>

	</table>
</td>
</tr>
<tr>
<td>
	<table width="100%" border="0" cellspacing="0" cellpadding="3">
	<%=strList%>
	</table>
</td>
</tr>

<tr>
<td>
	<table width="100%" border="0">
		<tr>
		<td width="84%" height="30"></td>
		<td width="16%" valign="bottom"><%=mailSN%> </td>
		<tr>
		<tr>
		<td width="86%" height="33"></td>
		<td width="14%" valign="bottom"><%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%></td>
		<tr>
		<tr>
		<td width="86%" height="40"></td>
		<td width="14%" valign="bottom"><%
		if DealLineDateTmp<>"" then
			if ExchangeTypeFlag<>"N" then
				response.write DealLineDateTmp
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<tr>
	</table>
</td>
</table>


<%if sys_City<>"高雄縣" and sys_City<>"嘉義縣" and sys_City<>"台中市" and sys_City<>"高雄市" then %>
<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" then%>
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
			<td colspan="3" height="30"><div align="center"><u><span class="style6">臺 灣 郵 政</span></u></div></td> 
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
<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
<%end if%>
		<br>
<%if sys_City<>"雲林縣" then %>
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
	if sys_City="南投縣" then
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
<%if sys_City="南投縣" or sys_City="雲林縣" or sys_City="台南市" or sys_City="台東縣" then %>

<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" then%>
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
			<td colspan="3" height="30"><div align="center"><u><span class="style6">臺 灣 郵 政</span></u></div></td> 
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
<%if sys_City<>"雲林縣" then %>
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
	if sys_City="南投縣" then
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
<%if sys_City="高雄市" then%>
printWindow(true,5.08,0,5.08,0);
<%elseif sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" or sys_City="高雄縣" then%>
window.print();
<%else%>
printWindow(true,5.08,5.08,5.08,5.08);
<%end if%>
</script>
</html>
