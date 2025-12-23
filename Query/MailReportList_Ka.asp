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
	font-size: 10pt;
	font-family: "標楷體";
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
	font-size: 14pt;
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
strExchangeType="select a.ExchangeTypeID from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 "&strwhere
set rsEType=conn.execute(strExchangeType)
if not rsEType.eof then
	if trim(rsEType("ExchangeTypeID"))="N" then
		ExchangeTypeFlag="N"
	else
		ExchangeTypeFlag="W"
	end if
else
	ExchangeTypeFlag="W"
end if
rsEType.close
set rsEType=nothing

strSendMailUnit="select b.UnitID,b.UnitName,b.Address,b.Tel from MemberData a,UnitInfo b " &_
		" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
set rsSendMailUnit=conn.execute(strSendMailUnit)
if not rsSendMailUnit.eof then
	
	if sys_City<>"高雄縣" and sys_City<>"台中市" then 
		UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
	Else
		If sys_City="高雄縣" And trim(rsSendMailUnit("UnitID"))="8J00" Then
			UnitName="停管中心(交通隊)"
		else
			UnitName=trim(rsSendMailUnit("UnitName"))
		End if
	end if
	UnitAddress=trim(rsSendMailUnit("Address"))
	UnitTel=trim(rsSendMailUnit("Tel"))
end if
rsSendMailUnit.close
set rsSendMailUnit=nothing

if sys_City="高雄縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
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
end if
set rs1=conn.execute(strSQL)
if sys_City="高雄縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and (e.ExchangeTypeID='N' and d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
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
		pagecnt=fix(Cint(rsCnt("cnt"))/20+0.9999999)
	end if
end if
rsCnt.close
set rsCnt=nothing
'response.write strSQL

MDate=""
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
		'<!-- smith 高雄縣 -->
		strList=strList&"<br>"
		strList=strList&"<br>"
		strList=strList&"<br>"
	
	pageNum=fix(CaseSN/20)+1
	for i=1 to 20
		if rs1.eof then exit for
		MailBatchNumber=trim(rs1("BatchNumber"))
		mailSN=mailSN+1
		CaseSN=CaseSN+1
		if sys_City="高雄縣" or sys_City="台中市" then
			strList=strList&"<tr height=""26"">"
		else
			strList=strList&"<tr>"		
		end if
		'順序號碼
		if sys_City<>"雲林縣" and sys_City<>"台南縣" and sys_City<>"台南市" then
			strList=strList&"<td align=""center""></td>"
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
		else
			theMailNumber="&nbsp;"
		end if
		rsH.close
		set rsH=nothing
		strList=strList&"<td align=""center"" width=""18""></td>"
		strList=strList&"<td align=""center"" width=""80"">"&theMailNumber&"</td>"
		GetMailMem=""
		GetMailAddress=""
		ZipName=""
		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			if ExchangeTypeFlag="N" then	'單退先抓A的driver，沒有的話再抓W的Driver,再沒有就抓W的owner
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress="(戶)"&trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
					else
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
					end if
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
'							set rsZip=conn.execute(strZip)
'							if not rsZip.eof then
'								ZipName=trim(rsZip("ZipName"))
'							end if
'							rsZip.close
'							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress="(戶)"&trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
'							set rsZip=conn.execute(strZip)
'							if not rsZip.eof then
'								ZipName=trim(rsZip("ZipName"))
'							end if
'							rsZip.close
'							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
					end if
					rsD2.close
					set rsD2=nothing
				end if
				rsD.close
				set rsD=nothing
			else	'入案直接抓W的Owner
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
				
'						strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
'						set rsZip=conn.execute(strZip)
'						if not rsZip.eof then
'							ZipName=trim(rsZip("ZipName"))
'						end if
'						rsZip.close
'						set rsZip=nothing
					if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
						GetMailMem="&nbsp;"
					else
						GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
					end if
					GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
				end if
				rsD2.close
				set rsD2=nothing
			end if
			
		else	'攔停抓Driver
			
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
		end if
		'收件人姓名
			strList=strList&"<td align=""left"" width=""460"">"&funcCheckFont(GetMailMem,15,1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&funcCheckFont(GetMailAddress,15,1)&"</td>"
			
		'收件地址
			'strList=strList&"<td align=""left"" width=""315"">"&GetMailAddress&"</td>"

		'strList=strList&"<td align=""center"">&nbsp;</td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		'strList=strList&"<td align=""center"">&nbsp;</td>"
		
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"" width=""40"">"&theMailMoneyTmp&"</td>"
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
			if sys_City="高雄縣" or sys_City="台中市" then
				strList=strList&"<tr height=""26"">"
			else
				strList=strList&"<tr>"
			end if
			'順序號碼
			'strList=strList&"<td align=""center"">&nbsp;</td>"
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


%>

<!-- smith for 高雄縣-->

<br>
<br>
<br>
<table width="780" border="0">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">

	<tr>
		<td width="4%"></td>
		<td valign="top" width="35%"><span class="style7">
		
		
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%response.write year(now)-1911
		%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%
		response.write right("00"&month(now),2)
		%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%
		response.write right("00"&day(now),2)
		%>
		<br>
		</span></td>
		<td></td>
		<td width="10%" align="left">
		<span class="style7">
		<%=pageNum%> of <%=pagecnt%>
		</span></td>	
	</tr>
	<!--smith 高雄縣-->
	<tr>
		<td width="4%"></td>
		<td ><span >&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<% response.write UnitName %> </span> </td>
	    <td> &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp<% response.write UnitAddress %>  </td>
	  
	
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
		<td width="80%" height="30"></td>
		<td width="20%" valign="bottom"><%=mailSN%> </td>
		<tr>
		<tr>
		<td width="78%" height="33"></td>
		<td width="22%" valign="bottom"><%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%></td>
		<tr>
	</table>
</td>
</table>


<%if sys_City<>"高雄縣" and sys_City<>"嘉義縣" and sys_City<>"台中市" then %>
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
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
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
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" or sys_City="高雄縣" then%>
window.print();
<%else%>
printWindow(true,15.08,10.08,5.08,5.08);
<%end if%>
</script>
</html>
