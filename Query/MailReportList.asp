
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
'response.write strwhere
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
		
		if sys_City="花蓮縣" or sys_City="保二總隊三大隊二中隊" then 
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
	strB="select distinct(a.BatchNumber) " &_
	" from DCILog a" &_
	",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
	set rsB=conn.execute(strB)
	While Not rsB.Eof
		strBDel="Delete from batchnumberjob where batchNumber='"&Trim(rsB("Batchnumber"))&"' and PrintTypeID=0"
		conn.execute strBDel

		strBIns="Insert into batchnumberjob values('"&Trim(rsB("Batchnumber"))&"',"&Trim(session("User_ID"))&",2,sysdate)"
		conn.execute strBIns
	rsB.MoveNext
	Wend
	rsB.close
	Set rsB=Nothing 
End If 

if sys_City="台東縣" then
	PageCaseCnt=20
else
	PageCaseCnt=20
end if

if sys_City="基隆市" then 
	strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
	",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
	" from DCILog a" &_
	",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and ((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"

elseif sys_City="澎湖縣" or sys_City="雲林縣" then 
	if ExchangeTypeFlag="N" then

		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
elseif sys_City="南投縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and a.DciReturnStatusID in ('S','N','h','c') and e.ExchangeTypeID='W'" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber,(select mailnumber from billmailhistory where billsn=a.billsn) mailnumber " &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8 and (f.EquipmentID<>'-1' or f.EquipmentID is null)) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by mailnumber,f.RecordDate"
	end if
elseif sys_City="台中縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and e.Status in ('S','N')" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
elseif sys_City="花蓮縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N','h'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
elseif sys_City="台中市" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and (e.Status in ('S','N','h') or (e.Status='n' and e.billcloseid='j')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
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
elseif sys_City="台南市" then
	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N','n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else	'交通隊說要用大宗掛號碼排序 1100119
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress,g.mailnumber" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.mailnumber,f.RecordDate"
	end if
elseif sys_City="台南縣" or sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="嘉義市" or sys_City="台東縣" Or sys_City="苗栗縣" then
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
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"
	end if
elseif sys_City="宜蘭縣" Then
	if ExchangeTypeFlag="N" Then
		if Trim(session("Unit_ID"))="TQ00" Or Trim(session("Unit_ID"))="TP00" Then
			ReturnStatusPrint="'S','N','n'"
		Else
			ReturnStatusPrint="'S','N'"
		End If 
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N' and e.Status in (" & ReturnStatusPrint & ")))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	Else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	End If 
	
else
	
	strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.BillFillDate,f.RecordDate,a.BatchNumber" &_
	",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
	" from DCILog a" &_
	",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"
end If
If  sys_City="台南市" Then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	strI="insert into Log values((select nvl(max(Sn),0)+1 from Log),360,"&Trim(Session("User_ID"))&",'"&Trim(Session("Ch_Name"))&"','"&userip&"',sysdate,'大宗交寄函件,"&Replace(strSQL,"'","""")&"')"
	'response.write strI
	Conn.execute strI
End If 
set rs1=conn.execute(strSQL)
if sys_City="基隆市" then 
	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
	",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and ((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
elseif sys_City="澎湖縣" or sys_City="雲林縣" then
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
elseif sys_City="南投縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and a.DciReturnStatusID in ('S','N','h','c') and e.ExchangeTypeID='W'" &_
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
elseif sys_City="台中縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and e.Status in ('S','N')" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
elseif sys_City="花蓮縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N','h'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
elseif sys_City="台中市" then
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and (e.Status in ('S','N','h') or (e.Status='n' and e.billcloseid='j')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) and NVL(f.EquiPmentID,1)<>-1 "&strwhere	
	end if
elseif sys_City="台南市" then
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and e.Status in ('S','N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere	
	end if
elseif sys_City="台南縣" then
	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
	",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
elseif sys_City="宜蘭縣" then
	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
	",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
	" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
	" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
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
		" where f.Sn=g.BillSn and f.Sn=a.BillSn and f.recordstateid=0 "&strwhere
else
	strMailDate="select g.MailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn and f.recordstateid=0 "&strwhere
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
		if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
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
		strSqlH="select MailNumber,mailchknumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from BillMailHistory where BillSN="&trim(rs1("BillSN"))
		set rsH=conn.execute(strSqlH)
		if not rsH.eof Then
			if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
				theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
			end if
			if sys_City="台中市" or sys_City="雲林縣" then
				
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=right("00000000" & trim(rsH("MailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=right("000000" & trim(rsH("StoreAndSendMailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end If
			ElseIf sys_City="保二總隊四大隊二中隊" Then	'南科
				if trim(rs1("ExchangeTypeID"))="W" Then
					if trim(rsH("mailchknumber"))<>"" and not isnull(rsH("mailchknumber")) Then
						theMailNumber=left(Replace(trim(rsH("mailchknumber"))," ",""),14)&"&nbsp;"
					elseif trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=trim(rsH("MailNumber"))&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=left(Replace(trim(rsH("StoreAndSendMailNumber"))," ",""),14)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end if
'			elseif sys_City="南投縣" and trim(Session("Unit_ID"))="05BA" then

'				theMailNumber="&nbsp;"
			elseif sys_City="南投縣" then
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
				end If
			elseif sys_City="嘉義縣" Then
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
						theMailNumber=Right("000000"&trim(rsH("MailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				elseif trim(rs1("ExchangeTypeID"))="N" then
					if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
						theMailNumber=Right("000000"&trim(rsH("StoreAndSendMailNumber")),6)&"&nbsp;"
					else
						theMailNumber="&nbsp;"
					end if
				else
					theMailNumber="&nbsp;"
				end if
			elseif sys_City="嘉義市" Then
				if trim(rs1("ExchangeTypeID"))="W" then
					if trim(rsH("MAILCHKNUMBER"))<>"" and not isnull(rsH("MAILCHKNUMBER")) then
						theMailNumber=replace(trim(rsH("MAILCHKNUMBER"))," ","")&"&nbsp;"
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
		ZipName=""
		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			'-------------------------------------------------------------------------------------------
			if sys_City="台東縣" then	'(OLD)1.先抓Owner有（就、住） 2.DrvierAddress 3.OwnerAddress
				'(NEW 2015/5/20)(入案)住居地 -> (查車)戶籍地 -> (入案)車籍地
				Response.flush
				strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,dcierrorcardata,Nwner,NwnerZip,NwnerAddress from BIllBaseDCIReturn where (BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') and ExchangeTypeID='W' and Status in('Y','S','n','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof Then
					if ExchangeTypeFlag="N" Then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
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
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
						end If
					Else
						If Trim(request("CitySpecFlag"))="TD01" Then '台東肇事要抓入案的駕駛
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							GetMailMem=trim(rsD("Driver"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress"))&"","臺","台"),ZipName,"")
						Else
							'---------------------------------
							if instr(trim(rsD("OwnerAddress")),"(住)")>1 or instr(trim(rsD("OwnerAddress")),"(就)")>1 or instr(trim(rsD("OwnerAddress")),"（住）")>1 or instr(trim(rsD("OwnerAddress")),"（就）")>1 Or instr(trim(rsD("OwnerAddress")),"(通)")>1 or instr(trim(rsD("OwnerAddress")),"（通）")>1 then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								GetMailMem=trim(rsD("Owner"))
								GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
							else
								strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where Exists (select carno from dcilog where BillSN="&trim(rs1("BillSN"))&" and CarNo='"&trim(rs1("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S'"
								Set rsD3=conn.execute(strSqlD)
								If Not rsD3.eof Then
									If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
										GetMailMem=trim(rsD("Owner"))

										strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=Nothing
										
										GetMailAddress=trim(rsD3("DriverHomeZip"))&ZipName&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")
									Else
										strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=nothing
										GetMailMem=trim(rsD("Owner"))
										GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
									End If
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									GetMailMem=trim(rsD("Owner"))
									GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
								rsD3.close
								Set rsD3=Nothing 
							end If
							'-------------------------------------
						End if
					End If 
				end if
				rsD.close
				set rsD=Nothing
			'-------------------------------------------------------------------------------------------
			elseif sys_City="宜蘭縣" Then
				strSqlD="select Sn,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from BIllBase where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and Recordstateid=0"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if ExchangeTypeFlag="N" then	
						
		
						GetMailMem=trim(rsD("Owner"))
						If trim(rsD("DriverAddress") &"")<>"" Then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
							GetMailAddress=trim(rsD("DriverZip"))&ZipName&replace(replace(trim(rsD("DriverAddress") &""),"臺","台"),ZipName,"")
						Else	'TITAN沒寫BILLBASE DriverAddress的話,先抓A DriverHomeAddress,再抓W DriverHomeAddress
							strSqlDciA="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where Carno in (select carno from dcilog where BillSN="&trim(rsD("SN"))&" and ExchangetypeID='A') and ExchangeTypeID='A' and Status='S'"
							set rsDciA=conn.execute(strSqlDciA)
							if not rsDciA.eof then
								if trim(rsDciA("DriverHomeAddress"))<>"" and not isnull(rsDciA("DriverHomeAddress")) Then
									strZip="select ZipName from Zip where ZipID='"&trim(rsDciA("DriverHomeZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									GetMailAddress=trim(rsDciA("DriverHomeZip"))&ZipName&replace(replace(trim(rsDciA("DriverHomeAddress") &""),"臺","台"),ZipName,"")

								end If
							else
								strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='W' "
								set rsD2=conn.execute(strSqlD2)
								if not rsD2.eof then
									if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
										strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=nothing

										GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")

									end if
								end if
								rsD2.close
								set rsD2=nothing
							end if
							rsDciA.close
							set rsDciA=nothing
						End If 
						

					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress") &""),"臺","台"),ZipName,"")
						'如果Billbase有寫以billbase為主
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
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
					End if
				end if
				rsD.close
				set rsD=Nothing
			'-------------------------------------------------------------------------------------------
			elseif sys_City="花蓮縣" then
				if ExchangeTypeFlag="N" then	'單退先抓A的driver，沒有的話再抓W的Driver,再沒有就抓W的owner
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rs1("BillSN"))&" and ExchangetypeID='A') and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							'GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
						else
							'GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
						end If
						
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof Then
							GetMailMem=trim(rsD2("Owner"))
						end if
						rsD2.close
						set rsD2=nothing
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
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
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=nothing
				else	'XXXX入案先抓A的OwnerNotifyAddress 2.W owner 3.W driver(2012/1/12)XXX
						'入案先抓A的OwnerNotifyAddress 3.A driver 3.W owner (2021/3/16)
					BitchHL=0
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,OwnerNotifyAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S' and Exists (select Billsn from Dcilog where CarNo=BIllBaseDCIReturn.CarNo and CarNo='"&trim(rs1("CarNo"))&"' and Billsn="&trim(rs1("BillSn"))&" and ExchangeTypeID='A')"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("OwnerNotifyAddress"))<>"" and not isnull(rsD("OwnerNotifyAddress")) then
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerNotifyAddress"))
						ElseIf  trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
							GetMailMem=trim(rsD("Owner"))
							If trim(rsD("DriverHomeZip"))<>"" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
							End if
							if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
								GetMailMem=" &nbsp;"
							else
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress"))&"","臺","台"),ZipName,"")

						Else
							BitchHL=1
						end if
					Else
						BitchHL=1
					End If 
					rsD.close
					set rsD=Nothing
					
					If BitchHL=1 Then 
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,Driver from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof Then
							If trim(rsD2("OwnerAddress"))<>"" And Not isnull(rsD2("OwnerAddress")) Then
								If trim(rsD2("OwnerZip"))<>"" then
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=Nothing
								End if
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")&"")&"","臺","台")&"",ZipName,"")
								
							End If 
						end if
						rsD2.close
						set rsD2=Nothing
					End if
				end If
			'-------------------------------------------------------------------------------------------
			elseif sys_City="基隆市XXX" then '改跟高雄一樣
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
					if ExchangeTypeFlag="N" then	'單退先抓W看有沒有做戶籍補正，沒有就抓owner
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
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
					else
						'入案直接抓owner
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
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
			'-------------------------------------------------------------------------------------------
			elseif sys_City="台中市" or sys_City="高雄縣" then
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
					set rsD2=Nothing
					
					if sys_City="台中市" then
						'If trim(GetMailMem & "")="" Then
						if trim(rs1("Owner") & "")<>"" then
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
					End if
				end If
			'-------------------------------------------------------------------------------------------
			ElseIf sys_City="嘉義市" Or sys_City="澎湖縣" Then
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
				else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1  then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN in(select sn from billbase where billno='"&trim(rs1("BillNo"))&"' and recordstateid=0) and ExchangetypeID='A') and ExchangetypeID='A'"
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
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
				end If
			ElseIf sys_City="高雄市" Or sys_City="基隆市" Or sys_City=ApconfigureCityName Or sys_City="苗栗縣" Or sys_City="保二總隊三大隊一中隊" Then
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
					If sys_City="高雄市" Or sys_City="基隆市" Or sys_City="保二總隊三大隊一中隊" Then '如果Billbase有寫以billbase為主
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
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1 then
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
					If sys_City="高雄市" Or sys_City="基隆市" Or sys_City="保二總隊三大隊一中隊" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								if sys_City="基隆市" then
									strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipNameBill=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing

									GetMailAddress=trim(rs1("OwnerZip"))&ZipNameBill&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipNameBill,"")
								else
									GetMailAddress=trim(rs1("OwnerZip"))&" "&trim(rs1("OwnerAddress"))
								end if
							End If
						End If 
					End If
				end If
			'-------------------------------------------------------------------------------------------
			elseif sys_City="台南市" Then
				if ExchangeTypeFlag="N" then
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						'單退先抓W看有沒有做戶籍補正，沒有的話再抓A,再沒有就抓owner
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
							'if sys_City="宜蘭縣" then
							'	ZipName=""
							'else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							'end if
							
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")
						else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
							set rsD=conn.execute(strSqlD)
							if not rsD.eof then
								if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
								else
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
								end if
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
							rsD.close
							set rsD=nothing
						end if

					end if
					rsD2.close
					set rsD2=nothing
					If sys_City="台南市" Then '如果Billbase有寫以billbase為主
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
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1 then
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
							" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S' " &_
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
					If sys_City="台南市" Then '如果Billbase有寫以billbase為主
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
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
						End If 
					End If
				end If
				
			'-------------------------------------------------------------------------------------------
			elseif sys_City="南投縣" then
				if ExchangeTypeFlag="N" Then
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
							set rsZip=Nothing
						If trim(rs1("BillTypeID"))="2" Then
							GetMailMem=trim(rsD("Owner"))
						Else
							GetMailMem=trim(rsD("Driver"))
						End If 
						If Not IsNull(rsD("DriverHomeAddress")) then
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
						End if
					end if 
					rsD.close

					If ifnull(GetMailAddress) Then
						strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn " &_
						" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S'" &_
						" and Carno in (select carno from dcilog where BillSN="&trim(rs1("BillSN")) &_
						" and ExchangetypeID='A' and dcireturnstatusid='S')"
						set rsD=conn.execute(strSqlD)
						if not rsD.eof then
							if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress"))  then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
								If trim(rs1("BillTypeID"))="2" Then
									GetMailMem=trim(rsD("Owner"))
								Else
									GetMailMem=trim(rsD("Driver"))
								End If 
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
									If trim(rs1("BillTypeID"))="2" Then
										GetMailMem=trim(rsD2("Owner"))
									Else
										GetMailMem=trim(rsD2("Driver"))
									End If 
									GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress")),"臺","台"),ZipName,"")
								else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									If trim(rs1("BillTypeID"))="2" Then
										GetMailMem=trim(rsD2("Owner"))
									Else
										GetMailMem=trim(rsD2("Driver"))
									End If 
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
						set rsD2=nothing
						'如果Billbase有寫以billbase為主
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
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If

				end If
			'-------------------------------------------------------------------------------------------
			elseif sys_City="嘉義縣" or sys_City="屏東縣" Then
				if ExchangeTypeFlag="N" then
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						ZipName=""

						if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
						end if
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					end if
					rsD.close
					set rsD=Nothing
				Else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1 then
							ZipName=""			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn " &_
							" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S' " &_
							" and Carno in (select carno from dcilog where BillSN="&trim(rs1("BillSN")) &_
							" and ExchangetypeID='A' and dcireturnstatusid='S')"
							Set rsD3=conn.execute(strSqlD)
							If Not rsD3.eof Then
								If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")&"(戶)"
								Else
									ZipName=""
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
							Else
								ZipName=""
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
							rsD3.close
							Set rsD3=Nothing 
						End if
					end if
					rsD2.close
					set rsD2=Nothing
					If sys_City="屏東縣" Then '如果Billbase有寫以billbase為主
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
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
						End If 
					End If

				End If 
			'-------------------------------------------------------------------------------------------
			elseif sys_City<>"彰化縣" and sys_City<>"澎湖縣" and sys_City<>"台南市" and sys_City<>"台南縣" and sys_City<>"宜蘭縣" then	'彰化澎湖單退要抓戶籍地址
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
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
					if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
						GetMailMem="&nbsp;"
					else
						GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
					end if
					GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
				end if
				rsD.close
				set rsD=Nothing
			'-------------------------------------------------------------------------------------------
			else
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
					if ExchangeTypeFlag="N" then	'單退先抓W看有沒有做戶籍補正，沒有的話再抓A,再沒有就抓owner
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
							'if sys_City="宜蘭縣" then
							'	ZipName=""
							'else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							'end if
							
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")
						else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
							set rsD=conn.execute(strSqlD)
							if not rsD.eof then
								if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
								else
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
								end if
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
							rsD.close
							set rsD=nothing
						end if
					else
						'入案直接抓owner
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
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
				set rsD2=nothing
			end If
		'=============================================================================================
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
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
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
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
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
				set rsD=Nothing
			'-------------------------------------------------------------------------------------------
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
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
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
						elseif sys_City="宜蘭縣" or sys_City="澎湖縣" or sys_City="南投縣" or sys_City="台東縣" or sys_City="花蓮縣" then
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
						end If
							If Not IsNull(rsD("OwnerAddress")) Then 
								GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
							End If 
					end if
				end if
				rsD.close
				set rsD=nothing
			end if
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
			if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
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
elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="台中市" or sys_City="屏東縣" then 
	ReportCount=1
else
	ReportCount=2
end if
if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
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
	if sys_City="南投縣" or sys_City="基隆市" or sys_City="台東縣" or sys_City="台中市" or sys_City="高雄市"  Then
			MailBatchNumber=""
			strBatch="select distinct(a.Batchnumber) from DciLog a,BillBase f where a.BillSN=f.SN "&_
				" and f.RecordStateID=0 "&strwhere
			set rsBatch=conn.execute(strBatch)
			If Not rsBatch.Bof Then rsBatch.MoveFirst 
			While Not rsBatch.Eof
				if sys_City="南投縣" Then
					If MailBatchNumber="" Then
						MailBatchNumber=trim(rsBatch("BatchNumber"))
					End If 
				Else 
					If MailBatchNumber="" Then
						MailBatchNumber=trim(rsBatch("BatchNumber"))
					Else
						MailBatchNumber=MailBatchNumber&","&trim(rsBatch("BatchNumber"))
					End If 
				End If 
				rsBatch.MoveNext
			Wend
			rsBatch.close
			set rsBatch=Nothing
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
		if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
			response.write "宜蘭縣政府交通處"
		Else
			response.write UnitName
		End If 
		
		%>
		</span></td>
	</tr>

	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%
		if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
			response.write "26060 宜蘭市縣政北路1號"
		Else
			response.write UnitAddress
		End If 
		
		%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%
		if sys_City="宜蘭縣" and (trim(Session("Ch_Name"))="楊玉燕" or trim(Session("Ch_Name"))="許雅琪") then 
			response.write "03-9251000"
		Else
			response.write UnitTel
		End If 
		%></u>
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
	if sys_City="南投縣"  or sys_City="基隆市" or sys_City="台東縣" or sys_City="台中市" or sys_City="高雄市"  then
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
window.print();
//printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
</html>
