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
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>郵費單</title>
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

.style1 {
	font-size: 19pt;
	line-height:22px;
	font-family: "標楷體";
}
.style2 {
	font-size: 12pt;
	font-family: "標楷體";
}
.style3 {
	font-size: 10pt;
	line-height:13px;
	font-family: "標楷體";}
.style4 {
	font-size: 11pt;
	line-height:14px;
	font-family: "標楷體";}
.style5 {
	font-size: 9pt;
	line-height:11px;
	font-family: "標楷體";
}
.style6 {
	font-size: 10px;
	font-family: "標楷體";
}
.style7 {font-size: 9pt; font-family: "標楷體"; }

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


  '取得送達監理所
	'逕舉的到案處所用BillBaseDCIReturn
	ReportStationArrayTemp=""
if sys_City="基隆市" then 

	strStReport="select distinct(e.DCIReturnStation) from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and ((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (f.UseTool=8))" &strwhere&" order by e.DCIReturnStation"

elseif sys_City="台南市" then 
	strStReport="select distinct(e.DCIReturnStation) from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &strwhere&" order by e.DCIReturnStation"
else
	strStReport="select distinct(e.DCIReturnStation) from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &strwhere&" order by e.DCIReturnStation"
end if
	set rsStReport=conn.execute(strStReport)
	If Not rsStReport.Bof Then rsStReport.MoveFirst 
	While Not rsStReport.Eof
		if ReportStationArrayTemp="" then
			ReportStationArrayTemp=trim(rsStReport("DCIReturnStation"))
		else
			ReportStationArrayTemp=ReportStationArrayTemp&","&trim(rsStReport("DCIReturnStation"))
		end if
	rsStReport.MoveNext
	Wend
	rsStReport.close
	set rsStReport=nothing

	'攔停的到案處所用MemberStation
	StopStationArrayTemp=""
	strStStop="select distinct(f.MemberStation) from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.SN and a.BillTypeID<>'2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		
		" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L') and f.RecordStateID=0"&strwhere&" order by f.MemberStation"
	set rsStStop=conn.execute(strStStop)
	If Not rsStStop.Bof Then rsStStop.MoveFirst 
	While Not rsStStop.Eof
		if StopStationArrayTemp="" then
			StopStationArrayTemp=trim(rsStStop("MemberStation"))
		else
			StopStationArrayTemp=StopStationArrayTemp&","&trim(rsStStop("MemberStation"))
		end if
	rsStStop.MoveNext
	Wend
	rsStStop.close
	set rsStStop=nothing

	StationArrayTemp=""
	ReportStationArray=split(ReportStationArrayTemp,",")
	StopStationArray=split(StopStationArrayTemp,",")
	for RSA=0 to ubound(ReportStationArray)
		if instr(StationArrayTemp,ReportStationArray(RSA))=0 then
			if (ReportStationArray(RSA)="20" or ReportStationArray(RSA)="21" or ReportStationArray(RSA)="22" or ReportStationArray(RSA)="23" or ReportStationArray(RSA)="24" or ReportStationArray(RSA)="29") and instr(StationArrayTemp,"22")=0 then
				if StationArrayTemp="" then
					StationArrayTemp="22"
				else
					StationArrayTemp=StationArrayTemp&",22"
				end if
			elseif (ReportStationArray(RSA)="20" or ReportStationArray(RSA)="21" or ReportStationArray(RSA)="22" or ReportStationArray(RSA)="23" or ReportStationArray(RSA)="24" or ReportStationArray(RSA)="29") and instr(StationArrayTemp,"22")>0 then
				StationArrayTemp=StationArrayTemp
			elseif (ReportStationArray(RSA)="30" or ReportStationArray(RSA)="31" or ReportStationArray(RSA)="32") and instr(StationArrayTemp,"32")=0 then
				if StationArrayTemp="" then
					StationArrayTemp="32"
				else
					StationArrayTemp=StationArrayTemp&",32"
				end if
			elseif (ReportStationArray(RSA)="30" or ReportStationArray(RSA)="31" or ReportStationArray(RSA)="32") and instr(StationArrayTemp,"32")>0 then
				StationArrayTemp=StationArrayTemp
			else
				if StationArrayTemp="" then
					StationArrayTemp=ReportStationArray(RSA)
				else
					StationArrayTemp=StationArrayTemp&","&ReportStationArray(RSA)
				end if
			end if
		end if
	next
	for SSA=0 to ubound(StopStationArray)
		if instr(StationArrayTemp,StopStationArray(SSA))=0 then
			if (StopStationArray(SSA)="20" or StopStationArray(SSA)="21" or StopStationArray(SSA)="22" or StopStationArray(SSA)="23" or StopStationArray(SSA)="24" or StopStationArray(SSA)="29") and instr(StationArrayTemp,"22")=0 then
				if StationArrayTemp="" then
					StationArrayTemp="22"
				else
					StationArrayTemp=StationArrayTemp&",22"
				end if
			elseif (StopStationArray(SSA)="20" or StopStationArray(SSA)="21" or StopStationArray(SSA)="22" or StopStationArray(SSA)="23" or StopStationArray(SSA)="24" or StopStationArray(SSA)="29") and instr(StationArrayTemp,"22")>0 then
				StationArrayTemp=StationArrayTemp
			elseif (StopStationArray(SSA)="30" or StopStationArray(SSA)="31" or StopStationArray(SSA)="32") and instr(StationArrayTemp,"32")=0 then
				if StationArrayTemp="" then
					StationArrayTemp="32"
				else
					StationArrayTemp=StationArrayTemp&",32"
				end if
			elseif (StopStationArray(SSA)="30" or StopStationArray(SSA)="31" or StopStationArray(SSA)="32") and instr(StationArrayTemp,"32")>0 then
				StationArrayTemp=StationArrayTemp
			else
				if StationArrayTemp="" then
					StationArrayTemp=StopStationArray(SSA)
				else
					StationArrayTemp=StationArrayTemp&","&StopStationArray(SSA)
				end if
			end if
		end if
	next
	
	StationCntTotal=0
	StationArray=split(StationArrayTemp,",")
'response.write StationArrayTemp&"<br>"
'response.write UBound(StationArray)+1
%>
<table width="710" align="center" border="0" cellspacing="0" cellpadding="3">
<tr>
<td align="center" colspan="3" height="45px"><span class="style1">特&nbsp;約&nbsp;郵&nbsp;件&nbsp;郵&nbsp;費&nbsp;單</span></td>
</tr>
<tr>
<td align="left" width="45%"><span class="style2">寄件人名稱：<%
	'寄件人
	if sys_City<>"宜蘭縣" then
		strSendMem="select Value from ApConfigure where ID=27"
		set rsSendMem=conn.execute(strSendMem)
		if not rsSendMem.eof then
			response.write trim(rsSendMem("Value"))
		end if
		rsSendMem.close
		set rsSendMem=nothing
	else
		strSendU="select Value from ApConfigure where ID=35"
		set rsSendU=conn.execute(strSendU)
		if not rsSendU.eof then
			response.write trim(rsSendU("Value"))
		end if
		rsSendU.close
		set rsSendU=nothing

		strSendMem="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
		set rsSendMem=conn.execute(strSendMem)
		if not rsSendMem.eof then
			response.write trim(rsSendMem("UnitName"))
		end if
		rsSendMem.close
		set rsSendMem=nothing
	end if
%></span></td>
<td align="left" width="40%"><span class="style2">交寄日期： <%
if sys_City<>"雲林縣" and sys_City<>"宜蘭縣" then
	response.write year(now)-1911
else
	response.write "&nbsp;"
end if
%> 年 <%
if sys_City<>"雲林縣" and sys_City<>"宜蘭縣" then
	response.write month(now)
else
	response.write "&nbsp;"
end if
%> 月  <%
if sys_City<>"雲林縣" and sys_City<>"宜蘭縣" then
	response.write day(now)
else
	response.write "&nbsp;"
end if
%> 日</span></td>
<td align="left" width="15%"><span class="style2">第&nbsp; &nbsp; &nbsp; &nbsp;號</span></td>
</tr>
</table>

<table width="710" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td rowspan="2" width="3%" align="center">&nbsp;</td>
<td rowspan="2" colspan="3" width="18%" align="center"><span class="style3">寄&nbsp; &nbsp;達<br><br>地&nbsp; &nbsp;區</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">郵件類別</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">航空</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">水陸路</span></td>
<td rowspan="2" width="6%" align="center"><span class="style3">件數</span></td>
<td rowspan="2" width="6%" align="center"><span class="style3">每&nbsp;件<br><br>重&nbsp;量</span></td>
<td colspan="2" align="center"><span class="style3">每件資費</span></td>
<td rowspan="2" width="11%" align="center"><span class="style3">郵&nbsp; 費<br><br>總&nbsp; 額</span></td>
<td colspan="3" align="center"><span class="style3">※收寄單位複核</span></td>
<td rowspan="2" width="11%" align="center"><span class="style3">備註</span></td>
</tr>
<tr>
<td width="6%" align="center"><span class="style3">郵費</span></td>
<td width="7%" align="center"><span class="style3">存證費</span></td>
<td width="8%" align="center"><span class="style3">件數</span></td>
<td colspan="2" align="center"><span class="style3">郵費總額</span></td>
</tr>
<%
if (UBound(StationArray)+1) >= 16 then
	ArrayCount=15
else
	ArrayCount=UBound(StationArray)
	SpaceCol=16-(UBound(StationArray)+1)
end if
MailTotalCnt=0
for i=0  to ArrayCount

%>
<tr>
<td height="20" align="center"><%=i+1%></td>
<td colspan="3"><span class="style3"><%
	'監理站名稱
	strSqlStationName1="select DCIstationName from Station where DCIstationID='"&trim(StationArray(i))&"'"
	set rsSN1=conn.execute(strSqlStationName1)
	if not rsSN1.eof then
		response.Write trim(rsSN1("DCIstationName"))
	end if
	rsSN1.close
	set rsSN1=nothing
%></span></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td align="center"><%
	'件數
	MailCount1=0
	'逕舉的到案處所用BillBaseDCIReturn
	if trim(StationArray(i))="22" then
		StationID="'20','21','22','23','24','29'"
	elseif trim(StationArray(i))="32" then
		StationID="'30','31','32'"
	else
		StationID="'"&trim(StationArray(i))&"'"
	end if
if sys_City="基隆市" then 

	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and ((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
elseif sys_City="台南市" then 
	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
else
	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
end if
	set rsCntReport=conn.execute(strCntReport)
	if not rsCntReport.eof then
		MailCount1=MailCount1+cint(rsCntReport("cnt"))
		MailCountTotal=MailCountTotal+cint(rsCntReport("cnt"))
	end if
	rsCntReport.close
	set rsCntReport=nothing
	
	'攔停的到案處所用MemberStation
	strCntStop="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.SN and a.BillTypeID<>'2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		
		" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ("&StationID&") and f.RecordStateID=0"&strwhere
	set rsCntStop=conn.execute(strCntStop)
	if not rsCntStop.eof then
		MailCount1=MailCount1+cint(rsCntStop("cnt"))
		MailCountTotal=MailCountTotal+cint(rsCntStop("cnt"))
	end if
	rsCntStop.close
	set rsCntStop=nothing
	response.write MailCount1
	MailTotalCnt=MailTotalCnt+MailCount1
%></td>
<td>&nbsp;</td>
<td align="center"><%
	'郵費
	if trim(request("MailMoneyType"))="1" then
		MailMoney=25
	elseif trim(request("MailMoneyType"))="2" then
		MailMoney=24
	elseif trim(request("MailMoneyType"))="3" then
		MailMoney=trim(request("MailMoneyValue"))
	elseif trim(request("MailMoneyType"))="4" then
		MailMoney=""
	end if
	if MailMoney<>"" then
		response.write MailMoney
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td align="center"><%
	'郵費總額
	if trim(request("MailMoneyType"))="1" then
		MailMoney=25
	elseif trim(request("MailMoneyType"))="2" then
		MailMoney=24
	elseif trim(request("MailMoneyType"))="3" then
		MailMoney=trim(request("MailMoneyValue"))
	elseif trim(request("MailMoneyType"))="4" then
		MailMoney=""
	end if
	if MailMoney<>"" then
		response.write MailMoney*MailCount1
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td width="10%">&nbsp;</td>
<td width="3%">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<%next
for i2=1 to SpaceCol
%>
<tr>
<td height="20" align="center"><%=i+i2%></td>
<td colspan="3">&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td width="10%">&nbsp;</td>
<td width="3%">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<%
next
%>
<tr>
<td colspan="4" height="20"><span class="style3">總 &nbsp;計</span></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td align="center"><%=MailTotalCnt%></td>
<td>&nbsp;</td>
<td align="center"><%
'總郵費
	if MailMoney<>"" then
		response.write MailMoney
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td align="center"><%
'總郵費總額
	if MailMoney<>"" then
		response.write MailMoney*MailTotalCnt
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td colspan="12" height="105">
<div>
<table width="350" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td align="center" width="34%"><span class="style3">應&nbsp;付&nbsp;郵&nbsp;資</span></td>
<td align="center" width="33%"><span class="style3">折&nbsp;扣&nbsp;郵&nbsp;資<br>基本折扣&nbsp; &nbsp; %<br>附加折扣 &nbsp; &nbsp;%</span></td>
<td align="center" width="33%"><span class="style3">實&nbsp;付&nbsp;郵&nbsp;資</span></td>
</tr>
<tr>
<td align="right" height="40"><span class="style3"><%
'總郵費總額
	if MailMoney<>"" then
		response.write MailTotalCnt*MailMoney
	else
		response.write "&nbsp;"
	end if
%>&nbsp; 元</span></td>
<td align="right"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
</tr>
</table>
</div>
</td>
<td colspan="4">
<div>
<table width="220" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td align="center" width="30%"><span class="style3">應付郵資</span></td>
<td align="center" width="40%"><span class="style3">折扣郵資<br>基本折扣&nbsp; %<br>附加折扣&nbsp; %</span></td>
<td align="center" width="30%"><span class="style3">實付郵資</span></td>
</tr>
<tr>
<td align="right" height="40"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
</tr>
</table>
</div>
</td>
</tr>
<tr>
<td colspan="2" height="55" align="center">
<span class="style3">實&nbsp;收<br>郵&nbsp;費<br>共&nbsp;計<br>(大寫)</span>
</td>
<td colspan="9" height="55" align="right" valign="top"><span class="style4"><br><br>
萬&nbsp; &nbsp; 仟&nbsp; &nbsp; 佰&nbsp; &nbsp; 拾&nbsp; &nbsp; 元&nbsp; &nbsp; 角整</span>
</td>
<td colspan="5" align="right" valign="top"><span class="style4">
<div align="left">※</div><br>
萬&nbsp; &nbsp; 仟&nbsp; &nbsp; 佰&nbsp; &nbsp; 拾&nbsp; &nbsp; 元&nbsp; &nbsp; 角整</span>
</td>
</tr>
<tr>
<td colspan="3" height="55" valign="top">
<span class="style4">寄件人簽章：</span>
</td>
<td colspan="8" height="105">
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="50%" align="left" valign="top"><span class="style4">
※<br>&nbsp; 收寄單位：<br><br>&nbsp; 經&nbsp;辦&nbsp;員：<br><br>&nbsp; 主&nbsp;&nbsp; &nbsp;管：</span>
</td>
<td align="right" width="50%"><img src="../Image/MailMoneyPic1.jpg" width="96" height="94" /></td>
</tr>
</table>
</td>
<td colspan="5">
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="58%" align="left" valign="top"><span class="style4">
※<br>&nbsp; 郵件單位：<br><br>&nbsp; 經&nbsp;辦&nbsp;員：<br><br>&nbsp; 主&nbsp;&nbsp; &nbsp;管：</span>
</td>
<td align="right" width="42%"><img src="../Image/MailMoneyPic1.jpg" width="96" height="94" /></td>
</tr>
</table>
</td>
</tr>
</table>

<table width="710" align="center" border="0" cellspacing="0" cellpadding="3">
<tr>
<td width="7%" valign="top">
<span class="style3">
注意：
</span>
</td>
<td width="93%">
<span class="style5">
(1)※記號欄由郵局填寫，本單上數字如有更正，應由相關人員在更正處蓋章。<br>
(2)本日實寄件數及所需郵費總額以郵局複核填寫者為準。折扣郵資僅限於郵件轉運局或指定局交寄已分區捆紮郵件適用。並於日(月)報表中註明當日(月)給予折扣郵資之總金額。<br>
(3)每次應請寄件人填寫一式四份，但交寄大宗存證信函之存證費採逐月結帳應填寫一式五份，其中一份經收寄單位簽收後退還寄件人。<br>
(4)交寄大宗郵件如包括本地及外地郵件，且已辦理分區捆紮者，請寄件人填寫一式六份。
</span>
</td>
</tr>
</table>

<%
'============監理站超過16個========================================
if (UBound(StationArray)+1) > 16 then

response.write "<div class=""PageNext""></div>"
%>
<table width="710" align="center" border="0" cellspacing="0" cellpadding="3">
<tr>
<td align="center" colspan="3" height="45px"><span class="style1">特&nbsp;約&nbsp;郵&nbsp;件&nbsp;郵&nbsp;費&nbsp;單</span></td>
</tr>
<tr>
<td align="left" width="45%"><span class="style2">寄件人名稱：<%
	'寄件人
	if sys_City<>"宜蘭縣" then
		strSendMem="select Value from ApConfigure where ID=27"
		set rsSendMem=conn.execute(strSendMem)
		if not rsSendMem.eof then
			response.write trim(rsSendMem("Value"))
		end if
		rsSendMem.close
		set rsSendMem=nothing
	else
		strSendU="select Value from ApConfigure where ID=35"
		set rsSendU=conn.execute(strSendU)
		if not rsSendU.eof then
			response.write trim(rsSendU("Value"))
		end if
		rsSendU.close
		set rsSendU=nothing

		strSendMem="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
		set rsSendMem=conn.execute(strSendMem)
		if not rsSendMem.eof then
			response.write trim(rsSendMem("UnitName"))
		end if
		rsSendMem.close
		set rsSendMem=nothing
	end if
%></span></td>
<td align="left" width="40%"><span class="style2">交寄日期： <%
if sys_City<>"雲林縣" then
	response.write year(now)-1911
end if
%> 年 <%
if sys_City<>"雲林縣" then
	response.write month(now)
end if
%> 月  <%
if sys_City<>"雲林縣" then
	response.write day(now)
end if
%> 日</span></td>
<td align="left" width="15%"><span class="style2">第&nbsp; &nbsp; &nbsp; &nbsp;號</span></td>
</tr>
</table>

<table width="710" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td rowspan="2" width="3%" align="center">&nbsp;</td>
<td rowspan="2" colspan="3" width="18%" align="center"><span class="style3">寄&nbsp; &nbsp;達<br><br>地&nbsp; &nbsp;區</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">郵件類別</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">航空</span></td>
<td rowspan="2" width="3%" align="center"><span class="style3">水陸路</span></td>
<td rowspan="2" width="6%" align="center"><span class="style3">件數</span></td>
<td rowspan="2" width="6%" align="center"><span class="style3">每&nbsp;件<br><br>重&nbsp;量</span></td>
<td colspan="2" align="center"><span class="style3">每件資費</span></td>
<td rowspan="2" width="11%" align="center"><span class="style3">郵&nbsp; 費<br><br>總&nbsp; 額</span></td>
<td colspan="3" align="center"><span class="style3">※收寄單位複核</span></td>
<td rowspan="2" width="11%" align="center"><span class="style3">備註</span></td>
</tr>
<tr>
<td width="6%" align="center"><span class="style3">郵費</span></td>
<td width="7%" align="center"><span class="style3">存證費</span></td>
<td width="8%" align="center"><span class="style3">件數</span></td>
<td colspan="2" align="center"><span class="style3">郵費總額</span></td>
</tr>
<%
if (UBound(StationArray)+1-16) >= 16 then
	ArrayCount2=15
else
	ArrayCount2=UBound(StationArray)
	SpaceCol2=16-(UBound(StationArray)+1-16)
end if
MailTotalCnt=0
for i=16  to ArrayCount2

%>
<tr>
<td height="20" align="center"><%=i+1-16%></td>
<td colspan="3"><span class="style3"><%
	'監理站名稱
	strSqlStationName1="select DCIstationName from Station where DCIstationID='"&trim(StationArray(i))&"'"
	set rsSN1=conn.execute(strSqlStationName1)
	if not rsSN1.eof then
		response.Write trim(rsSN1("DCIstationName"))
	end if
	rsSN1.close
	set rsSN1=nothing
%></span></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td align="center"><%
	'件數
	MailCount1=0
	'逕舉的到案處所用BillBaseDCIReturn
	if trim(StationArray(i))="22" then
		StationID="'20','21','22','23','24','29'"
	elseif trim(StationArray(i))="32" then
		StationID="'30','31','32'"
	else
		StationID="'"&trim(StationArray(i))&"'"
	end if
if sys_City="基隆市" then 

	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and ((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
elseif sys_City="台南市" then 
	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
else
	strCntReport="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.Sn and f.RecordStateID=0" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		
		" and d.DCIreturnStatus=1 and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and f.UseTool<>8) or (f.UseTool=8))" &_
		" and e.DCIReturnStation in ("&StationID&")"&strwhere
end if
	set rsCntReport=conn.execute(strCntReport)
	if not rsCntReport.eof then
		MailCount1=MailCount1+cint(rsCntReport("cnt"))
		MailCountTotal=MailCountTotal+cint(rsCntReport("cnt"))
	end if
	rsCntReport.close
	set rsCntReport=nothing
	
	'攔停的到案處所用MemberStation
	strCntStop="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBaseDCIReturn e,BillBase f" &_
		" where a.BillSN=f.SN and a.BillTypeID<>'2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.BillNo=e.BillNO) or (a.BillNo is null and e.BillNO is null)) and a.CarNo=e.CarNo" &_
		
		" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ("&StationID&") and f.RecordStateID=0"&strwhere
	set rsCntStop=conn.execute(strCntStop)
	if not rsCntStop.eof then
		MailCount1=MailCount1+cint(rsCntStop("cnt"))
		MailCountTotal=MailCountTotal+cint(rsCntStop("cnt"))
	end if
	rsCntStop.close
	set rsCntStop=nothing
	response.write MailCount1
	MailTotalCnt=MailTotalCnt+MailCount1
%></td>
<td>&nbsp;</td>
<td align="center"><%
	'郵費
	if trim(request("MailMoneyType"))="1" then
		MailMoney=25
	elseif trim(request("MailMoneyType"))="2" then
		MailMoney=24
	elseif trim(request("MailMoneyType"))="3" then
		MailMoney=trim(request("MailMoneyValue"))
	elseif trim(request("MailMoneyType"))="4" then
		MailMoney=""
	end if
	if MailMoney<>"" then
		response.write MailMoney
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td align="center"><%
	'郵費總額
	if trim(request("MailMoneyType"))="1" then
		MailMoney=25
	elseif trim(request("MailMoneyType"))="2" then
		MailMoney=24
	elseif trim(request("MailMoneyType"))="3" then
		MailMoney=trim(request("MailMoneyValue"))
	elseif trim(request("MailMoneyType"))="4" then
		MailMoney=""
	end if
	if MailMoney<>"" then
		response.write MailMoney*MailCount1
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td width="10%">&nbsp;</td>
<td width="3%">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<%next
for i2=1 to SpaceCol2
%>
<tr>
<td height="20" align="center"><%=i+i2-16%></td>
<td colspan="3">&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td width="10%">&nbsp;</td>
<td width="3%">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<%
next
%>
<tr>
<td colspan="4" height="20"><span class="style3">總 &nbsp;計</span></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td align="center"><%=MailTotalCnt%></td>
<td>&nbsp;</td>
<td align="center"><%
'總郵費
	if MailMoney<>"" then
		response.write MailMoney
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td align="center"><%
'總郵費總額
	if MailMoney<>"" then
		response.write MailMoney*MailTotalCnt
	else
		response.write "&nbsp;"
	end if
%></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td colspan="12" height="105">
<div>
<table width="350" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td align="center" width="34%"><span class="style3">應&nbsp;付&nbsp;郵&nbsp;資</span></td>
<td align="center" width="33%"><span class="style3">折&nbsp;扣&nbsp;郵&nbsp;資<br>基本折扣&nbsp; &nbsp; %<br>附加折扣 &nbsp; &nbsp;%</span></td>
<td align="center" width="33%"><span class="style3">實&nbsp;付&nbsp;郵&nbsp;資</span></td>
</tr>
<tr>
<td align="right" height="40"><span class="style3"><%
'總郵費總額
	if MailMoney<>"" then
		response.write MailTotalCnt*MailMoney
	else
		response.write "&nbsp;"
	end if
%>&nbsp; 元</span></td>
<td align="right"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
</tr>
</table>
</div>
</td>
<td colspan="4">
<div>
<table width="220" align="center" border="1" cellspacing="0" cellpadding="3">
<tr>
<td align="center" width="30%"><span class="style3">應付郵資</span></td>
<td align="center" width="40%"><span class="style3">折扣郵資<br>基本折扣&nbsp; %<br>附加折扣&nbsp; %</span></td>
<td align="center" width="30%"><span class="style3">實付郵資</span></td>
</tr>
<tr>
<td align="right" height="40"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
<td align="right"><span class="style3">元</span></td>
</tr>
</table>
</div>
</td>
</tr>
<tr>
<td colspan="2" height="55" align="center">
<span class="style3">實&nbsp;收<br>郵&nbsp;費<br>共&nbsp;計<br>(大寫)</span>
</td>
<td colspan="9" height="55" align="right" valign="top"><span class="style4"><br><br>
萬&nbsp; &nbsp; 仟&nbsp; &nbsp; 佰&nbsp; &nbsp; 拾&nbsp; &nbsp; 元&nbsp; &nbsp; 角整</span>
</td>
<td colspan="5" align="right" valign="top"><span class="style4">
<div align="left">※</div><br>
萬&nbsp; &nbsp; 仟&nbsp; &nbsp; 佰&nbsp; &nbsp; 拾&nbsp; &nbsp; 元&nbsp; &nbsp; 角整</span>
</td>
</tr>
<tr>
<td colspan="3" height="55" valign="top">
<span class="style4">寄件人簽章：</span>
</td>
<td colspan="8" height="105">
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="50%" align="left" valign="top"><span class="style4">
※<br>&nbsp; 收寄單位：<br><br>&nbsp; 經&nbsp;辦&nbsp;員：<br><br>&nbsp; 主&nbsp;&nbsp; &nbsp;管：</span>
</td>
<td align="right" width="50%"><img src="../Image/MailMoneyPic1.jpg" width="96" height="94" /></td>
</tr>
</table>
</td>
<td colspan="5">
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="58%" align="left" valign="top"><span class="style4">
※<br>&nbsp; 郵件單位：<br><br>&nbsp; 經&nbsp;辦&nbsp;員：<br><br>&nbsp; 主&nbsp;&nbsp; &nbsp;管：</span>
</td>
<td align="right" width="42%"><img src="../Image/MailMoneyPic1.jpg" width="96" height="94" /></td>
</tr>
</table>
</td>
</tr>
</table>

<table width="710" align="center" border="0" cellspacing="0" cellpadding="3">
<tr>
<td width="7%" valign="top">
<span class="style3">
注意：
</span>
</td>
<td width="93%">
<span class="style5">
(1)※記號欄由郵局填寫，本單上數字如有更正，應由相關人員在更正處蓋章。<br>
(2)本日實寄件數及所需郵費總額以郵局複核填寫者為準。折扣郵資僅限於郵件轉運局或指定局交寄已分區捆紮郵件適用。並於日(月)報表中註明當日(月)給予折扣郵資之總金額。<br>
(3)每次應請寄件人填寫一式四份，但交寄大宗存證信函之存證費採逐月結帳應填寫一式五份，其中一份經收寄單位簽收後退還寄件人。<br>
(4)交寄大宗郵件如包括本地及外地郵件，且已辦理分區捆紮者，請寄件人填寫一式六份。
</span>
</td>
</tr>
</table>

<%end if%>
</body>

<script language="javascript">
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" or sys_City="花蓮縣" then%>
window.print();
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
</html>
