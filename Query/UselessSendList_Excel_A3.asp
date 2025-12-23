<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%
	StationArrayTemp=""
	strwhere=request("SQLstr")

%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style4 {
	color: #FF0000;
	font-size: 11px
}
<!--
.style5 {
	font-size: 11px
}
-->
</style>
<title>無效清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<%
strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
Sys_UnitName=rsunit("UnitName")
rsunit.close
%>
<form name=myForm method="post">
<%	
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	EofFlag1=0
	EofFlag2=0
if sys_City<>"基隆市" then 
	strSQL="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((d.DCIreturnStatus=-1 or (e.DciErrorCarData in ('1','3','9','a','j','A','H','K','T'))) and UseTool<>8) or (d.DCIreturnStatus=-1 and UseTool=8))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
else
	strSQL="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((d.DCIreturnStatus=-1 or (e.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T')) or (e.DciErrorCarData='F' and (e.rule4<>'2607' or e.rule4 is null))) and UseTool<>8) or (d.DCIreturnStatus=-1 and UseTool=8))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
end if
	set rs1=conn.execute(strSQL)
	if trim(rs1("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rs1("cnt"))/40+0.9999999)
	end if
	rs1.close
	PrintSN=0
if sys_City<>"基隆市" then 
	strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation" &_
		",d.StatusContent,a.ExchangeTypeID,a.DCIReturnStatusID,a.FileName,e.DciErrorCarData from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((d.DCIreturnStatus=-1 or (e.DciErrorCarData in ('1','3','9','a','j','A','H','K','T'))) and UseTool<>8) or (d.DCIreturnStatus=-1 and UseTool=8))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
else
	strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation" &_
		",d.StatusContent,a.ExchangeTypeID,a.DCIReturnStatusID,a.FileName,e.DciErrorCarData from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((d.DCIreturnStatus=-1 or (e.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T')) or (e.DciErrorCarData='F' and (e.rule4<>'2607' or e.rule4 is null))) and UseTool<>8) or (d.DCIreturnStatus=-1 and UseTool=8))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
end if
	set rs1=conn.execute(strSQL)
	if rs1.Eof then 
		EofFlag1=1
	end if
	While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%>無效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：逕舉</td>
			<td align="right">Page <%=fix(PrintSN/40)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="left">編號</td>
			<td width="9%" align="left">單號<br>DCI檔名</td>
			<td width="9%" align="left">違規日期<br>違規時間</td>
			<td width="9%" align="left"><br>車號</td>
			<td width="8%" align="left">法條1<br>法條2</td>
			<td width="31%" align="left"><br>駕駛人 / 車主</td>
			<td width="10%" align="left">員警<br>舉發單位</td>
			<td width="11%" align="left">無效原因<br>代保管物</td>
			<td width="9%" align="left">到案處所</td>
		</tr>
<%		
		for i=1 to 40
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>		<tr>
			<td><%
			'序號編號
			response.write PrintSN
			%></td>
			<td><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'DCI檔名
			if trim(rs1("FileName"))<>"" and not isnull(rs1("FileName")) then
				response.write "<font size=1>"&trim(rs1("FileName"))&"</font>"
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'違規日期違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write year(rs1("IllegalDate"))-1911&Right("00"&month(rs1("IllegalDate")),2)&Right("00"&day(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'車號,簡示車種
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車" 
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			%></td>
			<td><%
			'法條
			RuleStr=""
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))&"<br>"
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule2"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule2"))
				end if
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule3"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule3"))
				end if
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule4"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule4"))
				end if
			end if
			if RuleStr="" then
				response.write "&nbsp;"
			else
				response.write RuleStr
			end if
			%></td>
			<td><%
			'抓取BillBaseDCIReturn的資料
			DciOwner=""
			DciOwnerAddress=""
			DciDriverHomeAddress=""
			DCIStation=""
			if trim(rs1("BillNO"))="" or isnull(rs1("BillNO")) then
				strBillDci="select * from BillBaseDCIReturn" &_
					" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
					" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
					" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
			else
				strBillDci="select * from BillBaseDCIReturn" &_
					" where BillNO='"&trim(rs1("BillNO"))&"'" &_
					" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
					" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
					" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
			end if
			set rsBDci=conn.execute(strBillDci)
			if not rsBDci.eof then
				DciOwner=trim(rsBDci("Owner"))
				DciOwnerAddress=trim(rsBDci("OwnerAddress"))
				DciDriverHomeAddress=trim(rsBDci("DriverHomeAddress"))
				DCIStation=trim(rsBDci("DCIreturnStation"))
			end if
			rsBDci.close
			set rsBDci=nothing
			'車主
			if trim(rs1("BillTypeID"))="2" then
				response.write funcCheckFont(DciOwner,14,1)
			else
				response.write funcCheckFont(trim(rs1("Driver")),14,1)
			end if
			GetMailAddress=""
			if trim(rs1("BillTypeID"))="2" then
				if DciOwnerAddress<>"" and not isnull(DciOwnerAddress) then
					GetMailAddress=DciOwnerAddress
				end if
			else
				if DciDriverHomeAddress<>"" and not isnull(DciDriverHomeAddress) then
					GetMailAddress=DciDriverHomeAddress
				end if
			end if
			response.write "<br>"
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,14,1)&"</font>"
			%></td>
			<td><%
			'員警
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUN=conn.execute(strUName)
				if not rsUN.eof then
					response.write "<font size=1>"&trim(rsUN("UnitName"))&"</font>"
				end if
				rsUN.close
				set rsUN=nothing
			end if
			%></td>
			<td><%
			'無效原因
			if trim(rs1("DCIReturnStatusID"))<>"" and not isnull(rs1("DCIReturnStatusID")) then
				if trim(rs1("DCIReturnStatusID"))="Y" then
					response.write "<span class=""style4"">"&trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))&"</span>"
				else
					response.write trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))
				end if
				if instr("1359ajAFHKLTV",trim(rs1("DciErrorCarData")))<>-1 then
					CarErr=""
					strCarErr="select StatusContent from DciReturnStatus where DciActionID='WE'" &_
						" and DCIReturn='"&trim(rs1("DciErrorCarData"))&"'"
					set rsCarErr=conn.execute(strCarErr)
					if not rsCarErr.eof then
						CarErr=trim(rsCarErr("StatusContent"))
					end if
					rsCarErr.close
					set rsCarErr=nothing
					response.write "<br>"
					response.write "<span class=""style5"">"&trim(rs1("DciErrorCarData"))&" "&CarErr&"</span>"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'代保管物
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "0"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td><%
			'到案處所
			'攔停用BillBase=MemberStation 逕舉用BillBaseDCIReturn=DCIreturnStation
			if trim(rs1("BillTypeID"))="2" then 
				stationID=DCIStation
			else
				stationID=rs1("MemberStation")
			end if
			if trim(stationID)<>"" and not isnull(stationID) then
				strMemberStation="select DCIStationName from Station where DCIstationID='"&trim(stationID)&"'"
				set rsMS=conn.execute(strMemberStation)
				if not rsMS.eof then
					response.write "<font size=1>"&trim(rsMS("DCIStationName"))&"</font>"
				end if
				rsMS.close
				set rsMS=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>
<%
		rs1.MoveNext
		next
%>	
	</table>	
<%
	
	Wend
	rs1.close
	set rs1=nothing
if trim(PrintSN)<>"0" then
%>
共計： <%=PrintSN%>  &nbsp;筆<br>
<div class="PageNext"></div>
<%
end if
	strSQL="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID<>'2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (d.DCIreturnStatus=-1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	if trim(rs1("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rs1("cnt"))/40+0.9999999)
	end if
	rs1.close
	PrintSN=0
	strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation,e.DciErrorCarData" &_
		",d.StatusContent,a.ExchangeTypeID,a.DCIReturnStatusID,a.FileName from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID<>'2'" &_
		" and f.RecordStateID=0 and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (d.DCIreturnStatus=-1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	if rs1.Eof then 
		EofFlag2=1
	end if
	While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%>無效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：攔停</td>
			<td align="right">Page <%=fix(PrintSN/40)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="left">編號</td>
			<td width="9%" align="left">單號<br>DCI檔名</td>
			<td width="9%" align="left">違規日期<br>違規時間</td>
			<td width="9%" align="left"><br>車號</td>
			<td width="8%" align="left">法條1<br>法條2</td>
			<td width="31%" align="left"><br>駕駛人 / 車主</td>
			<td width="10%" align="left">員警<br>舉發單位</td>
			<td width="11%" align="left">無效原因<br>代保管物</td>
			<td width="9%" align="left">到案處所</td>
		</tr>
<%		
		for i=1 to 40
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>		<tr>
			<td><%
			'序號編號
			response.write PrintSN
			%></td>
			<td><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'DCI檔名
			if trim(rs1("FileName"))<>"" and not isnull(rs1("FileName")) then
				response.write "<font size=1>"&trim(rs1("FileName"))&"</font>"
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'違規日期違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write year(rs1("IllegalDate"))-1911&Right("00"&month(rs1("IllegalDate")),2)&Right("00"&day(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'車號,簡示車種
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車" 
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			%></td>
			<td><%
			'法條
			RuleStr=""
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))&"<br>"
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule2"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule2"))
				end if
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule3"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule3"))
				end if
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule4"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule4"))
				end if
			end if
			if RuleStr="" then
				response.write "&nbsp;"
			else
				response.write RuleStr
			end if
			%></td>
			<td><%
			'抓取BillBaseDCIReturn的資料
			DciOwner=""
			DciOwnerAddress=""
			DciDriverHomeAddress=""
			DCIStation=""
			if trim(rs1("BillNO"))="" or isnull(rs1("BillNO")) then
				strBillDci="select * from BillBaseDCIReturn" &_
					" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
					" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
					" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
			else
				strBillDci="select * from BillBaseDCIReturn" &_
					" where BillNO='"&trim(rs1("BillNO"))&"'" &_
					" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
					" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
					" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
			end if
			set rsBDci=conn.execute(strBillDci)
			if not rsBDci.eof then
				DciDriver=trim(rsBDci("Driver"))
				DciOwner=trim(rsBDci("Owner"))
				DciOwnerAddress=trim(rsBDci("OwnerAddress"))
				DciDriverHomeAddress=trim(rsBDci("DriverHomeAddress"))
				DCIStation=trim(rsBDci("DCIreturnStation"))
			end if
			rsBDci.close
			set rsBDci=nothing
			'車主
			if trim(rs1("BillTypeID"))="2" then
				response.write funcCheckFont(DciOwner,14,1)
			else
				response.write funcCheckFont(DciDriver,14,1)
			end if
			GetMailAddress=""
			if trim(rs1("BillTypeID"))="2" then
				if DciOwnerAddress<>"" and not isnull(DciOwnerAddress) then
					GetMailAddress=DciOwnerAddress
				end if
			else
				if DciDriverHomeAddress<>"" and not isnull(DciDriverHomeAddress) then
					GetMailAddress=DciDriverHomeAddress
				end if
			end if
			response.write "<br>"
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,14,1)&"</font>"
			%></td>
			<td><%
			'員警
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUN=conn.execute(strUName)
				if not rsUN.eof then
					response.write "<font size=1>"&trim(rsUN("UnitName"))&"</font>"
				end if
				rsUN.close
				set rsUN=nothing
			end if
			%></td>
			<td><%
			'無效原因
			if trim(rs1("DCIReturnStatusID"))<>"" and not isnull(rs1("DCIReturnStatusID")) then
				if trim(rs1("DCIReturnStatusID"))="Y" then
					response.write "<span class=""style4"">"&trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))&"</span>"
				else
					response.write trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))
				end if
				if instr("1359ajAFHKLTV",trim(rs1("DciErrorCarData")))<>-1 then
					CarErr=""
					strCarErr="select StatusContent from DciReturnStatus where DciActionID='WE'" &_
						" and DCIReturn='"&trim(rs1("DciErrorCarData"))&"'"
					set rsCarErr=conn.execute(strCarErr)
					if not rsCarErr.eof then
						CarErr=trim(rsCarErr("StatusContent"))
					end if
					rsCarErr.close
					set rsCarErr=nothing
					response.write "<br>"
					response.write "<span class=""style5"">"&trim(rs1("DciErrorCarData"))&" "&CarErr&"</span>"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'代保管物
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "0"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td><%
			'到案處所
			'攔停用BillBase=MemberStation 逕舉用BillBaseDCIReturn=DCIreturnStation
			if trim(rs1("BillTypeID"))="2" then 
				stationID=DCIStation
			else
				stationID=rs1("MemberStation")
			end if
			if trim(stationID)<>"" and not isnull(stationID) then
				strMemberStation="select DCIStationName from Station where DCIstationID='"&trim(stationID)&"'"
				set rsMS=conn.execute(strMemberStation)
				if not rsMS.eof then
					response.write "<font size=1>"&trim(rsMS("DCIStationName"))&"</font>"
				end if
				rsMS.close
				set rsMS=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>

<%
		rs1.MoveNext
		next
%>	
	</table>
	
<%
	wend
	rs1.close
	set rs1=nothing
	if trim(PrintSN)<>"0" then
%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
<%end if%>
<%
if EofFlag1=1 and EofFlag2=1 then
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%>無效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：</td>
			<td align="right">Page <%=fix(PrintSN/40)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="left">編號</td>
			<td width="9%" align="left">單號<br>DCI檔名</td>
			<td width="9%" align="left">違規日期<br>違規時間</td>
			<td width="9%" align="left"><br>車號</td>
			<td width="8%" align="left">法條1<br>法條2</td>
			<td width="31%" align="left"><br>駕駛人 / 車主</td>
			<td width="10%" align="left">員警<br>舉發單位</td>
			<td width="11%" align="left">無效原因<br>代保管物</td>
			<td width="9%" align="left">到案處所</td>
		</tr>
	</table>
<%
end if
%>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>