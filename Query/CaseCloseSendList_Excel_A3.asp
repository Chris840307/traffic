<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
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
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>結案清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<%
	strCity="select * from Apconfigure where ID=35"
	set rsCity=conn.execute(strCity)
	if not rsCity.eof then
		Sys_UnitName=rsCity("Value")
	end if
	rsCity.close
	set rsCity=Nothing
	
strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
Sys_UnitName=Sys_UnitName & rsunit("UnitName")
rsunit.close
%>
<form name=myForm method="post">
<%
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
		'=======================攔停===============================
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f where a.BillSN=f.SN and f.BillTypeID<>'2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
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

		EofFlag1=0
		EofFlag2=0
		PrintSN=0
	if ExchangeTypeFlag="N" then
		strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation,a.ExchangeTypeID" &_
		",a.DciReturnStatusID,a.FileName" &_
		",d.StatusContent,a.DCIReturnStatusID from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f,BillMailHistory g where a.BillSN=f.SN and f.BillTypeID<>'2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and f.SN=g.BillSn" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation,a.ExchangeTypeID" &_
		",a.DciReturnStatusID,a.FileName" &_
		",d.StatusContent,a.DCIReturnStatusID from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f where a.BillSN=f.SN and f.BillTypeID<>'2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
		set rs1=conn.execute(strSQL)
		if rs1.Eof then 
			EofFlag1=1
		end if
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
			if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%><%
				If sys_City="花蓮縣" Then
					response.write "舉發違反道路交通管理事件通知單退件寄存已結案清冊"
				Else
					response.write "結案清冊"
				End if
				%></font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：攔停</td>
			<td align="right">Page <%=fix(PrintSN/20)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" or sys_City="花蓮縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="0" cellspacing="0">
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
<%		for i=1 to 20
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
				response.write funcCheckFont(DciOwner,15,1)
			else
				response.write DciDriver
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
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,15,1)&"</font>"
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
				response.write trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))
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
<%end if%>
<%		'==========================逕舉=========================
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f where a.BillSN=f.SN and f.BillTypeID='2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
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

		PrintSN=0
	if ExchangeTypeFlag="N" then
		strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation,a.ExchangeTypeID" &_
		",a.DciReturnStatusID,a.FileName" &_
		",d.StatusContent,a.DCIReturnStatusID from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f,BillMailHistory g where a.BillSN=f.SN and f.BillTypeID='2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and f.SN=g.BillSN" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2,f.Rule3" &_
		",f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation,a.ExchangeTypeID" &_
		",a.DciReturnStatusID,a.FileName" &_
		",d.StatusContent,a.DCIReturnStatusID from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBase f where a.BillSN=f.SN and f.BillTypeID='2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and ((a.ExchangeTypeID='E' and a.DCIReturnStatusID='n')" &_
		" or (a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e'))" &_
		" or (a.ExchangeTypeID='N' and a.DCIReturnStatusID='n'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
		set rs1=conn.execute(strSQL)
		if rs1.Eof then 
			EofFlag2=1
		end if
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
			if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%><%
				If sys_City="花蓮縣" Then
					response.write "舉發違反道路交通管理事件通知單退件寄存已結案清冊"
				Else
					response.write "結案清冊"
				End if
				%></font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：逕舉</td>
			<td align="right">Page <%=fix(PrintSN/20)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" or sys_City="花蓮縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="0" cellspacing="0">
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
<%		for i=1 to 20
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
				response.write funcCheckFont(DciOwner,15,1)
			else
				response.write trim(rs1("Driver"))
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
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,15,1)&"</font>"
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
				response.write trim(rs1("DCIReturnStatusID"))&trim(rs1("StatusContent"))
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
<%end if%>
<%
if EofFlag1=1 and EofFlag2=1 then
%>
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%><%
				If sys_City="花蓮縣" Then
					response.write "舉發違反道路交通管理事件通知單退件寄存已結案清冊"
				Else
					response.write "結案清冊"
				End if
				%></font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：</td>
			<td align="right">Page <%=fix(PrintSN/20)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="0" cellspacing="0">
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