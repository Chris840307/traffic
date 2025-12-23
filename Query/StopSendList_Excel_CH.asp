<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
.style5 {font-family:新細明體; color=0044ff; line-height:13px; font-size: 11px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>攔停舉發移送清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

 'and a.BillTypeID<>'2'
%>
<%
	'頁數
	PageNum=1
	StationArrayTemp=""
	strwhere=request("SQLstr")
	strStation="select distinct(f.MemberStation) from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere&" order by MemberStation"
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then rsStation.MoveFirst 
	While Not rsStation.Eof
		if StationArrayTemp="" then
			StationArrayTemp=trim(rsStation("MemberStation"))
		else
			StationArrayTemp=StationArrayTemp&","&trim(rsStation("MemberStation"))
		end if
	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing

	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=nothing
%>
</head>
<body>
<form name=myForm method="post">

<%
	StationArray=split(StationArrayTemp,",")
%>
<%
	strUnitName2="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		TitleUnitName2=trim(rsUnitName2("UnitName"))
	end if
	rsUnitName2.close
	set rsUnitName2=nothing

	strUnitName="select Value from ApConfigure where ID=40"
	set rsUnitName=conn.execute(strUnitName)
	if not rsUnitName.eof then
		TitleUnitName=trim(rsUnitName("value"))&" "&TitleUnitName2
	end if
	rsUnitName.close
	set rsUnitName=nothing

	PrintSNtotal=0	'編號

	'台北市交通裁決所舉發單列表
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
		DciStationName="台北市交通事件裁決所"
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation in ('20','21','22','23','24','29')" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/21+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			" from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation in ('20','21','22','23','24','29')" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then
%>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
			response.write "<div class=""PageNext""></div>"
		end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><%=TitleUnitName%>&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left">站所：<%
		response.write DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"></td>
			<td width="10%">單號</td>
			<td width="9%">違規日期</td>
			<td width="9%"></td>
			<td width="7%"></td>
			<td width="18%"></td>
			<td width="16%">舉發單位</td>
			<td width="8%">員警</td>
			<td width="8%">扣件</td>
			<td width="10%">備註</td>
		</tr>
		<tr>
			<td>編號</td>
			<td>入案日期</td>
			<td>違規時間</td>
			<td>車號</td>
			<td>法條</td>
			<td>駕駛人/車主</td>
			<td>駕籍資料</td>
			<td></td>
			<td>車籍</td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 21
			if rs1.eof then exit for
%>
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"><%
			PrintSN=PrintSN+1
			PrintSNtotal=PrintSNtotal+1
			response.write PrintSNtotal
			%></td>
			<td width="10%"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write rs1("BillNO")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%response.write trim(rs1("CarNo"))%></td>
			<td width="7%"><span class="style5"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="18%"><span class="style5"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="16%"><span class="style5"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					response.write trim(rsUnit("UnitName"))
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write "<b>" & rsBF("Content") & "</b>"
				'smith add <b> tag
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="10%"><%
			'檔名
			response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			%></td>
		</tr>
		<tr>
			<td></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%response.write trim(rs1("CarSimpleID"))%></td>
			<td><span class="style5"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			'駕籍
			if trim(rs1("DCIErrorIDData"))="0" then
				response.write "0 正常"
			elseif trim(rs1("DCIErrorIDData"))<>"" and not isnull(rs1("DCIErrorIDData")) then
				strDriverData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorIDData"))&"'"
				set rsDD=conn.execute(strDriverData)
				if not rsDD.eof then
					response.write trim(rs1("DCIErrorIDData"))&" "&trim(rsDD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsDD.close
				set rsDD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td></td>
			<td><span class="style5"><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><%
			'批號
			response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
			%></td>
		</tr>
		</table>
		</td>
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

%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%if SA<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if

	end if

	'高雄市交通事件裁決所列表
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
		DciStationName="高雄市交通事件裁決所"
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation in ('30','31','32')" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/21+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			" from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation in ('30','31','32')" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then
%>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
			response.write "<div class=""PageNext""></div>"
		end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><%=TitleUnitName%>&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left">站所：<%
		response.write DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"></td>
			<td width="10%">單號</td>
			<td width="9%">違規日期</td>
			<td width="9%"></td>
			<td width="7%"></td>
			<td width="18%"></td>
			<td width="16%">舉發單位</td>
			<td width="8%">員警</td>
			<td width="8%">扣件</td>
			<td width="10%">備註</td>
		</tr>
		<tr>
			<td>編號</td>
			<td>入案日期</td>
			<td>違規時間</td>
			<td>車號</td>
			<td>法條</td>
			<td>駕駛人/車主</td>
			<td>駕籍資料</td>
			<td></td>
			<td>車籍</td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 21
			if rs1.eof then exit for
%>
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"><%
			PrintSN=PrintSN+1
			PrintSNtotal=PrintSNtotal+1
			response.write PrintSNtotal
			%></td>
			<td width="10%"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write rs1("BillNO")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%response.write trim(rs1("CarNo"))%></td>
			<td width="7%"><span class="style5"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="18%"><span class="style5"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="16%"><span class="style5"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					response.write trim(rsUnit("UnitName"))
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write "<b>" & rsBF("Content") & "</b>"
				'smith add <b> tag
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="10%"><%
			'檔名
			response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			%></td>
		</tr>
		<tr>
			<td></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%response.write trim(rs1("CarSimpleID"))%></td>
			<td><span class="style5"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			'駕籍
			if trim(rs1("DCIErrorIDData"))="0" then
				response.write "0 正常"
			elseif trim(rs1("DCIErrorIDData"))<>"" and not isnull(rs1("DCIErrorIDData")) then
				strDriverData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorIDData"))&"'"
				set rsDD=conn.execute(strDriverData)
				if not rsDD.eof then
					response.write trim(rs1("DCIErrorIDData"))&" "&trim(rsDD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsDD.close
				set rsDD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td></td>
			<td><span class="style5"><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><%
			'批號
			response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
			%></td>
		</tr>
		</table>
		</td>
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

%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%if SA<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if

	end if

	'其他監理所舉發單列表
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
	if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then
		DciStationName=""
		strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(StationArray(SA))&"'"
		set rsSN=conn.execute(strSqlStationName)
		if not rsSN.eof then
			DciStationName=trim(rsSN("DCIstationName"))
		end if
		rsSN.close
		set rsSN=nothing
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation='"&trim(StationArray(SA))&"'" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/21+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			" from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation='"&trim(StationArray(SA))&"'" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then
%>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
			response.write "<div class=""PageNext""></div>"
		end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><%=TitleUnitName%>&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left">站所：<%
		response.write trim(StationArray(SA))&" "&DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"></td>
			<td width="10%">單號</td>
			<td width="9%">違規日期</td>
			<td width="9%"></td>
			<td width="7%"></td>
			<td width="18%"></td>
			<td width="16%">舉發單位</td>
			<td width="8%">員警</td>
			<td width="8%">扣件</td>
			<td width="10%">備註</td>
		</tr>
		<tr>
			<td>編號</td>
			<td>入案日期</td>
			<td>違規時間</td>
			<td>車號</td>
			<td>法條</td>
			<td>駕駛人/車主</td>
			<td>駕籍資料</td>
			<td></td>
			<td>車籍</td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 21
			if rs1.eof then exit for
%>
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"><%
			PrintSN=PrintSN+1
			PrintSNtotal=PrintSNtotal+1
			response.write PrintSNtotal
			%></td>
			<td width="10%"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write rs1("BillNO")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%response.write trim(rs1("CarNo"))%></td>
			<td width="7%"><span class="style5"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="18%"><span class="style5"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="16%"><span class="style5"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					response.write trim(rsUnit("UnitName"))
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="8%"><span class="style5"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write "<b>" & rsBF("Content") & "</b>"
				'smith add <b> tag
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="10%"><%
			'檔名
			response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			%></td>
		</tr>
		<tr>
			<td></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%response.write trim(rs1("CarSimpleID"))%></td>
			<td><span class="style5"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style5"><%
			'駕籍
			if trim(rs1("DCIErrorIDData"))="0" then
				response.write "0 正常"
			elseif trim(rs1("DCIErrorIDData"))<>"" and not isnull(rs1("DCIErrorIDData")) then
				strDriverData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorIDData"))&"'"
				set rsDD=conn.execute(strDriverData)
				if not rsDD.eof then
					response.write trim(rs1("DCIErrorIDData"))&" "&trim(rsDD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsDD.close
				set rsDD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td></td>
			<td><span class="style5"><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><%
			'批號
			response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
			%></td>
		</tr>
		</table>
		</td>
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

%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%if SA<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	end if
	next
%>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,3.08,3.08,3.08,3.08);
</script>
<%conn.close%>