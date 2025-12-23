<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 8000
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style2 {font-family:新細明體; color=0044ff; line-height:23px; font-size: 18px}
.style3 {font-family:新細明體; color=0044ff; line-height:22px; font-size: 16px}
.style5 {font-family:新細明體; color=0044ff; line-height:15px; font-size: 11px}
.style6 {font-family:新細明體; color=0044ff; line-height:25px; font-size: 20px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>逕行舉發移送清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
'權限
'AuthorityCheck(234)

 'and a.BillTypeID<>'2'
 strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%
	'頁數
	PageNum=1
	StationArrayTemp=""
	strwhere=request("SQLstr")
	strStation="select distinct(e.DCIReturnStation) from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere&" order by DCIReturnStation"
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then rsStation.MoveFirst 
	While Not rsStation.Eof
		if StationArrayTemp="" then
			StationArrayTemp=trim(rsStation("DCIReturnStation"))
		else
			StationArrayTemp=StationArrayTemp&","&trim(rsStation("DCIReturnStation"))
		end if
	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing

	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
		" BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=nothing
'response.write strSQL
%>
</head>
<body>
<form name=myForm method="post">
<%if sys_City<>"嘉義縣" and sys_City<>"高雄市" then%>
<center><span class="style2">舉發違反道路交通事件通知單逕行舉發移送清冊</span></font></center>
	<table width="600" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="3" cellspacing="0" align="center">
		<tr>
			<td width="33%" align="center"><span class="style3">受文單位</span></td>
			<td width="33%" align="center"><span class="style3">移送件數</span></td>
			<td width="33%" align="center"><span class="style3">備考</span></td>
		</tr>
<%	StationCntTotal=0
	'台北市交通裁決所數量
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
%>
		<tr>
			<td><span class="style3"><%
			'受文單位
			response.write "台北市交通事件裁決所"
			%></span></td>
			<td align="center"><span class="style3"><%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation in ('20','21','22','23','24','29') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and e.DCIReturnStation in ('20','21','22','23','24','29') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCloseCntReport=conn.execute(strCloseCntReport)
		if not rsCloseCntReport.eof then
			if trim(rsCloseCntReport("cnt"))>0 then
				response.write "結案 "&trim(rsCloseCntReport("cnt"))&" 件"
			else
				response.write "&nbsp;"
			end if
		end if
		rsCloseCntReport.close
		set rsCloseCntReport=nothing
			%></span></td>
		</tr>
<%

	end if

	'高雄市交通事件裁決所數量
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
%>
		<tr>
			<td><span class="style3"><%
			'受文單位
			response.write "高雄市交通事件裁決所"
			%></span></td>
			<td align="center"><span class="style3"><%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation in ('30','31','32') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and e.DCIReturnStation in ('30','31','32') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCloseCntReport=conn.execute(strCloseCntReport)
		if not rsCloseCntReport.eof then
			if trim(rsCloseCntReport("cnt"))>0 then
				response.write "結案 "&trim(rsCloseCntReport("cnt"))&" 件"
			else
				response.write "&nbsp;"
			end if
		end if
		rsCloseCntReport.close
		set rsCloseCntReport=nothing
			%></span></td>
		</tr>
<%
	end if
	'其他監理所數量
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then
%>
		<tr>
			<td><span class="style3"><%
			'受文單位
		strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(StationArray(SA))&"'"
		set rsSN=conn.execute(strSqlStationName)
		if not rsSN.eof then
			response.write trim(rsSN("DCIstationName"))
		end if
		rsSN.close
		set rsSN=nothing
			%></span></td>
			<td align="center"><span class="style3"><%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCloseCntReport=conn.execute(strCloseCntReport)
		if not rsCloseCntReport.eof then
			if trim(rsCloseCntReport("cnt"))>0 then
				response.write "結案 "&trim(rsCloseCntReport("cnt"))&" 件"
			else
				response.write "&nbsp;"
			end if
		end if
		rsCloseCntReport.close
		set rsCloseCntReport=nothing
			%></span></td>
		</tr>
<%		end if
	next
%>
		<tr>
			<td><span class="style3">小計</span></td>
			<td align="center"><span class="style3"><%=StationCntTotal%></span></td>
			<td>&nbsp;</td>
		</tr>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
	<div class="PageNext"></div>
<%else
	StationCntTotal=0
	'台北市交通裁決所數量
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation in ('20','21','22','23','24','29') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
	end if

	'高雄市交通事件裁決所數量
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation in ('30','31','32') and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
	end if

	'其他監理所數量
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then

		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L') and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
		" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		StationCntTotal=StationCntTotal+StationCnt
		end if
	next
end if
%>
<%
	strUnitName2="select UnitName,UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		if sys_City="高雄市" then
			strT2="select UnitName from UnitInfo where UnitID='"&trim(rsUnitName2("UnitTypeID"))&"'"
			set rsT2=conn.execute(strT2)
			if not rsT2.eof then
				TitleUnitName2=trim(rsT2("UnitName"))
			end if
			rsT2.close
			set rsT2=nothing
		else
			TitleUnitName2=trim(rsUnitName2("UnitName"))
		end if
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
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation in ('20','21','22','23','24','29') and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/16+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID,f.EquipMentID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation in ('20','21','22','23','24','29') and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
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
	<table width="100%" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><span class="style2"><%=TitleUnitName%>&nbsp;逕行舉發移送清冊</span></td>
		</tr>
		<tr>
			<td align="left"><span class="style3">站所：<%
		response.write "<strong><font class=""style6"">"&"台北市交通事件裁決所"&"</font></strong>"
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/16)+1%> of <%=pagecnt%></span></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"></td>
			<td width="10%"><span class="style3">單號</span></td>
			<td width="9%"><span class="style3"><!-- 違規日期 --></span></td>
			<td width="9%"></td>
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="18%"><span class="style3">舉發單位</span></td>
			<td width="9%"><span class="style3">員警</span></td>
			<td width="10%"><span class="style3">扣件</span></td>
			<td width="9%"><span class="style3"><!-- 備註 --></span></td>
		</tr>
		<tr>
			<td><span class="style3"><!-- 編號 --></span></td>
			<td><span class="style3">入案日期</span></td>
			<td><span class="style3">違規日期<!-- 違規時間 --></span></td>
			<td><span class="style3">車號</span></td>
			<td><span class="style3">法條</span></td>
			<td><span class="style3">駕駛人/車主</span></td>
			<td><span class="style3">駕籍資料</span></td>
			<td></td>
			<td><span class="style3">車籍資料</span></td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 16
			if rs1.eof then exit for
			Response.flush
%>
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"><span class="style3"><%
			PrintSNtotal=PrintSNtotal+1
			PrintSN=PrintSN+1
			'response.write PrintSNtotal
			%></span></td>
			<td width="10%"><span class="style3"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write rs1("BillNO")
				else
					response.write "<strong>"&rs1("BillNO")&"</strong>"
				end if
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="9%"><span class="style3"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="9%"><span class="style3"><%response.write trim(rs1("CarNo"))%></span></td>
			<td width="8%"><span class="style3"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="18%"><span class="style3"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),15,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="18%"><span class="style3"><%
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
			<td width="9%"><span class="style3"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="10%"><span class="style3"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="9%"><span class="style3"><%
			'檔名
			response.write "&nbsp;"
			%></span></td>
		</tr>
		<tr>
			<td></td>
			<td><span class="style3"><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
			'if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			'	response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			'else
				response.write "&nbsp;"
			'end if
			%></span></td>
			<td><span class="style3"><%response.write trim(rs1("CarSimpleID"))%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),15,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
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
			<td><span class="style3"><%
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
			<td><span class="style3"><%
			'批號
			response.write "&nbsp;"
			%></span></td>
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
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation in ('30','31','32') and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/16+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID,f.EquipMentID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation in ('30','31','32') and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
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
	<table width="100%" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><span class="style2"><%=TitleUnitName%>&nbsp;逕行舉發移送清冊</span></td>
		</tr>
		<tr>
			<td align="left"><span class="style3">站所：<%
		response.write "<strong><font class=""style6"">"&"高雄市交通事件裁決所"&"</font></strong>"
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/16)+1%> of <%=pagecnt%></span></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"></td>
			<td width="10%"><span class="style3">單號</span></td>
			<td width="9%"><span class="style3"><!-- 違規日期 --></span></td>
			<td width="9%"></td>
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="18%"><span class="style3">舉發單位</span></td>
			<td width="9%"><span class="style3">員警</span></td>
			<td width="10%"><span class="style3">扣件</span></td>
			<td width="9%"><span class="style3"><!-- 備註 --></span></td>
		</tr>
		<tr>
			<td><span class="style3"><!-- 編號 --></span></td>
			<td><span class="style3">入案日期</span></td>
			<td><span class="style3">違規日期<!-- 違規時間 --></span></td>
			<td><span class="style3">車號</span></td>
			<td><span class="style3">法條</span></td>
			<td><span class="style3">駕駛人/車主</span></td>
			<td><span class="style3">駕籍資料</span></td>
			<td></td>
			<td><span class="style3">車籍資料</span></td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 16
			if rs1.eof then exit for
			Response.flush
%>
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"><span class="style3"><%
			PrintSNtotal=PrintSNtotal+1
			PrintSN=PrintSN+1
			'response.write PrintSNtotal
			%></span></td>
			<td width="10%"><span class="style3"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write rs1("BillNO")
				else
					response.write "<strong>"&rs1("BillNO")&"</strong>"
				end if
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="9%"><span class="style3"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="9%"><span class="style3"><%response.write trim(rs1("CarNo"))%></span></td>
			<td width="8%"><span class="style3"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="18%"><span class="style3"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),15,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="18%"><span class="style3"><%
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
			<td width="9%"><span class="style3"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="10%"><span class="style3"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="9%"><span class="style3"><%
			'檔名
			response.write "&nbsp;"
			%></span></td>
		</tr>
		<tr>
			<td></td>
			<td><span class="style3"><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
			'if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			'	response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			'else
				response.write "&nbsp;"
			'end if
			%></span></td>
			<td><span class="style3"><%response.write trim(rs1("CarSimpleID"))%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),15,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
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
			<td><span class="style3"><%
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
			<td><span class="style3"><%
			'批號
			response.write "&nbsp;"
			%></span></td>
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
	'其他堅理所列表
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
	if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then
		DciStationName=""
		strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(StationArray(SA))&"'"
		set rsSN=conn.execute(strSqlStationName)
			DciStationName=trim(rsSN("DCIstationName"))
		if not rsSN.eof then
		end if
		rsSN.close
		set rsSN=nothing
		PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/16+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID,f.EquipMentID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and (((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
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
	<table width="100%" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><span class="style2"><%=TitleUnitName%>&nbsp;逕行舉發移送清冊</span></td>
		</tr>
		<tr>
			<td align="left"><span class="style3">站所：<%
		response.write "<strong><font class=""style6"">"&DciStationName&"</font></strong>"
	%>&nbsp; &nbsp; &nbsp; &nbsp;移送日期：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/16)+1%> of <%=pagecnt%></span></td>
		</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"></td>
			<td width="10%"><span class="style3">單號</span></td>
			<td width="9%"><span class="style3"><!-- 違規日期 --></span></td>
			<td width="9%"></td>
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="18%"><span class="style3">舉發單位</span></td>
			<td width="9%"><span class="style3">員警</span></td>
			<td width="10%"><span class="style3">扣件</span></td>
			<td width="9%"><span class="style3"><!-- 備註 --></span></td>
		</tr>
		<tr>
			<td><span class="style3"><!-- 編號 --></span></td>
			<td><span class="style3">入案日期</span></td>
			<td><span class="style3">違規日期<!-- 違規時間 --></span></td>
			<td><span class="style3">車號</span></td>
			<td><span class="style3">法條</span></td>
			<td><span class="style3">駕駛人/車主</span></td>
			<td><span class="style3">駕籍資料</span></td>
			<td></td>
			<td><span class="style3">車籍資料</span></td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to 16
			if rs1.eof then exit for
			Response.flush
%>
	<tr>
	<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="0%"><span class="style3"><%
			PrintSNtotal=PrintSNtotal+1
			PrintSN=PrintSN+1
			'response.write PrintSNtotal
			%></span></td>
			<td width="10%"><span class="style3"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write rs1("BillNO")
				else
					response.write "<strong>"&rs1("BillNO")&"</strong>"
				end if
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="9%"><span class="style3"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><span class="style3"><%response.write trim(rs1("CarNo"))%></span></td>
			<td width="8%"><span class="style3"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="18%"><span class="style3"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),15,1)
			else
				response.write "&nbsp;"
			end if	
			%></span></td>
			<td width="18%"><span class="style3"><%
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
			<td width="9%"><span class="style3"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="10%"><span class="style3"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></span></td>
			<td width="9%"><span class="style3"><%
			'檔名
			response.write "&nbsp;"
			%></span></td>
		</tr>
		<tr>
			<td></td>
			<td><span class="style3"><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
			'if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
			'	response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			'else
				response.write "&nbsp;"
			'end if
			%></span></td>
			<td><span class="style3"><%response.write trim(rs1("CarSimpleID"))%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			%></span></td>
			<td><span class="style3"><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),15,1)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td><span class="style3"><%
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
			<td><span class="style3"><%
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
			<td><span class="style3"><%
			'批號
			response.write "&nbsp;"
			%></span></td>
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

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>