<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>寄存送達清冊(已結案)</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)
%>
<%
	if sys_City="台中市" then
		CloseDciReturnStatusID="DciReturnStatusID in ('n')"
	elseif sys_City="南投縣" or sys_City="台中縣" then
		CloseDciReturnStatusID="DciReturnStatusID not in ('S','N','h')"
	else
		CloseDciReturnStatusID="DciReturnStatusID not in ('S','N')"
	end if
	strwhere=request("SQLstr")
	'逕舉的到案處所用BillBaseDCIReturn
	ReportStationArrayTemp=""
	strStReport="select distinct(e.DCIReturnStation) from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
	",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
	" and a.ReturnMarkType='4'" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a" &_
		" ,BillBaseDCIReturn e where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')"
	set rsStReport=conn.execute(strStReport)
	If Not rsStReport.Bof Then 
		rsStReport.MoveFirst 
	else
		response.write "查無資料!"
	end if	
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
	strStStop="select distinct(f.MemberStation) from DCILog a,MemberData b" &_
	",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
	" and f.RecordStateID=0" &_
	" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
	" and a.DCIReturnStatusID='S' and a.ReturnMarkType='4'" &_
	" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
	" and a.RecordMemberID=b.MemberID(+) "&strwhere
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
			if StationArrayTemp="" then
				StationArrayTemp=ReportStationArray(RSA)
			else
				StationArrayTemp=StationArrayTemp&","&ReportStationArray(RSA)
			end if
		end if
	next
	for SSA=0 to ubound(StopStationArray)
		if instr(StationArrayTemp,StopStationArray(SSA))=0 then
			if StationArrayTemp="" then
				StationArrayTemp=StopStationArray(SSA)
			else
				StationArrayTemp=StationArrayTemp&","&StopStationArray(SSA)
			end if
		end if
	next
%>
</head>
<body>
<form name=myForm method="post">
<%if sys_City<>"台南市" then %>
	<center><font size="3">舉發違反道路交通事件通知單寄存送達(已結案)移送清冊</font></center>
	<table width="80%" border="1" cellpadding="3" cellspacing="0" align="center">
		<tr>
			<td width="33%" align="center"><span class="style3">受文單位</span></td>
			<td width="33%" align="center"><span class="style3">移送件數</span></td>
			<td width="33%" align="center"><span class="style3">備考</span></td>
		</tr>
<%
	StationCntTotal=0
	StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫

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
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation in ('20','21','22','23','24','29')"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.MemberStation in ('20','21','22','23','24','29')" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
'		CloseCnt1=0
'		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.BillTypeID='2'" &_
'			" and a.ReturnMarkType='4' and a.DciReturnStatusID='n'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
'			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
'			" and e.DCIReturnStation in ('20','21','22','23','24','29')"
'		set rsCntReport=conn.execute(strCntReport)
'		if not rsCntReport.eof then
'			CloseCnt1=cint(trim(rsCntReport("cnt")))
'		end if
'		rsCntReport.close
'		set rsCntReport=nothing
'
'		'攔停
'		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.MemberStation in ('20','21','22','23','24','29')" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID<>'2'" &_
'			" and a.ReturnMarkType='4'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere
'		set rsCntStop=conn.execute(strCntStop)
'		if not rsCntStop.eof then
'			CloseCnt1=CloseCnt1+cint(trim(rsCntStop("cnt")))
'		end if
'		rsCntStop.close
'		set rsCntStop=nothing
'
'		if CloseCnt1>0 then
'			response.write "結案 "&CloseCnt1&" 件"
'		else
			response.write "&nbsp;"
'		end if
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
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation in ('30','31','32')"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.MemberStation in ('30','31','32')" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
'		CloseCnt2=0
'		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.BillTypeID='2'" &_
'			" and a.ReturnMarkType='4' and a.DciReturnStatusID='n'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
'			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
'			" and e.DCIReturnStation in ('30','31','32')"
'		set rsCntReport=conn.execute(strCntReport)
'		if not rsCntReport.eof then
'			CloseCnt2=cint(trim(rsCntReport("cnt")))
'		end if
'		rsCntReport.close
'		set rsCntReport=nothing
'
'		'攔停
'		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.MemberStation in ('30','31','32')" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.BillTypeID<>'2'" &_
'			" and a.ReturnMarkType='4' and a.DciReturnStatusID='n'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere
'		set rsCntStop=conn.execute(strCntStop)
'		if not rsCntStop.eof then
'			CloseCnt2=CloseCnt2+cint(trim(rsCntStop("cnt")))
'		end if
'		rsCntStop.close
'		set rsCntStop=nothing
'
'		if CloseCnt2>0 then
'			response.write "結案 "&CloseCnt2&" 件"
'		else
			response.write "&nbsp;"
'		end if
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
			if StationNameArray="" then
				StationNameArray=trim(rsSN("DCIstationName"))
			else
				StationNameArray=StationNameArray&","&trim(rsSN("DCIstationName"))
			end if
			response.write trim(rsSN("DCIstationName"))
		end if
		rsSN.close
		set rsSN=nothing
			%></span></td>
			<td align="center"><span class="style3"><%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"'"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
'		CloseCnt3=0
'		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.BillTypeID='2'" &_
'			" and a.ReturnMarkType='4' and a.DciReturnStatusID='n'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
'			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
'			" and e.DCIReturnStation='"&trim(StationArray(SA))&"'"
'		set rsCntReport=conn.execute(strCntReport)
'		if not rsCntReport.eof then
'			CloseCnt3=cint(trim(rsCntReport("cnt")))
'		end if
'		rsCntReport.close
'		set rsCntReport=nothing
'
'		'攔停
'		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
'			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.BillTypeID<>'2'" &_
'			" and a.ReturnMarkType='4' and a.DciReturnStatusID='n'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) "&strwhere
'		set rsCntStop=conn.execute(strCntStop)
'		if not rsCntStop.eof then
'			CloseCnt3=CloseCnt3+cint(trim(rsCntStop("cnt")))
'		end if
'		rsCntStop.close
'		set rsCntStop=nothing
'
'		if CloseCnt3>0 then
'			response.write "結案 "&CloseCnt3&" 件"
'		else
			response.write "&nbsp;"
'		end if
		%></span></td>
		</tr>
<%		else
			if StationNameArray="" then
				StationNameArray=" "
			else
				StationNameArray=StationNameArray&", "
			end if
		end if
	next
%>
		<tr>
			<td><span class="style3">小計</span></td>
			<td align="center"><span class="style3"><%=StationCntTotal%></span></td>
			<td>&nbsp;</td>
		</tr>
	</table>
	<center><%
	PageNum=1
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
	<div class="PageNext"></div>
<%else%>
<%
	StationCntTotal=0
	StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫

	'台北市交通裁決所數量
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation in ('20','21','22','23','24','29')"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.MemberStation in ('20','21','22','23','24','29')" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
	end if

	'高雄市交通事件裁決所數量
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
			" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation in ('30','31','32')"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.MemberStation in ('30','31','32')" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
			" and a.ReturnMarkType='4'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
	end if

	'其他監理所數量
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then
			strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(StationArray(SA))&"'"
			set rsSN=conn.execute(strSqlStationName)
			if not rsSN.eof then
				if StationNameArray="" then
					StationNameArray=trim(rsSN("DCIstationName"))
				else
					StationNameArray=StationNameArray&","&trim(rsSN("DCIstationName"))
				end if
			end if
			rsSN.close
			set rsSN=nothing
			StationCnt=0
			strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b" &_
				",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
				" and f.RecordStateID=0" &_
				" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
				" and a.ReturnMarkType='4'" &_
				" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
				" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
				" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
				" and e.DCIReturnStation='"&trim(StationArray(SA))&"'"
			set rsCntReport=conn.execute(strCntReport)
			if not rsCntReport.eof then
				StationCnt=StationCnt+trim(rsCntReport("cnt"))
			end if
			rsCntReport.close
			set rsCntReport=nothing

			'攔停
			strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
				",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
				" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
				" and f.RecordStateID=0" &_
				" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID<>'2'" &_
				" and a.ReturnMarkType='4'" &_
				" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
				" and a.RecordMemberID=b.MemberID(+) "&strwhere
			set rsCntStop=conn.execute(strCntStop)
			if not rsCntStop.eof then
				StationCnt=StationCnt+trim(rsCntStop("cnt"))
			end if
			rsCntStop.close
			set rsCntStop=nothing
			StationCntTotal=StationCntTotal+StationCnt
		else
			if StationNameArray="" then
				StationNameArray=" "
			else
				StationNameArray=StationNameArray&", "
			end if
		end if
	next
%>
<%end if%>
<%	
	CaseSn=0
	'台北市交通裁決所列表
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
	'逕舉
	PrintSN=0
	strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.Rule1,f.Rule2,f.Rule3" &_
		",e.billcloseid " &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID,a.ExchangeTypeID,a.DciReturnStatusID from DciLog a where a.BillSN is not null "&strwhere&") a,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN" &_
		" and f.RecordStateID=0 and f.SN=g.BillSn" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation in ('20','21','22','23','24','29') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation in ('20','21','22','23','24','29') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<center><font size="3">舉發違反道路交通事件通知單寄存送達(已結案)移送清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%="台北市交通事件裁決所"%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>
			<td align="center">送達書號</td>
			<td align="center">貼條號碼</td>
			<td align="center">退件原因</td>
			<td align="center">單退狀態/送達狀態</td>
		</tr>
<%		
	for i=1 to 45
		if rs1.eof then exit for
		PrintSN=PrintSN+1
%>		<tr>
			<td align="center"><%
			'編號
			CaseSn=CaseSn+1
			response.write CaseSn
			%></td>
			<td align="center" nowrap><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center" nowrap><%
			'車號
			if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
				response.write trim(rs1("CarNO"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),18,1)
			else
				response.write "&nbsp;"
			end if				
			%></td>
			<td align="center" nowrap><%
			'法條一
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條二
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條三
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write trim(rs1("Rule3"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'送達書號
			ReturnReason=""
			MailNumberTmp=""
			strMail="select MailNumber,StoreAndSendMailNumber,StoreAndSendGovNumber,ReturnResonID,StoreAndSendReturnResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
			set rsMail=conn.execute(strMail)
			if not rsMail.eof then
				if trim(rsMail("StoreAndSendGovNumber"))="" or isnull(rsMail("StoreAndSendGovNumber")) then
					response.write "&nbsp;"
				else
					response.write trim(rsMail("StoreAndSendGovNumber"))
				end if
				if (trim(rsMail("StoreAndSendMailNumber"))="" or isnull(rsMail("StoreAndSendMailNumber"))) and (trim(rsMail("MailNumber"))<>"" and not isnull(rsMail("MailNumber"))) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("MailNumber"))
					
				elseif trim(rsMail("StoreAndSendMailNumber"))<>"" and not isnull(rsMail("StoreAndSendMailNumber")) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("StoreAndSendMailNumber"))
				end if
				'退件原因
					strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
					set rsCode=conn.execute(strCode)
					if not rsCode.eof then
						ReturnReason=trim(rsCode("Content"))
					end if
					rsCode.close
					set rsCode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%

			if MailNumberTmp="" then
				response.write "&nbsp;"
			else
				response.write MailNumberTmp
			end if
			rsMail.close
			set rsMail=nothing
			%></td>
			<td nowrap><%
			'退件原因
			if ReturnReason="" then
				response.write "&nbsp;"
			else
				response.write ReturnReason
			end if
			%></td>
			<td nowrap><%
			'單退狀態
			strStatus1="select DciReturnStatusID from Dcilog a where a.ExchangeTypeID='N' and a.ReturnMarkType='4' and a.BillSn="&trim(rs1("BillSn"))&strwhere
			set rsStatus1=conn.execute(strStatus1)
			if not rsStatus1.eof then
				strStatus="select * from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsStatus1("DciReturnStatusID"))&"'"
				set rsStatus=conn.execute(strStatus)
				if not rsStatus.eof then
					response.write trim(rsStatus("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsStatus.close
				set rsStatus=nothing
				
			else
				response.write "&nbsp;"
			end if	
			rsStatus1.close
			set rsStatus1=nothing

			if sys_City="台中市" then
				strBillDciClose="select a.billcloseid,b.Content from BillBaseDCIReturn a,DciCode b" &_
				" where a.BillNO='"&trim(rs1("BillNO"))&"'" &_
				" and a.CarNo='"&trim(rs1("CarNo"))&"' and" &_
				" a.ExchangeTypeID='N' and a.billcloseid=b.ID and B.TypeID='9'" 
				set rsBDciClose=conn.execute(strBillDciClose)
				if not rsBDciClose.eof then
					response.write " / "&trim(rsBDciClose("billcloseid"))&trim(rsBDciClose("Content"))
				else
					response.write "&nbsp;"
				end if
				rsBDciClose.close
				set rsBDciClose=nothing
			end if
			%></td>
		</tr>
<%
		rs1.MoveNext
		next
%>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
		Wend
		rs1.close
		set rs1=nothing
%>
<%if SA2<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	end if

	'高雄市交通事件裁決所列表
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
	'逕舉
	PrintSN=0
	strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.Rule1,f.Rule2,f.Rule3" &_
		",e.billcloseid " &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID,a.ExchangeTypeID,a.DciReturnStatusID from DciLog a where a.BillSN is not null "&strwhere&") a,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN" &_
		" and f.RecordStateID=0 and f.SN=g.BillSn" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<center><font size="3">舉發違反道路交通事件通知單寄存送達(已結案)移送清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%="高雄市交通事件裁決所"%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>
			<td align="center">送達書號</td>
			<td align="center">貼條號碼</td>
			<td align="center">退件原因</td>
			<td align="center">單退狀態/送達狀態</td>
		</tr>
<%		
	for i=1 to 45
		if rs1.eof then exit for
		PrintSN=PrintSN+1
%>		<tr>
			<td align="center"><%
			'編號
			CaseSn=CaseSn+1
			response.write CaseSn
			%></td>
			<td align="center" nowrap><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center" nowrap><%
			'車號
			if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
				response.write trim(rs1("CarNO"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),18,1)
			else
				response.write "&nbsp;"
			end if				
			%></td>
			<td align="center" nowrap><%
			'法條一
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條二
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條三
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write trim(rs1("Rule3"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'送達書號
			ReturnReason=""
			MailNumberTmp=""
			strMail="select MailNumber,StoreAndSendMailNumber,StoreAndSendGovNumber,ReturnResonID,StoreAndSendReturnResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
			set rsMail=conn.execute(strMail)
			if not rsMail.eof then
				if trim(rsMail("StoreAndSendGovNumber"))="" or isnull(rsMail("StoreAndSendGovNumber")) then
					response.write "&nbsp;"
				else
					response.write trim(rsMail("StoreAndSendGovNumber"))
				end if
				if (trim(rsMail("StoreAndSendMailNumber"))="" or isnull(rsMail("StoreAndSendMailNumber"))) and (trim(rsMail("MailNumber"))<>"" and not isnull(rsMail("MailNumber"))) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("MailNumber"))
					
				elseif trim(rsMail("StoreAndSendMailNumber"))<>"" and not isnull(rsMail("StoreAndSendMailNumber")) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("StoreAndSendMailNumber"))
				end if
				'退件原因
					strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
					set rsCode=conn.execute(strCode)
					if not rsCode.eof then
						ReturnReason=trim(rsCode("Content"))
					end if
					rsCode.close
					set rsCode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%

			if MailNumberTmp="" then
				response.write "&nbsp;"
			else
				response.write MailNumberTmp
			end if
			rsMail.close
			set rsMail=nothing
			%></td>
			<td nowrap><%
			'退件原因
			if ReturnReason="" then
				response.write "&nbsp;"
			else
				response.write ReturnReason
			end if
			%></td>
			<td nowrap><%
			'單退狀態
			strStatus1="select DciReturnStatusID from Dcilog a where a.ExchangeTypeID='N' and a.ReturnMarkType='4' and a.BillSn="&trim(rs1("BillSn"))&strwhere
			set rsStatus1=conn.execute(strStatus1)
			if not rsStatus1.eof then
				strStatus="select * from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsStatus1("DciReturnStatusID"))&"'"
				set rsStatus=conn.execute(strStatus)
				if not rsStatus.eof then
					response.write trim(rsStatus("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsStatus.close
				set rsStatus=nothing
				
			else
				response.write "&nbsp;"
			end if	
			rsStatus1.close
			set rsStatus1=nothing

			if sys_City="台中市" then
				strBillDciClose="select a.billcloseid,b.Content from BillBaseDCIReturn a,DciCode b" &_
				" where a.BillNO='"&trim(rs1("BillNO"))&"'" &_
				" and a.CarNo='"&trim(rs1("CarNo"))&"' and" &_
				" a.ExchangeTypeID='N' and a.billcloseid=b.ID and B.TypeID='9'" 
				set rsBDciClose=conn.execute(strBillDciClose)
				if not rsBDciClose.eof then
					response.write " / "&trim(rsBDciClose("billcloseid"))&trim(rsBDciClose("Content"))
				else
					response.write "&nbsp;"
				end if
				rsBDciClose.close
				set rsBDciClose=nothing
			end if
			%></td>
		</tr>
<%
		rs1.MoveNext
		next
%>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
	Wend
	rs1.close
	set rs1=nothing
%>
<%if SA2<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	end if


	StationName=split(StationNameArray,",")
	for SA2=0 to ubound(StationArray)
		if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA2)))<=0 then
	'逕舉
	PrintSN=0
	strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.Rule1,f.Rule2,f.Rule3" &_
		",e.billcloseid " &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID,a.ExchangeTypeID,a.DciReturnStatusID from DciLog a where a.BillSN is not null "&strwhere&") a,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN" &_
		" and f.RecordStateID=0 and f.SN=g.BillSn" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<center><font size="3">舉發違反道路交通事件通知單寄存送達(已結案)移送清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%=StationName(SA2)%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>
			<td align="center">送達書號</td>
			<td align="center">貼條號碼</td>
			<td align="center">退件原因</td>
			<td align="center">單退狀態/送達狀態</td>
		</tr>
<%		
	for i=1 to 45
		if rs1.eof then exit for
		PrintSN=PrintSN+1
%>		<tr>
			<td align="center"><%
			'編號
			CaseSn=CaseSn+1
			response.write CaseSn
			%></td>
			<td align="center" nowrap><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center" nowrap><%
			'車號
			if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
				response.write trim(rs1("CarNO"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),18,1)
			else
				response.write "&nbsp;"
			end if				
			%></td>
			<td align="center" nowrap><%
			'法條一
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條二
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'法條三
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write trim(rs1("Rule3"))
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td align="center" nowrap><%
			'送達書號
			ReturnReason=""
			MailNumberTmp=""
			strMail="select MailNumber,StoreAndSendMailNumber,StoreAndSendGovNumber,ReturnResonID,StoreAndSendReturnResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
			set rsMail=conn.execute(strMail)
			if not rsMail.eof then
				if trim(rsMail("StoreAndSendGovNumber"))="" or isnull(rsMail("StoreAndSendGovNumber")) then
					response.write "&nbsp;"
				else
					response.write trim(rsMail("StoreAndSendGovNumber"))
				end if
				if (trim(rsMail("StoreAndSendMailNumber"))="" or isnull(rsMail("StoreAndSendMailNumber"))) and (trim(rsMail("MailNumber"))<>"" and not isnull(rsMail("MailNumber"))) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("MailNumber"))
					
				elseif trim(rsMail("StoreAndSendMailNumber"))<>"" and not isnull(rsMail("StoreAndSendMailNumber")) then
					'貼條號碼
					MailNumberTmp=trim(rsMail("StoreAndSendMailNumber"))
				end if
				'退件原因
					strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
					set rsCode=conn.execute(strCode)
					if not rsCode.eof then
						ReturnReason=trim(rsCode("Content"))
					end if
					rsCode.close
					set rsCode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%

			if MailNumberTmp="" then
				response.write "&nbsp;"
			else
				response.write MailNumberTmp
			end if
			rsMail.close
			set rsMail=nothing
			%></td>
			<td nowrap><%
			'退件原因
			if ReturnReason="" then
				response.write "&nbsp;"
			else
				response.write ReturnReason
			end if
			%></td>
			<td nowrap><%
			'單退狀態
			strStatus1="select DciReturnStatusID from Dcilog a where a.ExchangeTypeID='N' and a.ReturnMarkType='4' and a.BillSn="&trim(rs1("BillSn"))&strwhere
			set rsStatus1=conn.execute(strStatus1)
			if not rsStatus1.eof then
				strStatus="select * from DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsStatus1("DciReturnStatusID"))&"'"
				set rsStatus=conn.execute(strStatus)
				if not rsStatus.eof then
					response.write trim(rsStatus("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsStatus.close
				set rsStatus=nothing
				
			else
				response.write "&nbsp;"
			end if	
			rsStatus1.close
			set rsStatus1=nothing

			if sys_City="台中市" then
				strBillDciClose="select a.billcloseid,b.Content from BillBaseDCIReturn a,DciCode b" &_
				" where a.BillNO='"&trim(rs1("BillNO"))&"'" &_
				" and a.CarNo='"&trim(rs1("CarNo"))&"' and" &_
				" a.ExchangeTypeID='N' and a.billcloseid=b.ID and B.TypeID='9'" 
				set rsBDciClose=conn.execute(strBillDciClose)
				if not rsBDciClose.eof then
					response.write " / "&trim(rsBDciClose("billcloseid"))&trim(rsBDciClose("Content"))
				else
					response.write "&nbsp;"
				end if
				rsBDciClose.close
				set rsBDciClose=nothing
			end if
			%></td>
		</tr>
<%
		rs1.MoveNext
		next
%>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
	Wend
	rs1.close
	set rs1=nothing
%>
<%if SA2<>ubound(StationArray) then%>
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
window.print();

</script>
<%conn.close%>