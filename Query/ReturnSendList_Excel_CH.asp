<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
.style4 {font-family:新細明體;  line-height:19px;font-size: 12pt}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>退件清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
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
%>
<%
	strwhere=request("SQLstr")
	'逕舉的到案處所用BillBaseDCIReturn
	ReportStationArrayTemp=""
	strStReport="select distinct(DCIReturnStation) from (select a.BillNo,a.CarNo from DCILog a,MemberData b," &_
		"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID='N' and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a" &_
		" ,BillBaseDCIReturn e where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')"
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
	'response.write strStReport
	'response.end
	'攔停的到案處所用MemberStation
	StopStationArrayTemp=""
	strStStop="select distinct(f.MemberStation) from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID='N'" &_
		" and a.BillTypeID<>'2' and a.ExchangeTypeID=d.DCIActionID(+)"&_
		" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere
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
<%	StationCntTotal=0
	StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫
	StationCntArray=""	'將每個處所的件數存到陣列

	'台北市交通裁決所數量
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
%>
		<%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b," &_
			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N'" &_
			" and a.BillTypeID='2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e" &_
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
		" and f.RecordStateID=0" &_
		" and f.MemberStation in ('20','21','22','23','24','29')" &_
		" and a.ExchangeTypeID='N'" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		StationCntTaipei=StationCnt	'台北數量
			%>
<%
	end if

	'高雄市交通事件裁決所數量
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
%>
		<%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b," &_
			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N'" &_
			" and a.BillTypeID='2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e" &_
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
		" and f.RecordStateID=0" &_
		" and f.MemberStation in ('30','31','32')" &_
		" and a.ExchangeTypeID='N'" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		StationCntKS=StationCnt	'高雄數量
			%>
<%
	end if

	'其他間理所數量
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA)))<=0 then
%>
		<%
			'受文單位
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
			%>
			<%
			'件數
		'逕舉
		StationCnt=0
		strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b," &_
			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N'" &_
			" and a.BillTypeID='2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e" &_
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
		" and f.RecordStateID=0" &_
		" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
		" and a.ExchangeTypeID='N'" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere
		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt

		if StationCntArray="" then
			StationCntArray=StationCnt
		else
			StationCntArray=StationCntArray&","&StationCnt
		end if
			%>
<%		else
			if StationNameArray="" then
				StationNameArray=" "
			else
				StationNameArray=StationNameArray&","&" "
			end if
			if StationCntArray="" then
			StationCntArray=0
			else
				StationCntArray=StationCntArray&",0"
			end if
		end if
	next
%>

<%	StationName=split(StationNameArray,",")
	StationCnt=split(StationCntArray,",")

	TitleValue=""
	strUnitName2="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		TitleUnitName2=trim(rsUnitName2("UnitName"))
	end if
	rsUnitName2.close
	set rsUnitName2=nothing

	strTitle="select Value from Apconfigure where ID=40"
	set rsTitle=conn.execute(strTitle)
	if not rsTitle.eof then
		TitleValue=rsTitle("Value")&" "&TitleUnitName2
	end if
	rsTitle.close
	set rsTitle=nothing

	'台北市交通裁決所舉發單列表
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
	PrintSN=0
%>
<%		'逕舉and攔停
		strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.CarSimpleID,f.IllegalDate" &_
		",f.Rule1,f.Rule2,f.BillUnitID,f.BillMem1,f.BillMem2" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID from DciLog a where a.BillSN is not null "&strwhere&") a" &_
		" ,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and f.SN=g.BillSn" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation in ('20','21','22','23','24','29') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation in ('20','21','22','23','24','29') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"

		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>		
	<table width="710" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" height="28" colspan="2"><span class="style4"><%
		
		response.write TitleValue&"&nbsp(所有退件)資料"

		if trim(StationCnt(SA2))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(trim(StationCntTaipei))/21+0.9999999)
		end if
	%></span></td>
	</tr>
	<tr>
	<td width="65%">到案處所：<%="台北市交通事件裁決所"%>
	&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
	列印日期：<%=now%>
	</td>
	<td align="right" width="35%">
	Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td></td>
	</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%">單號</td>
					<td width="10%">違規日期</td>
					<td width="10%"></td>
					<td width="10%"></td>
					<td width="23%"></td>
					<td width="16%"></td>
					<td width="10%"></td>
					<td width="11%"></td>
				</tr>
				<tr>
					<td></td>
					<td>違規時間</td>
					<td>車號</td>
					<td>法條</td>
					<td>駕駛人/車主</td>
					<td>舉發單位</td>
					<td>員警</td>
					<td>退件原因</td>
				</tr>
			</table>
			</td>
		<tr>
<%		for i=1 to 21
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%"><%
				'單號
				if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
					response.write trim(rs1("BillNO"))
				else
					response.write "&nbsp;"
				end if
				%></td>
					<td width="10%"><%
					'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td width="10%"><%
				'車號
				if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
					response.write trim(rs1("CarNO"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="10%"><%
				'法條一
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="23%"></td>
					<td width="16%"><%
					'舉發單位
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write rsUnit("UnitName")
			else
				response.write "&nbsp;"
			end if
			rsUnit.close
			set rsUnit=nothing
					%></td>
					<td width="10%"><%
					'員警1
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td width="11%"><%
				ReturnReason=""
				strMail="select MailNumber,StoreAndSendMailNumber,ReturnResonID,StoreAndSendReturnResonID,OpenGovResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
				set rsMail=conn.execute(strMail)
				if not rsMail.eof then
					'退件原因
						strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
						set rsCode=conn.execute(strCode)
						if not rsCode.eof then
							response.write trim(rsMail("UserMarkResonID"))&" "&trim(rsCode("Content"))
						end if
						rsCode.close
						set rsCode=nothing
				else
					response.write "&nbsp;"
				end if

				rsMail.close
				set rsMail=nothing
				%></td>
				</tr>
				<tr>
					<td></td>
					<td><%
					'違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td><%
					'車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				response.write trim(rs1("CarSimpleID"))
			else
				response.write "&nbsp;"
			end if	
					%></td>
					<td><%
				'法條二
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td><%
				'車主姓名
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write funcCheckFont(rs1("Owner"),14,1)
				else
					response.write "&nbsp;"
				end if				
				%></td>
					<td></td>
					<td><%
					'員警2
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write trim(rs1("BillMem2"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td></td>
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
<%if SA2<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	end if

		'高雄市交通事件裁決所列表
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
	PrintSN=0
%>
<%		'逕舉and攔停
		strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.CarSimpleID,f.IllegalDate" &_
		",f.Rule1,f.Rule2,f.BillUnitID,f.BillMem1,f.BillMem2" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID from DciLog a where a.BillSN is not null "&strwhere&") a" &_
		" ,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and f.SN=g.BillSn" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"

		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>		
	<table width="710" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" height="28" colspan="2"><span class="style4"><%
		response.write TitleValue&"&nbsp(所有退件)資料"

		if trim(StationCnt(SA2))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(trim(StationCntKS))/21+0.9999999)
		end if
	%></span></td>
	</tr>
	<tr>
	<td width="65%">到案處所：<%="高雄市交通事件裁決所"%>
	&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
	列印日期：<%=now%>
	</td>
	<td align="right" width="35%">
	Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td></td>
	</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%">單號</td>
					<td width="10%">違規日期</td>
					<td width="10%"></td>
					<td width="10%"></td>
					<td width="23%"></td>
					<td width="16%"></td>
					<td width="10%"></td>
					<td width="11%"></td>
				</tr>
				<tr>
					<td></td>
					<td>違規時間</td>
					<td>車號</td>
					<td>法條</td>
					<td>駕駛人/車主</td>
					<td>舉發單位</td>
					<td>員警</td>
					<td>退件原因</td>
				</tr>
			</table>
			</td>
		<tr>
<%		for i=1 to 21
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%"><%
				'單號
				if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
					response.write trim(rs1("BillNO"))
				else
					response.write "&nbsp;"
				end if
				%></td>
					<td width="10%"><%
					'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td width="10%"><%
				'車號
				if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
					response.write trim(rs1("CarNO"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="10%"><%
				'法條一
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="23%"></td>
					<td width="16%"><%
					'舉發單位
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write rsUnit("UnitName")
			else
				response.write "&nbsp;"
			end if
			rsUnit.close
			set rsUnit=nothing
					%></td>
					<td width="10%"><%
					'員警1
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td width="11%"><%
				ReturnReason=""
				strMail="select MailNumber,StoreAndSendMailNumber,ReturnResonID,StoreAndSendReturnResonID,OpenGovResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
				set rsMail=conn.execute(strMail)
				if not rsMail.eof then
					'退件原因
						strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
						set rsCode=conn.execute(strCode)
						if not rsCode.eof then
							response.write trim(rsMail("UserMarkResonID"))&" "&trim(rsCode("Content"))
						end if
						rsCode.close
						set rsCode=nothing
				else
					response.write "&nbsp;"
				end if

				rsMail.close
				set rsMail=nothing
				%></td>
				</tr>
				<tr>
					<td></td>
					<td><%
					'違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td><%
					'車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				response.write trim(rs1("CarSimpleID"))
			else
				response.write "&nbsp;"
			end if	
					%></td>
					<td><%
				'法條二
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td><%
				'車主姓名
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write funcCheckFont(rs1("Owner"),14,1)
				else
					response.write "&nbsp;"
				end if				
				%></td>
					<td></td>
					<td><%
					'員警2
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write trim(rs1("BillMem2"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td></td>
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
<%if SA2<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	end if

	'其他監理所列表
	for SA2=0 to ubound(StationName)
	'response.write StationCntArray
	'response.write StationCnt(SA2)
	if instr("20,21,22,23,24,29,30,31,32",trim(StationArray(SA2)))<=0 then
	PrintSN=0
%>

<%		'逕舉and攔停
		strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,f.CarSimpleID,f.IllegalDate" &_
		",f.Rule1,f.Rule2,f.BillUnitID,f.BillMem1,f.BillMem2" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID from DciLog a where a.BillSN is not null "&strwhere&") a" &_
		" ,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and f.SN=g.BillSn" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')))" &_
		" order by g.UserMarkDate"
'response.write strSQL
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>		
	<table width="710" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" height="28" colspan="2"><span class="style4"><%
		response.write TitleValue&"&nbsp(所有退件)資料"

		if trim(StationCnt(SA2))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(trim(StationCnt(SA2)))/21+0.9999999)
		end if
	%></span></td>
	</tr>
	<tr>
	<td width="65%">到案處所：<%=StationName(SA2)%>
	&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
	列印日期：<%=now%>
	</td>
	<td align="right" width="35%">
	Page <%=fix(PrintSN/21)+1%> of <%=pagecnt%></td></td>
	</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%">單號</td>
					<td width="10%">違規日期</td>
					<td width="10%"></td>
					<td width="10%"></td>
					<td width="23%"></td>
					<td width="16%"></td>
					<td width="10%"></td>
					<td width="11%"></td>
				</tr>
				<tr>
					<td></td>
					<td>違規時間</td>
					<td>車號</td>
					<td>法條</td>
					<td>駕駛人/車主</td>
					<td>舉發單位</td>
					<td>員警</td>
					<td>退件原因</td>
				</tr>
			</table>
			</td>
		<tr>
<%		for i=1 to 21
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="10%"><%
				'單號
				if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
					response.write trim(rs1("BillNO"))
				else
					response.write "&nbsp;"
				end if
				%></td>
					<td width="10%"><%
					'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td width="10%"><%
				'車號
				if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
					response.write trim(rs1("CarNO"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="10%"><%
				'法條一
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="23%"></td>
					<td width="16%"><%
					'舉發單位
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write rsUnit("UnitName")
			else
				response.write "&nbsp;"
			end if
			rsUnit.close
			set rsUnit=nothing
					%></td>
					<td width="10%"><%
					'員警1
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td width="11%"><%
				ReturnReason=""
				strMail="select MailNumber,StoreAndSendMailNumber,ReturnResonID,StoreAndSendReturnResonID,OpenGovResonID,UserMarkResonID from BillMailHistory where BillSN='"&trim(rs1("BillSN"))&"'"
				set rsMail=conn.execute(strMail)
				if not rsMail.eof then
					'退件原因
						strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rsMail("UserMarkResonID"))&"'"
						set rsCode=conn.execute(strCode)
						if not rsCode.eof then
							response.write trim(rsMail("UserMarkResonID"))&" "&trim(rsCode("Content"))
						end if
						rsCode.close
						set rsCode=nothing
				else
					response.write "&nbsp;"
				end if

				rsMail.close
				set rsMail=nothing
				%></td>
				</tr>
				<tr>
					<td></td>
					<td><%
					'違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td><%
					'車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				response.write trim(rs1("CarSimpleID"))
			else
				response.write "&nbsp;"
			end if	
					%></td>
					<td><%
				'法條二
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td><%
				'車主姓名
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write funcCheckFont(rs1("Owner"),14,1)
				else
					response.write "&nbsp;"
				end if				
				%></td>
					<td></td>
					<td><%
					'員警2
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write trim(rs1("BillMem2"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td></td>
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

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>