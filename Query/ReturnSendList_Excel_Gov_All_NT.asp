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
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"嘉義縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體;  line-height:19px;font-size: 12pt}
.style5 {font-family:新細明體;  line-height:14px;font-size: 8pt}
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
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>退件清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->

<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)
%>
<%
	if sys_City="台中市" then
		CloseDciReturnStatusID="DciReturnStatusID in ('n','h','S','N')"
	else
		CloseDciReturnStatusID="DciReturnStatusID is not null"
	end if
	strwhere=request("SQLstr")
	'逕舉的到案處所用BillBaseDCIReturn
	ReportStationArrayTemp=""
	strStReport="select distinct(DCIReturnStation) from (select f.SN,a.BillNo,a.CarNo from DCILog a,MemberData b," &_
		"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&" and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+)" &_
		" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a" &_
		" ,BillBaseDCIReturn e,BillMailHistory g where a.SN=g.BillSn and a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L') and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q') order by DCIReturnStation"
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
	'response.write strStReport
	'response.end
	'攔停的到案處所用MemberStation
	StopStationArrayTemp=""
	strStStop="select distinct(f.MemberStation) from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.SN=g.BillSn and f.RecordStateID=0" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and a.BillTypeID<>'2' and a.ExchangeTypeID=d.DCIActionID(+)"&_
		" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')"&strwhere& "  order by f.MemberStation"
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
	Next
	SortStationArrayTemp=StationArrayTemp
	Arr_SortStationArrayTemp=Split(SortStationArrayTemp,",")
	for i=0 to ubound(Arr_SortStationArrayTemp)
        vFlag   = true
        vTempNO = i
        vTemp   = cint(Arr_SortStationArrayTemp(vTempNO))

        do while(vFlag)
            if 0 < vTempNO then
                if vTemp < cint(Arr_SortStationArrayTemp(vTempNO - 1)) then  'pe
                    Arr_SortStationArrayTemp(vTempNO)    = Arr_SortStationArrayTemp(vTempNO - 1)
                    Arr_SortStationArrayTemp(vTempNO - 1)= vTemp
                else
                    vFlag   = false
                end if
            else
                vFlag   = false
            end if
            vTempNO = vTempNO - 1
        loop
    Next
    StationArrayTemp=""
	For i=0 To UBound(Arr_SortStationArrayTemp)
		If StationArrayTemp="" Then
			StationArrayTemp=Arr_SortStationArrayTemp(i)
		Else
			StationArrayTemp=StationArrayTemp&","&Arr_SortStationArrayTemp(i)
		End If 
	Next 
%>

</head>
<body>
<form name=myForm method="post">
<%if (sys_City<>"台南市" and sys_City<>"南投縣") or (sys_City="南投縣" and (trim(Session("Unit_ID"))="05CB" or trim(Session("Unit_ID"))="05FG")) then %>
	<center><font size="3">舉發違反道路交通事件通知單退件移送清冊</font></center>
	<%
	if sys_City="高雄市" then
		if instr(trim(strwhere),"BatchNumber")>0 then
			response.write "<center>批號："&replace(mid(trim(strwhere),instr(trim(strwhere),"BatchNumber")+16,instr(trim(strwhere),"')")-(instr(trim(strwhere),"BatchNumber")+16)),"'","")&"</center>"
		end if
	end if
	%>
	<table width="80%" border="1" cellpadding="3" cellspacing="0" align="center">
		<tr>
			<td width="33%" align="center"><span class="style3">受文單位</span></td>
			<td width="33%" align="center"><span class="style3">移送件數</span></td>
			<td width="33%" align="center"><span class="style3">備考</span></td>
		</tr>
<%	StationCntTotal=0
	StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫
	StationCntArray=""	'將每個處所的件數存到陣列

	

	'其他間理所數量=========================================================================
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		'if instr("20,21,22,23,24,25,26,30,31,32,40,41,46,60,61,63,68",trim(StationArray(SA)))<=0 then
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
		strCntReport="select count(*) as cnt from (select f.SN,a.BillNo,a.CarNo from DCILog a,MemberData b," &_
			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
			" and a.BillTypeID='2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e,BillMailHistory g" &_
			" where a.BillNo=e.BillNO and a.SN=g.BillSn and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')"
	
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing

		'攔停
		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.SN=g.BillSn and f.RecordStateID=0" &_
		" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q') "&strwhere
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
		response.write StationCnt
			%></span></td>
			<td><span class="style3"><%
			'結案件數
		'逕舉
'		CloseCnt3=0
'		strCntReport="select count(*) as cnt from (select f.SN,a.BillNo,a.CarNo from DCILog a,MemberData b," &_
'			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
'			" and f.RecordStateID=0" &_
'			" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n'" &_
'			" and a.BillTypeID='2'" &_
'			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e,BillMailHistory g" &_
'			" where a.BillNo=e.BillNO and a.SN=g.BillSn and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n')" &_
'			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')"
'	
'		set rsCntReport=conn.execute(strCntReport)
'		if not rsCntReport.eof then
'			CloseCnt3=cint(trim(rsCntReport("cnt")))
'		end if
'		rsCntReport.close
'		set rsCntReport=nothing
'
'		'攔停
'		strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
'		",DCIReturnStatus d,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
'		" and f.SN=g.BillSn and f.RecordStateID=0" &_
'		" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
'		" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n'" &_
'		" and a.BillTypeID<>'2'" &_
'		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
'		" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q') "&strwhere
'		set rsCntStop=conn.execute(strCntStop)
'		if not rsCntStop.eof then
'			CloseCnt3=CloseCnt3+cint(trim(rsCntStop("cnt")))
'		end if
'		rsCntStop.close
'		set rsCntStop=nothing
'		if CloseCnt3>0 then
'			response.write "結案 "&CloseCnt3&" 件"
'		else
			response.write "&nbsp;"
'		end if
			%></span></td>
		</tr>
<%		'else
'			if StationNameArray="" then
'				StationNameArray=" "
'			else
'				StationNameArray=StationNameArray&","&" "
'			end if
'			if StationCntArray="" then
'				StationCntArray=0
'			else
'				StationCntArray=StationCntArray&",0"
'			end if
'		end if
	next
%>
		<tr>
			<td><span class="style3">小計</span></td>
			<td align="center"><span class="style3"><%=StationCntTotal%></span></td>
			<td>&nbsp;</td>
		</tr>
	</table>
	<div class="PageNext">&nbsp;</div>
<%else%>
<%
	StationCntTotal=0
	StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫
	StationCntArray=""	'將每個處所的件數存到陣列
	
	'其他間理所數量=====================================================
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		'if instr("20,21,22,23,24,25,26,30,31,32,40,41,46,60,61,63,68",trim(StationArray(SA)))<=0 then
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
		strCntReport="select count(*) as cnt from (select f.SN,a.BillNo,a.CarNo from DCILog a,MemberData b," &_
			"DCIReturnStatus d,BillBase f where a.BillSN=f.SN" &_
			" and f.RecordStateID=0" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
			" and a.BillTypeID='2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' "&strwhere&") a,BillBaseDCIReturn e,BillMailHistory g" &_
			" where a.BillNo=e.BillNO and a.SN=g.BillSn and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L')" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')"
	
			set rsCntReport=conn.execute(strCntReport)
			if not rsCntReport.eof then
				StationCnt=StationCnt+trim(rsCntReport("cnt"))
			end if
			rsCntReport.close
			set rsCntReport=nothing

			'攔停
			strCntStop="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
			" and f.SN=g.BillSn and f.RecordStateID=0" &_
			" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
			" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and a.ReturnMarkType='3' and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q') "&strwhere
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
'		else
'			if StationNameArray="" then
'				StationNameArray=" "
'			else
'				StationNameArray=StationNameArray&","&" "
'			end if
'			if StationCntArray="" then
'				StationCntArray=0
'			else
'				StationCntArray=StationCntArray&",0"
'			end if
'		end if
	next
%>
<%end if%>
<%	StationName=split(StationNameArray,",")
	StationCnt=split(StationCntArray,",")

	TitleValue=""
	strUnitName2="select UnitName,UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		if sys_City="屏東縣" then
			TitleUnitName2=replace(rsUnitName2("UnitName"),"屏東縣政府警察局","")
		elseif sys_City="高雄市" then
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

	strTitle="select Value from Apconfigure where ID=40"
	set rsTitle=conn.execute(strTitle)
	if not rsTitle.eof then
		TitleValue=rsTitle("Value")&" "&TitleUnitName2
	end if
	rsTitle.close
	set rsTitle=nothing

	'其他監理所列表=============================================================================
	pagetmp=0
	for SA2=0 to ubound(StationName)
	'response.write StationCntArray
	'response.write StationCnt(SA2)
	'if instr("20,21,22,23,24,25,26,30,31,32,40,41,46,60,61,63,68",trim(StationArray(SA2)))<=0 then
	PrintSN=0
if pagetmp>0 then%>
	<div class="PageNext">&nbsp;</div>
<%end if
	pagetmp=pagetmp+1
%>

<%		'逕舉and攔停
		strSQL="select a.BillSN,a.BillNO,a.CarNO,a.BatchNumber,e.Driver,e.Owner,f.CarSimpleID,f.IllegalDate" &_
		",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,f.BillMem1,f.BillMem2,f.BillTypeID" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID,a.BatchNumber,a.ExchangeTypeID,a.DciReturnStatusID from DciLog a where a.BillSN is not null "&strwhere&") a" &_
		" ,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and f.SN=g.BillSn" &_
		" and a.ExchangeTypeID='N' and a."&CloseDciReturnStatusID&"" &_
		" and ((a.BillTypeID='2' and e.DCIReturnStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))" &_
		" or (a.BillTypeID<>'2' and f.MemberStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n','L'))) and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')" &_
		" order by g.UserMarkDate"
'response.write strSQL
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>		
	<table width="710" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" height="28" colspan="2"><span class="style4"><%
		response.write TitleValue&"&nbsp退件(公示)資料"

		if trim(StationCnt(SA2))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(trim(StationCnt(SA2)))/23+0.9999999)
		end if
	%></span></td>
	</tr>
	<tr>
	<td width="80%">到案處所：<%=trim(StationArray(SA2)) & " " & StationName(SA2)%>
	&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
	列印日期：<%=now%>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <%
	if sys_City="基隆市" then
		if trim(rs1("BatchNumber"))<>"" and not isnull(rs1("BatchNumber")) then
			response.write "作業批號："&trim(rs1("BatchNumber"))
		end if	
	end if
	%>
	</td>
	<td align="right" width="20%">
	Page <%=fix(PrintSN/23)+1%> of <%=pagecnt%></td></td>
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
					<td width="16%"><%
					if sys_City="基隆市" then
						response.write "舉發單位"
					end if
					%></td>
					<td width="10%"></td>
					<td width="11%"></td>
				</tr>
				<tr>
					<td></td>
					<td>違規時間</td>
					<td>車號</td>
					<td>法條</td>
					<td>駕駛人/車主</td>
					<td><%
					if sys_City="基隆市" then
						response.write "車主證號"
					else
						response.write "舉發單位"
					end if
					%></td>
					<td>員警</td>
					<td>退件原因</td>
				</tr>
			</table>
			</td>
		<tr>
<%		for i=1 to 23
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
				if len(rsUnit("UnitName"))>7 then
					response.write "<span class=""style5"">"&rsUnit("UnitName")&"</span>"
				else
					response.write rsUnit("UnitName")
				end if
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
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					response.write "<br>"&trim(rs1("Rule3"))
				end if	
				%></td>
					<td><%
			if sys_City="台東縣" or ((sys_City="基隆市" Or sys_City="南投縣") and trim(rs1("BillTypeID"))="1") then
				'駕駛姓名
				if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
					response.write funcCheckFont(trim(rs1("Driver")),18,1)&"/"
				else
					response.write "&nbsp;"
				end if		
			end if
				'車主姓名
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write funcCheckFont(rs1("Owner"),18,1)
				else
					response.write "&nbsp;"
				end if				
				%></td>
					<td><%
				'車主證號
				if sys_City="基隆市" then
					strRet="select OwnerID from BillBaseDciReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='N'"
					set rsRet=conn.execute(strRet)
					if not rsRet.eof then
						if trim(rsRet("OwnerID"))<>"" and not isnull(rsRet("OwnerID")) then
							response.write trim(rsRet("OwnerID"))
						else
							response.write "&nbsp;"
						end if		
					end if
					rsRet.close
					set rsRet=nothing
				end if
				%></td>
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
<%
	'end if
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
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
window.print();
<%elseif sys_City="宜蘭縣" then%>
printWindow(true,7,5.08,5.08,5.08);
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
<%conn.close%>