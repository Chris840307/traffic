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
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"苗栗縣" and sys_City<>"嘉義縣" then%>
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
.style4 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
.style5 {font-family:新細明體; color=0044ff; line-height:13px; font-size: 11px}
.style6 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="苗栗縣" or sys_City="嘉義縣" then%>
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
<title>攔停舉發移送清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 6800
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

	If sys_City="苗栗縣" Then
		if trim(request("Selt_MemberStation"))<>"" then strwhere=strwhere&" and f.MemberStation='"&trim(request("Selt_MemberStation"))&"'"

		strB="select distinct(a.BatchNumber) from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
		set rsB=conn.execute(strB)
		While Not rsB.Eof
			strBDel="Delete from batchnumberjob where batchNumber='"&Trim(rsB("Batchnumber"))&"' and PrintTypeID=0"
			conn.execute strBDel

			strBIns="Insert into batchnumberjob values('"&Trim(rsB("Batchnumber"))&"',"&Trim(session("User_ID"))&",0,sysdate)"
			conn.execute strBIns
		rsB.MoveNext
		Wend
		rsB.close
		Set rsB=Nothing 
	End If 

	strStation="select distinct(f.MemberStation) from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere&" order by MemberStation"
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then
		rsStation.MoveFirst 
	else
		response.write "查無資料，請確認此批舉發單是攔停舉發單!"
	end if
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
<%if (sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"台南市" and sys_City<>"台中市" and sys_City<>"苗栗縣" and sys_City<>"南投縣") or (sys_City="南投縣" and trim(Session("Unit_ID"))="05CB") then
	if sys_City="高雄縣" then
		response.write "<br><br><br><br><br>"
	end if
%>
<center><font size="3">舉發違反道路交通事件通知單攔停舉發移送清冊</font></center>
	<table width="600" border="1" cellpadding="3" cellspacing="0" align="center">
		<tr>
			<td width="33%" align="center"><span class="style3">受文單位</span></td>
			<td width="33%" align="center"><span class="style3">移送件數</span></td>
			<td width="33%" align="center"><span class="style3">備考</span></td>
		</tr>
<%	StationCntTotal=0
	'台北市交通事件裁決所
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
%>
		<tr>
			<td><span class="style3"><%
			'受文單位
			response.write "台北市交通事件裁決所"
			%></span></td>
			<td align="center"><span class="style3"><%
			'件數
		'攔停
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ('20','21','22','23','24','29') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and f.MemberStation in ('20','21','22','23','24','29') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ('30','31','32') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and f.MemberStation in ('30','31','32') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation='"&trim(StationArray(SA))&"' and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		strCloseCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('S','d','e')" &_
		" and f.MemberStation='"&trim(StationArray(SA))&"' and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
<%else%>
<%
	StationCntTotal=0
	'台北市交通事件裁決所
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"29")>0 then
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ('20','21','22','23','24','29') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
		'攔停
		StationCnt=0
		strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and f.MemberStation in ('30','31','32') and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
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
			StationCnt=0
			strCntReport="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and f.MemberStation='"&trim(StationArray(SA))&"' and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere
			set rsCntReport=conn.execute(strCntReport)
			if not rsCntReport.eof then
				StationCnt=StationCnt+trim(rsCntReport("cnt"))
			end if
			rsCntReport.close
			set rsCntReport=nothing
			StationCntTotal=StationCntTotal+StationCnt
		end if
	Next
	Response.flush
%>
<%end if%>
<%
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

	strUnitName="select Value from ApConfigure where ID=40"
	set rsUnitName=conn.execute(strUnitName)
	if not rsUnitName.eof then
             if sys_City="高雄縣" then
		TitleUnitName="高雄縣政府警察局 "&TitleUnitName2
             Else
				if sys_City="苗栗縣" Then
					TitleUnitName=trim(rsUnitName("value"))
				Else
					TitleUnitName=trim(rsUnitName("value"))&" "&TitleUnitName2
				End If 
                
             end if
	end if
	rsUnitName.close
	set rsUnitName=nothing

	PrintSNtotal=0	'編號
	if sys_City="嘉義縣" or sys_City="嘉義市" or sys_City="南投縣" or sys_City="台南市" or sys_City="宜蘭縣" or sys_City="屏東縣" or sys_City="台東縣" then
		PageCount=20
	else
		PageCount=23
	end if

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
				pagecnt=fix(Cint(rsCnt("cnt"))/PageCount+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.RuleSpeed, f.IllegalSpeed,f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber,f.DealLineDate,f.BillFillDate" &_
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
			if PageNum>1 then
				response.write "<div class=""PageNext""></div>"
			end if
		end if
	if sys_City="高雄縣" then
		response.write "<br><br><br><br><br>"
	end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><b><%=TitleUnitName%></b>&nbsp;&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left"><%
	if sys_City="高雄市" then
		strS="select * from station where dcistationid='22'"
		set rsS=conn.execute(strS)
		if not rsS.eof then
			response.write trim(rsS("stationaddress"))
		end if
		rsS.close
		set rsS=nothing 
	end if 
			%><br>站所：<%
		response.write DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;<%
	if sys_City="台中市" then
		response.write "列印日期"
	else
		response.write "移送日期"
	end if
	%>：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;<%
	if sys_City="苗栗縣Q" then
		if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID in (select UnitTypeID from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"')"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write "分局名稱："&trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
		end If
	Else
		response.write "(本批案件已透過中華電信數據分公司作入案管制)"
	End if
	%>&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/PageCount)+1%> of <%=pagecnt%></td>
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
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="11%">舉發單位</td>
			<td width="9%">員警</td>
			<td width="10%">扣件</td>
			<%if sys_City="雲林縣" then%>
			<td width="11%">應到案日期</span></td>
			<%else%>
			<td width="11%">備註<span class='style4'>&nbsp;&nbsp;&nbsp;(超重)</span></td>
			<%end if%>
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
			<td>車籍資料</td>
			<%if sys_City="雲林縣" then%>
			<td>填單日期</td>
			<%else%>
			<td></td>
			<%end if%>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to PageCount
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
			<td width="8%"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="18%"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(trim(rs1("Driver")),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td width="11%"><span class="style6"><%
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
			<td width="9%"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="10%"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write "<b><span class='style3'>" & rsBF("Content") & "</span></b>"				
				'smith add <b> tag
				'response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td width="11%"><%
			if sys_City="雲林縣" then
				'應到案日
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write gInitDT(rs1("DealLineDate"))
				else
					response.write "&nbsp;"
				end if
			else
				'檔名
				response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			end if
			%></td>
		</tr>
		<tr>
			<td><%

			%></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				'response.write trim(rs1("DCICaseInDate"))
				'smith make font small
				response.write "<span class='style4'>"& trim(rs1("DCICaseInDate"))&"</span>"
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
			<td><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				response.write "<br>"&trim(rs1("Rule4"))
			end if
			%></td>
			<td><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
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
			<td><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof Then
					if trim(rs1("DCIErrorCarData"))="F" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					elseif trim(rs1("DCIErrorCarData"))="$" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					else
						response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
					end if
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if sys_City="雲林縣" then
				'填單日
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write gInitDT(rs1("BillFillDate"))
				else
					response.write "&nbsp;"
				end if
			else
				'批號
				response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
				if sys_City="苗栗縣" Then
					response.write "<strong>"
					strRSeq="select count(*) as cnt from billbase where SN in (select Billsn from Dcilog where batchnumber='"&trim(rs1("BatchNumber"))&"') and RecordDate<to_date('"&year(rs1("RecordDate"))&"/"&month(rs1("RecordDate"))&"/"&day(rs1("RecordDate"))&" "&hour(rs1("RecordDate"))&":"&minute(rs1("RecordDate"))&":"&second(rs1("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0"
					Set rsRSeq=conn.execute(strRSeq)
					If Not rsRSeq.eof Then
						response.write "_"
						response.write Trim(rsRSeq("cnt"))+1
					End If
					rsRSeq.close
					Set rsRSeq=Nothing 

					response.write "</strong>"
				End if
				'超重部份
				Sys_IllegalSpeed="":Sys_RuleSpeed=""
				Sys_IllegalSpeed=trim(rs1("IllegalSpeed"))
				Sys_RuleSpeed=trim(rs1("RuleSpeed"))
				SYS_Rule1=trim(rs1("Rule1"))
				
				if left(SYS_Rule1,2)="29" and trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
					response.write  "<span class='style5'><strong>&nbsp;"&Sys_IllegalSpeed-Sys_RuleSpeed&"噸</strong></span>"
				end if		
			end if
			%></td>
		</tr>
		</table>
		</td>
		</tr>
<%			Response.flush
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
	攔停資料請核對移送表與二、(三)聯違規通知單聯、扣件物，倘若不符請通知本局，逾三日未回覆視同相符無誤。
	<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%

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
				pagecnt=fix(Cint(rsCnt("cnt"))/PageCount+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.RuleSpeed, f.IllegalSpeed,f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber,f.DealLineDate,f.BillFillDate" &_
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
<%		end if
		if PageNum>1 then
			response.write "<div class=""PageNext""></div>"
		end if
		
	if sys_City="高雄縣" then
		response.write "<br><br><br><br><br>"
	end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><b><%=TitleUnitName%></b>&nbsp;&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left"><%
	if sys_City="高雄市" then
		strS="select * from station where dcistationid='32'"
		set rsS=conn.execute(strS)
		if not rsS.eof then
			response.write trim(rsS("stationaddress"))
		end if
		rsS.close
		set rsS=nothing 
	end if 
			%><br>站所：<%
		response.write DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;<%
	if sys_City="台中市" then
		response.write "列印日期"
	else
		response.write "移送日期"
	end if
	%>：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;<%
	if sys_City="苗栗縣Q" then
		if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID in (select UnitTypeID from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"')"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write "分局名稱："&trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
		end If
	Else
		response.write "(本批案件已透過中華電信數據分公司作入案管制)"
	End if
	%>&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/PageCount)+1%> of <%=pagecnt%></td>
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
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="11%">舉發單位</td>
			<td width="9%">員警</td>
			<td width="10%">扣件</td>
			<%if sys_City="雲林縣" then%>
			<td width="11%">應到案日期</span></td>
			<%else%>
			<td width="11%">備註<span class='style4'>&nbsp;&nbsp;&nbsp;(超重)</span></td>
			<%end if%>
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
			<td>車籍資料</td>
			<%if sys_City="雲林縣" then%>
			<td>填單日期</td>
			<%else%>
			<td></td>
			<%end if%>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to PageCount
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
			<td width="8%"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="18%"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(trim(rs1("Driver")),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td width="11%"><span class="style6"><%
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
			<td width="9%"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="10%"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof				
				response.write "<b><span class='style3'>" & rsBF("Content") & "</span></b>"
				'smith add <b> tag
				'response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td width="11%"><%
			if sys_City="雲林縣" then
				'應到案日
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write gInitDT(rs1("DealLineDate"))
				else
					response.write "&nbsp;"
				end if
			else
				'檔名
				response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			end if
			%></td>
		</tr>
		<tr>
			<td><%
				
			%></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				'response.write trim(rs1("DCICaseInDate"))
				'smith make font small
				response.write "<span class='style4'>"& trim(rs1("DCICaseInDate"))&"</span>"				
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
			<td><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				response.write "<br>"&trim(rs1("Rule4"))
			end if
			%></td>
			<td><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
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
			<td><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					if trim(rs1("DCIErrorCarData"))="F" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					elseif trim(rs1("DCIErrorCarData"))="$" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					else
						response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
					end if
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if sys_City="雲林縣" then
				'填單日
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write gInitDT(rs1("BillFillDate"))
				else
					response.write "&nbsp;"
				end if
			else
				response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
				if sys_City="苗栗縣" Then
					response.write "<strong>"
					strRSeq="select count(*) as cnt from billbase where SN in (select Billsn from Dcilog where batchnumber='"&trim(rs1("BatchNumber"))&"') and RecordDate<to_date('"&year(rs1("RecordDate"))&"/"&month(rs1("RecordDate"))&"/"&day(rs1("RecordDate"))&" "&hour(rs1("RecordDate"))&":"&minute(rs1("RecordDate"))&":"&second(rs1("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0"
					Set rsRSeq=conn.execute(strRSeq)
					If Not rsRSeq.eof Then
						response.write "_"
						response.write Trim(rsRSeq("cnt"))+1
					End If
					rsRSeq.close
					Set rsRSeq=Nothing 

					response.write "</strong>"
				End if
				'超重部份
				Sys_IllegalSpeed="":Sys_RuleSpeed=""
				Sys_IllegalSpeed=trim(rs1("IllegalSpeed"))
				Sys_RuleSpeed=trim(rs1("RuleSpeed"))
				SYS_Rule1=trim(rs1("Rule1"))
				
				if left(SYS_Rule1,2)="29" and trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
					response.write  "<span class='style5'><strong>&nbsp;"&Sys_IllegalSpeed-Sys_RuleSpeed&"噸</strong></span>"
				end if			
			end if
			%></td>
		</tr>
		</table>
		</td>
		</tr>
<%			Response.flush
		rs1.MoveNext
		next
%>
	</table>
<%
		Wend
		rs1.close
		set rs1=nothing

%>
	共計： <%=PrintSN%>  &nbsp;筆
	<br>
	攔停資料請核對移送表與二、(三)聯違規通知單聯、扣件物，倘若不符請通知本局，逾三日未回覆視同相符無誤。
	<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
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
				pagecnt=fix(Cint(rsCnt("cnt"))/PageCount+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select f.RuleSpeed, f.IllegalSpeed,f.SN,f.BillNo,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate,e.DCIReturnCarType" &_
			",f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,e.Driver,e.DriverHomeZip,e.DriverHomeAddress" &_
			",f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData,e.DCIErrorIDData" &_
			",e.Owner,f.TrafficAccidentType,d.DCIReturnStatus,a.FileName,a.BatchNumber,f.DealLineDate,f.BillFillDate" &_
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
<%		end if
		if PageNum>1 then
			response.write "<div class=""PageNext""></div>"
		end if
	if sys_City="高雄縣" then
		response.write "<br><br><br><br><br>"
	end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><b><%=TitleUnitName%></b>&nbsp;&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left"><%
	if sys_City="高雄市" then
		strS="select * from station where dcistationid='"&trim(StationArray(SA))&"'"
		set rsS=conn.execute(strS)
		if not rsS.eof then
			response.write trim(rsS("stationaddress"))
		end if
		rsS.close
		set rsS=nothing 
	end if 
			%><br>站所：<%
		response.write trim(StationArray(SA))&" "&DciStationName
	%>&nbsp; &nbsp; &nbsp; &nbsp;<%
	if sys_City="台中市" then
		response.write "列印日期"
	else
		response.write "移送日期"
	end if
	%>：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;<%
	if sys_City="苗栗縣Q" then
		if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID in (select UnitTypeID from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"')"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write "分局名稱："&trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
		end If
	Else
		response.write "(本批案件已透過中華電信數據分公司作入案管制)"
	End if
	%>&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/PageCount)+1%> of <%=pagecnt%></td>
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
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="11%">舉發單位</td>
			<td width="9%">員警</td>
			<td width="10%">扣件</td>
			<%if sys_City="雲林縣" then%>
			<td width="11%">應到案日期</span></td>
			<%else%>
			<td width="11%">備註<span class='style4'>&nbsp;&nbsp;&nbsp;(超重)</span></td>
			<%end if%>
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
			<td>車籍資料</td>
			<%if sys_City="雲林縣" then%>
			<td>填單日期</td>
			<%else%>
			<td></td>
			<%end if%>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to PageCount
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
			<td width="9%" ><%response.write trim(rs1("CarNo"))%></td>
			<td width="8%"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="18%"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(trim(rs1("Driver")),14,1)
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td width="11%"><span class="style6"><%
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
			<td width="9%"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="10%"><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;</b>"
			end if
			While Not rsBF.Eof
				response.write "<b><span class='style3'>" & rsBF("Content") & "</span></b>"
				'smith add <b> tag
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td width="11%"><%
			if sys_City="雲林縣" then
				'應到案日
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write gInitDT(rs1("DealLineDate"))
				else
					response.write "&nbsp;"
				end if
			else
				'檔名
				response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			end if
			%></td>
		</tr>
		<tr>
			<td><%
				
			%></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then			
				'smith make font small
				response.write "<span class='style4'>"& trim(rs1("DCICaseInDate"))&"</span>"
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
			<td><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				response.write "<br>"&trim(rs1("Rule4"))
			end if
			%></td>
			<td><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),14,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
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
			<td><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					if trim(rs1("DCIErrorCarData"))="F" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					elseif trim(rs1("DCIErrorCarData"))="$" then
						response.write "<strong>"&trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))&"</strong>"
					else
						response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
					end if
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if sys_City="雲林縣" then
				'填單日
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write gInitDT(rs1("BillFillDate"))
				else
					response.write "&nbsp;"
				end if
			else
				'批號
				response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
				if sys_City="苗栗縣" Then
					response.write "<strong>"
					strRSeq="select count(*) as cnt from billbase where SN in (select Billsn from Dcilog where batchnumber='"&trim(rs1("BatchNumber"))&"') and RecordDate<to_date('"&year(rs1("RecordDate"))&"/"&month(rs1("RecordDate"))&"/"&day(rs1("RecordDate"))&" "&hour(rs1("RecordDate"))&":"&minute(rs1("RecordDate"))&":"&second(rs1("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0"
					Set rsRSeq=conn.execute(strRSeq)
					If Not rsRSeq.eof Then
						response.write "_"
						response.write Trim(rsRSeq("cnt"))+1
					End If
					rsRSeq.close
					Set rsRSeq=Nothing 

					response.write "</strong>"
				End if
				'超重部份
				Sys_IllegalSpeed="":Sys_RuleSpeed=""
				Sys_IllegalSpeed=trim(rs1("IllegalSpeed"))
				Sys_RuleSpeed=trim(rs1("RuleSpeed"))
				SYS_Rule1=trim(rs1("Rule1"))
				
				if left(SYS_Rule1,2)="29" and trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
					response.write  "<span class='style5'><strong>&nbsp;"&Sys_IllegalSpeed-Sys_RuleSpeed&"噸</strong></span>"
				end if
			end if			
			%></td>
		</tr>
		</table>
		</td>
		</tr>
<%			Response.flush
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
	攔停資料請核對移送表與二、(三)聯違規通知單聯、扣件物，倘若不符請通知本局，逾三日未回覆視同相符無誤。
	<br>
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
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="苗栗縣" or sys_City="嘉義縣" then%>
window.print();
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
<%conn.close%>