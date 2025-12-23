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
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
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
<title>寄存送達期滿清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)
%>
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	sDay1=request("Day1")
	sDay2=request("Day2")
    strwhere= " ( a.UserMarkDate between " &_
              " To_Date('" & gOutDT(sDay1)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS')" &_
     		  " and To_Date('" & gOutDT(sDay2)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS') )"

	'逕舉的到案處所用BillBaseDCIReturn

	ReportStationArrayTemp=""
	strStReport="select distinct(e.DCIReturnStation) from (select f.BillNo,f.CarNo from mailstationreturn a" &_
		",BillBase f where " & strwhere &_
		" and a.billno=f.billno and f.RecordStateID=0" & ") a " &_
		" ,BillBaseDCIReturn e where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'"

    'response.write strStReport
    'response.end

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
	strStStop="select distinct(f.MemberStation) from mailstationreturn a" &_
		",BillBase f where a.billno=f.billno" &_
		" and f.RecordStateID=0" &_
		" and f.Billtypeid='1'" &_
		" and " & strwhere
	'response.write strStStop
    'response.end
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
	<center><font size="3">寄存期滿清冊</font></center>
	<table width="70%" border="1" cellpadding="2" cellspacing="0" align="center">
		<tr>
			<td width="33%" align="center"><span class="style3">受文單位</span></td>
			<td width="67%" align="center"><span class="style3">移送件數</span></td>
		   <!--	<td width="33%" align="center"><span class="style3">備考</span></td> -->
		</tr>
        <%	StationCntTotal=0
            StationNameArray=""	'將到案處所中文名稱存到陣列裡,清冊就不用再讀資料庫
            '台北市交通裁決所數量
            if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"25")>0 or instr(StationArrayTemp,"26")>0 then
		pagenum=1        
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
		strCntReport="select count(*) as cnt from (select f.BillNo,f.CarNo from mailstationreturn a" &_
                    ",BillBase f where " & strwhere &_
                    " and a.billno=f.billno and f.RecordStateID=0" & ") a " &_
                    " ,BillBaseDCIReturn e ,billbase f where a.BillNo=e.BillNO and a.billno=f.billno and f.Billtypeid='2' and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'" &_

                    " and e.DCIReturnStation in ('20','21','22','23','24','25','26')"

		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing


		'攔停
		strCntStop="select count(*) as cnt from mailstationreturn a" &_
                    ",BillBase f , BillBaseDCIReturn e where " & strwhere &_
                    " and a.billno=f.billno" &_
                    " and f.RecordStateID=0" &_
                    " and f.billtypeid<>'2'" &_
                    " and a.billno=e.BillNO and e.ExchangeTypeID='W' and e.Status='Y'" &_
                    " and f.MemberStation in ('20','21','22','23','24','25','26')"

		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
		<!--	<td><span class="style3"> --><%
			'結案件數
		'逕舉
		'CloseCnt1=0
		'strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID='2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
		'" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'" &_
		'" and e.DCIReturnStation in ('20','21','22','23','24','25','26')"
		'set rsCntReport=conn.execute(strCntReport)
		'if not rsCntReport.eof then
		'	CloseCnt1=cint(trim(rsCntReport("cnt")))
		'end if
		'rsCntReport.close
		'set rsCntReport=nothing

		'攔停
		'strCntStop="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and f.MemberStation in ('20','21','22','23','24','25','26')" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID<>'2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere
		'set rsCntStop=conn.execute(strCntStop)
		'if not rsCntStop.eof then
		'	CloseCnt1=CloseCnt1+cint(trim(rsCntStop("cnt")))
		'end if
		'rsCntStop.close
		'set rsCntStop=nothing
        '
		'if CloseCnt1>0 then
		'	response.write "結案 "&CloseCnt1&" 件"
		'else
		'	response.write "&nbsp;"
		'end if
			%><!--</span></td>-->
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
		strCntReport="select count(*) as cnt from (select f.BillNo,f.CarNo from mailstationreturn a" &_
					",BillBase f where " & strwhere &_
					" and a.billno=f.billno and f.RecordStateID=0" & ") a " &_
					" ,BillBaseDCIReturn e ,billbase f where a.BillNo=e.BillNO and a.billno=f.billno and f.billtypeid='2' and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'" &_
					" and e.DCIReturnStation in ('30','31','32')"
		set rsCntReport=conn.execute(strCntReport)
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		'response.write strCntReport
		'攔停
		strCntStop="select count(*) as cnt from mailstationreturn a" &_
                    ",BillBase f , BillBaseDCIReturn e where " & strwhere &_
                    " and a.billno=f.billno" &_
                    " and f.RecordStateID=0" &_
                    " and a.billno=e.BillNO and e.ExchangeTypeID='W' and e.Status='Y'" &_
                    " and f.billtypeid<>'2'" &_                    
					" and f.MemberStation in ('30','31','32')"

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
		'CloseCnt2=0
		'strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID='2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
		'" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'" &_
		'" and e.DCIReturnStation in ('30','31','32')"
		'set rsCntReport=conn.execute(strCntReport)
		'if not rsCntReport.eof then
		'	CloseCnt2=cint(trim(rsCntReport("cnt")))
		'end if
		'rsCntReport.close
		'set rsCntReport=nothing
		'攔停
		'strCntStop="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and f.MemberStation in ('30','31','32')" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID<>'2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere
		'set rsCntStop=conn.execute(strCntStop)
		'if not rsCntStop.eof then
		'	CloseCnt2=CloseCnt2+cint(trim(rsCntStop("cnt")))
		'end if
		'rsCntStop.close
		'set rsCntStop=nothing
        '
		'if CloseCnt2>0 then
		'	response.write "結案 "&CloseCnt2&" 件"
		'else
		'	response.write "&nbsp;"
		'end if
		%> <!--  </span></td> -->
		</tr>
<%
	end if

	'其他監理所數量
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
		if instr("20,21,22,23,24,25,26,30,31,32",trim(StationArray(SA)))<=0 then
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
		strCntReport="select count(*) as cnt from (select f.BillNo,f.CarNo from mailstationreturn a" &_
                    ",BillBase f where " & strwhere &_
                    " and a.billno=f.billno and f.billtypeid='2' and f.RecordStateID=0" & ") a " &_
                    " ,BillBaseDCIReturn e where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' "&_
		    "  and e.Status='Y'" &_
                    " and e.DCIReturnStation='"&trim(StationArray(SA))&"'"
		set rsCntReport=conn.execute(strCntReport)
if trim(StationArray(SA))="44" then
	'response.write strCntReport
	'response.write "<br>"
end if
		if not rsCntReport.eof then
			StationCnt=StationCnt+trim(rsCntReport("cnt"))
		end if
		rsCntReport.close
		set rsCntReport=nothing
		'response.write strCntReport
		'攔停
		strCntStop="select count(*) as cnt from mailstationreturn a" &_
                    ",BillBase f  ,BillBaseDCIReturn e where " & strwhere &_
                    " and a.billno=f.billno" &_
                    " and f.RecordStateID=0" &_
					" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
					" and f.BillTypeID<>'2'" &_
                    " and a.billno=e.BillNO and e.ExchangeTypeID='W' and e.Status='Y'"

		set rsCntStop=conn.execute(strCntStop)
		if not rsCntStop.eof then
			StationCnt=StationCnt+trim(rsCntStop("cnt"))
		end if
if trim(StationArray(SA))="44" then
	'response.write strCntStop
	'response.write "<br>"
end if
		rsCntStop.close
		set rsCntStop=nothing
		StationCntTotal=StationCntTotal+StationCnt
		response.write StationCnt
			%></span></td>
		   <!--	<td><span class="style3"> --><%
			'結案件數
		'逕舉
		'CloseCnt3=0
		'strCntReport="select count(*) as cnt from (select a.BillNo,a.CarNo from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID='2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere&") a,BillBaseDCIReturn e" &_
		'	" where a.BillNo=e.BillNO and a.CarNo=e.CarNo and e.ExchangeTypeID='W' and e.Status='Y'" &_
		'	" and e.DCIReturnStation='"&trim(StationArray(SA))&"'"
		'set rsCntReport=conn.execute(strCntReport)
		'if not rsCntReport.eof then
		'	CloseCnt3=cint(trim(rsCntReport("cnt")))
		'end if
		'rsCntReport.close
		'set rsCntReport=nothing

		'攔停
		'strCntStop="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		'",BillBase f where a.BillSN=f.SN" &_
		'" and f.RecordStateID=0" &_
		'" and f.MemberStation='"&trim(StationArray(SA))&"'" &_
		'" and a.ExchangeTypeID='N' and a.DciReturnStatusID='n' and a.BillTypeID<>'2'" &_
		'" and a.ReturnMarkType='5'" &_
		'" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere
		'set rsCntStop=conn.execute(strCntStop)
		'if not rsCntStop.eof then
		'	CloseCnt3=CloseCnt3+cint(trim(rsCntStop("cnt")))
		'end if
		'rsCntStop.close
		'set rsCntStop=nothing
        '
		'if CloseCnt3>0 then
		'	response.write "結案 "&CloseCnt3&" 件"
		'else
		'	response.write "&nbsp;"
		'end if
			%> <!-- </span></td>-->
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
		
		</tr>
	</table>
	<center><%
	PageNum=1
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1

	%></center>
	<div class="PageNext"></div>

<%	CaseSn=0
	'台北市交通裁決所列表
	if instr(StationArrayTemp,"20")>0 or instr(StationArrayTemp,"21")>0 or instr(StationArrayTemp,"22")>0 or instr(StationArrayTemp,"23")>0 or instr(StationArrayTemp,"24")>0 or instr(StationArrayTemp,"25")>0 or instr(StationArrayTemp,"26")>0 then
	'逕舉 and 攔停
	PrintSN=0

	strSQL="select f.SN,f.BillNO,f.CarNO,e.Owner,e.OwnerZip,f.Rule1,f.Rule2,f.Rule3" &_
		" from (select a.BillNo,a.UserMarkDate from mailstationreturn a where " & strwhere & ") a " &_
        " ,BillBaseDCIReturn e,BillBase f,mailstationreturn g" &_
		" where a.Billno=f.Billno" &_
		" and f.RecordStateID=0 and f.Billno=g.Billno" &_
		" and a.BillNo=e.BillNO " &_
		" and ((f.BillTypeID='2' and 	e.DCIReturnStation in ('20','21','22','23','24','25','26') and e.ExchangeTypeID='W' and e.Status='Y')" &_
		" or (f.BillTypeID<>'2' and     f.MemberStation in ('20','21','22','23','24','25','26') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n')))" &_
		" order by g.UserMarkDate"

	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
	pagenum=1
%>
	<center><font size="3">寄存期滿清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%="台北市交通事件裁決所"%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
<%if sys_City="嘉義縣" then %>
			<td align="center">鄉鎮別</td>
<%end if%>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>
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
			<td nowrap><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write trim(rs1("Owner"))
			else
				response.write "&nbsp;"
			end if				
			%></td>
<%if sys_City="嘉義縣" then %>
			<td><%
			if trim(rs1("OwnerZip"))<>"" and not isnull(rs1("OwnerZip")) then
				strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					response.write trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			else
				response.write "&nbsp;"
			end if	
			%></td>
<%end if%>
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
	PrintSN=0
	CaseSn=0
	PageNum=1
	'高雄市交通事件裁決所列表
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then

	if sys_City="宜蘭縣" then 
		CaseSn=0
	end if

	
	'逕舉and 攔停
	PrintSN=0
	strSQL="select f.SN,f.BillNO,f.CarNO,e.Owner,e.OwnerZip,f.Rule1,f.Rule2,f.Rule3" &_
		" from (select a.BillNo,a.UserMarkDate from mailstationreturn a where " & strwhere & ") a " &_
        " ,BillBaseDCIReturn e,BillBase f,mailstationreturn g" &_
		" where a.Billno=f.Billno" &_
		" and f.RecordStateID=0 and f.Billno=g.Billno" &_
		" and a.BillNo=e.BillNO " &_
		" and ((f.BillTypeID='2' and e.DCIReturnStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status='Y')" &_
		" or (f.BillTypeID<>'2' and f.MemberStation in ('30','31','32') and e.ExchangeTypeID='W' and e.Status in ('Y','S','n')))" &_
		" order by g.UserMarkDate"

	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<center><font size="3">寄存期滿清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%="高雄市交通事件裁決所"%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
<%if sys_City="嘉義縣" then %>
			<td align="center">鄉鎮別</td>
<%end if%>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>

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
			<td nowrap><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write trim(rs1("Owner"))
			else
				response.write "&nbsp;"
			end if				
			%></td>
<%if sys_City="嘉義縣" then %>
			<td><%
			if trim(rs1("OwnerZip"))<>"" and not isnull(rs1("OwnerZip")) then
				strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					response.write trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			else
				response.write "&nbsp;"
			end if	
			%></td>
<%end if%>
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

	'其他監理所列表
	StationName=split(StationNameArray,",")
	for SA2=0 to ubound(StationArray)
		if instr("20,21,22,23,24,25,26,30,31,32",trim(StationArray(SA2)))<=0 then
			if sys_City="宜蘭縣" then 
				CaseSn=0
			end if
	'逕舉and 攔停
	PrintSN=0
	CaseSn=0
	pagenum=1
	strSQL="select f.SN,f.BillNO,f.CarNO,e.Owner,e.OwnerZip,f.Rule1,f.Rule2,f.Rule3" &_
		" from (select a.BillNo,a.UserMarkDate from mailstationreturn a where " & strwhere & ") a " &_
        " ,BillBaseDCIReturn e,BillBase f,mailstationreturn g" &_
		" where a.Billno=f.Billno" &_
		" and f.RecordStateID=0 and f.Billno=g.Billno" &_
		" and a.BillNo=e.BillNO " &_
		" and ((f.BillTypeID='2' and e.DCIReturnStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status='Y')" &_
		" or (f.BillTypeID<>'2' and f.MemberStation='"&trim(StationArray(SA2))&"' and e.ExchangeTypeID='W' and e.Status in ('Y','S','n')))" &_
		" order by g.UserMarkDate"

	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<center><font size="3">寄存期滿清冊</font></center>
	列印日期：<%=now%>
	<br>
	到案處所：<%="&nbsp;"&StationName(SA2)%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center">編號</td>
			<td align="center">違規單號</td>
			<td align="center">車號</td>
			<td align="center">車主姓名</td>
<%if sys_City="嘉義縣" then %>
			<td align="center">鄉鎮別</td>
<%end if%>
			<td align="center">法條一</td>
			<td align="center">法條二</td>
			<td align="center">法條三</td>

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
			<td nowrap><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write trim(rs1("Owner"))
			else
				response.write "&nbsp;"
			end if				
			%></td>
<%if sys_City="嘉義縣" then %>
			<td><%
			if trim(rs1("OwnerZip"))<>"" and not isnull(rs1("OwnerZip")) then
				strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					response.write trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			else
				response.write "&nbsp;"
			end if	
			%></td>
<%end if%>
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

<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
window.print();
<%elseif sys_City="宜蘭縣" then%>
printWindow(true,7,5.08,5.08,5.08);
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
<%conn.close%>