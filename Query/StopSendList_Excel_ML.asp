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
.style1 {font-family:新細明體; line-height:19px; font-size: 16px}
.style2 {font-family:新細明體; line-height:25px; font-size: 20px}
.style3 {font-family:新細明體; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體; line-height:18px; font-size: 12px}
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
	RecordDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
	RecordDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
	strwhere=" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	strwhere=strwhere&" and f.MemberStation='54'"

	If sys_City="苗栗縣" Then
		strB="select distinct(a.Batchnumber) from DCILog a" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f,UnitInfo b where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and b.UnitID=f.BillUnitID " & _
		" and f.RecordStateID=0 "&strwhere
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

	strStation="select distinct(b.UnitTypeID) from DCILog a" &_
		",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f,UnitInfo b where f.SN=a.BillSN" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and b.UnitID=f.BillUnitID " & _
		" and f.RecordStateID=0 "&strwhere&" order by UnitTypeID"
		'response.write strStation
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then
		rsStation.MoveFirst 
	else
		response.write "查無資料，請確認此批舉發單是攔停舉發單!"
	end if
	While Not rsStation.Eof
		if StationArrayTemp="" then
			StationArrayTemp=trim(rsStation("UnitTypeID"))
		else
			StationArrayTemp=StationArrayTemp&","&trim(rsStation("UnitTypeID"))
		end if
	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing

	strCnt="select count(*) as cnt from DCILog a,DCIReturnStatus d" &_
		",BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and f.RecordStateID=0 "&strwhere
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=nothing

	'response.write StationArrayTemp&"<br>"&DBcnt
%>
</head>
<body>
<table width="400" border="0" cellpadding="3" cellspacing="0" align="center">
<tr>
	<td colspan="2" align="center"><span class="style2"><strong>苗栗縣警察局</strong></span>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"  class="style2">查獲違反道路交通管理事件移送表(攔停)<br>&nbsp;
	</td>
</tr>
<tr>
	<td align="right" class="style2">移送單位：
	</td>
	<td ><span class="style2">苗栗監理站</span>
	</td>
</tr>
<tr>
	<td align="right" class="style2">收件入期：
	</td>
	<td class="style2"><%
	response.write Left(request("RecordDate1"),Len(Trim(request("RecordDate1")))-4)
	response.write "/"&mid(request("RecordDate1"),Len(Trim(request("RecordDate1")))-3,2)
	response.write "/"&right(request("RecordDate1"),2)
	%></td>
</tr>
</table>
<table width="400" border="1" cellpadding="3" cellspacing="0" align="center">
<tr>
	<td class="style2">分局名稱
	</td>
	<td align="center" class="style2">攔停件數
	</td>
</tr>
<%
'單位順序排序
UnitArrayOrder=Split(StationArrayTemp,",")
UnitArrayOrderTemp=""
strOrder="select * from UnitInfo where UnitID in ('"&replace(StationArrayTemp,",","','")&"') order by UnitOrder"
Set rsOrder=conn.execute(strOrder)
While Not rsOrder.Eof
	For j=0 To UBound(UnitArrayOrder)
		If Trim(UnitArrayOrder(j))=Trim(rsOrder("UnitID")) Then 
			if UnitArrayOrderTemp="" then
				UnitArrayOrderTemp=trim(rsOrder("UnitID"))
			else
				UnitArrayOrderTemp=UnitArrayOrderTemp&","&trim(rsOrder("UnitID"))
			end If
		End If 
	Next 
	rsOrder.MoveNext
Wend
rsOrder.close
Set rsOrder=Nothing

UnitArray=Split(UnitArrayOrderTemp,",")
UnitNameArrayTemp=""
UnitCnt=0
'全部單位都要列出
UnitCntArrayTemp="03BA,03B6,3N00,3O00,3P00,3R00,3Q00"
UnitCntArray=Split(UnitCntArrayTemp,",")
For i=0 To UBound(UnitCntArray)
%>
<tr>
	<td class="style2"><%
	strUN="select UnitName from UnitInfo where UnitID='"&Trim(UnitCntArray(i))&"'"
	Set rsUn=conn.execute(strUN)
	If Not rsUn.eof Then
		response.write Trim(rsUN("UnitName"))
		If UnitNameArrayTemp="" Then
			UnitNameArrayTemp=Trim(rsUN("UnitName"))
		Else
			UnitNameArrayTemp=UnitNameArrayTemp&"#@#"&Trim(rsUN("UnitName"))
		End If 
	End If
	rsUn.close 
	Set rsUn=Nothing 
	%></td>
	<td align="center" class="style2"><%
	strCnt="select count(*) as cnt from DCILog a,DCIReturnStatus d" &_
		",BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
		" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
		" and a.BillTypeID<>'2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and f.RecordStateID=0 "&strwhere&" and f.BillUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(UnitCntArray(i))&"')"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		UnitCnt=UnitCnt+CInt(rsCnt("Cnt"))
		response.write rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=Nothing
	
	%></td>
</tr>
<%
next
%>
<tr>
	<td align="center" class="style2">共計：</td>
	<td align="center" class="style2"><%=" &nbsp; &nbsp; "&UnitCnt&" 件"%></td>
</tr>
<tr>
	<td align="center" class="style2">&nbsp;</td>
	<td align="center" class="style2"><%=" &nbsp; &nbsp; &nbsp; &nbsp;包"%></td>
</tr>
</table>
<table width="600" border="0" cellpadding="3" cellspacing="0" align="center">
<tr>
	<td ><span class="style2">
	<br>
攔停資料請核對移送表與二、(三)聯違規通知單聯、扣件物，倘若不符請通知本局，逾三日未回覆視同相符無誤。<br><br><br>
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 移送單位： 苗栗監理站<br><br>
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 簽收人：
	</span>
	</td>
</tr>
</table>
<br><br><br><br><br>
<center>
<font class="style2">
移送車牌明細
</font>
</center>
<table width="500" border="1" cellpadding="3" cellspacing="0" align="center">
<tr>
	<td align="center" class="style2">分局</td>
	<td align="center" class="style2">車號</td>
	<td align="center" class="style2">備註</td>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
</table>
<center><%
	'response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	'PageNum=PageNum+1
%></center>
<!-- <div class="PageNext">&nbsp;</div>
 --><%
'列表===============================================
	PrintSNtotal=0	'編號
	
For i=0 To UBound(UnitArray)
	If Trim(UnitArray(i))="03BA" then
		PageCount=20
	Else
		PageCount=20
	End If 
	PrintSN=0
		strCnt="select count(*) as cnt from DCILog a,MemberData b" &_
			",DCIReturnStatus d, BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN" &_
			" and a.BillNo=e.BillNO and f.MemberStation='54'" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere &_
			" and f.BillUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(UnitArray(i))&"')"
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
			" and a.BillNo=e.BillNO and f.MemberStation='54'" &_
			" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
			" and a.DCIReturnStatusID=e.Status and e.ExchangeTypeID='W'" &_
			" and a.DCIReturnStatusID in ('Y','S','n','L')" &_
			" and a.BillTypeID<>'2'" &_
			" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
			" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0 "&strwhere &_
			" and f.BillUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(UnitArray(i))&"')" &_
			" order by f.RecordMemberID,f.RecordDate"
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then
%>
		<center><%
	If PageNum>1 then
		response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	End if
	PageNum=PageNum+1
	%></center>
<%		end if
			response.write "<div class=""PageNext""></div>"
%>
	<table width="710" border="0" align="center" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center"><font size="3"><b><%=TitleUnitName%></b>&nbsp;&nbsp;攔停舉發移送清冊</font></td>
		</tr>
		<tr>
			<td align="left">站所：<%
		response.write "54 苗栗監理站"
	%>&nbsp; &nbsp; &nbsp; &nbsp;<%
		response.write "移送日期"
	%>：<%=Right("000"&year(now)-1911,3)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;<%
	if sys_City="苗栗縣" then
		strUN="select UnitName from UnitInfo where UnitID='"&Trim(UnitArray(i))&"'"
		Set rsUn=conn.execute(strUN)
		If Not rsUn.eof Then
			response.write "分局名稱："&Trim(rsUN("UnitName"))&"&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;"
		End If
		rsUn.close 
		Set rsUn=Nothing 
	Else
		response.write "(本批案件已透過中華電信數據分公司作入案管制)"
	End if
	%>&nbsp; &nbsp; &nbsp;Page <%=fix(PrintSN/PageCount)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" align="center" cellpadding="1" cellspacing="0">
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
<%		for Q=1 to PageCount
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
			<td width="11%"><span class="style4"><%
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
			<td><span class="style4"><%
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
	<table width="710" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr><td>
		共計： <%=PrintSN%>  &nbsp;筆<br>
		攔停資料請核對移送表與二、(三)聯違規通知單聯、扣件物，倘若不符請通知本局，逾三日未回覆視同相符無誤。
		<br>
		</td></tr>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%if i<>ubound(UnitArray) then%>
	<div class="PageNext"></div>
<%end if
Next 
%>
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