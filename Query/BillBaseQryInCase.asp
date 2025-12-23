

<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

.style4 {
	color: #FF0000;
	font-size: 16px
	}

-->
</style>
<title>舉發單入案</title>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<!--#include file="../Common/Banner.asp"-->
<% Server.ScriptTimeout = 800 %>
<%
'權限
'AuthorityCheck(250)
strAuthority = GenPkiTicket
session("Ticket") = strAuthority
RecordDate=split(gInitDT(date),"-")

'抓縣市
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

'組成查詢SQL字串
if request("DB_Selt")="Selt" then
		strwhere=""
		if trim(request("RecordDateCheck"))="1" then
			if request("RecordDate")<>"" and request("RecordDate1")<>""then
				RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
				RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"
				if strwhere<>"" then
					strwhere=strwhere&" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				else
					strwhere=" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				end if
			end if
		end if
		if trim(request("RecordDate_h"))<>"" or trim(request("RecordDate1_h"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			else
				strwhere=" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			end if
		end if
		if request("Sys_BillUnitID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillUnitID ="&request("Sys_BillUnitID")
			else
				strwhere=" and a.BillUnitID="&request("Sys_BillUnitID")
			end if
		end if
		if request("Sys_BillMem")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and (a.BillMem1='"&request("Sys_BillMem")&"' or a.BillMem2='"&request("Sys_BillMem")&"' or a.BillMem3='"&request("Sys_BillMem")&"')"
			else
				strwhere=" and (a.BillMem1='"&request("Sys_BillMem")&"' or a.BillMem2='"&request("Sys_BillMem")&"' or a.BillMem3='"&request("Sys_BillMem")&"')"
			end if
		end if
		if request("Sys_RecordUnit")<>"" and request("Sys_RecordMemberID")="" then
			strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
		end if
		if request("Sys_RecordMemberID")<>"" then
			strwhere=strwhere&" and a.RecordMemberID ="&request("Sys_RecordMemberID")
		end if
		if request("billtype")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillTypeID='"&request("billtype")&"'"
			else
				strwhere=" and a.BillTypeID='"&request("billtype")&"'"
			end if
		end if
		if request("Sys_BillNo")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillNo='"&request("Sys_BillNo")&"'"
			else
				strwhere=" and a.BillNo='"&request("Sys_BillNo")&"'"
			end if
		end if
		if request("Sys_CarNo")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.CarNo like '%"&request("Sys_CarNo")&"%'"
			else
				strwhere=" and a.CarNo like '%"&request("Sys_CarNo")&"%'"
			end if
		end if
		if request("Sys_Driver")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.Driver='"&request("Sys_Driver")&"'"
			else
				strwhere=" and a.Driver='"&request("Sys_Driver")&"'"
			end if
		end if
		if request("Sys_DriverID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.DriverID='"&request("Sys_DriverID")&"'"
			else
				strwhere=" and a.DriverID='"&request("Sys_DriverID")&"'"
			end if
		end if
		
		if trim(request("DCIstatus"))="0" then
			if trim(request("sys_BatcuNumber"))<>"" then
				if sys_City="彰化縣" or sys_City="高雄縣" then
					if strwhere<>"" then
						strwhere=strwhere&" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID is not null and BatchNumber='"&trim(request("sys_BatcuNumber"))&"')"
					else
						strwhere=" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID is not null and BatchNumber='"&trim(request("sys_BatcuNumber"))&"')"
					end if
				else
					if strwhere<>"" then
						strwhere=strwhere&" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S' and BatchNumber='"&trim(request("sys_BatcuNumber"))&"')"
					else
						strwhere=" and a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S' and BatchNumber='"&trim(request("sys_BatcuNumber"))&"')"
					end if
				end if
			Else
				if sys_City="台南市x" Then
					If trim(request("billtype"))="2" Then
						If trim(request("BillUseTool"))="0" or trim(request("BillUseTool"))="All" then
							strwhere=strwhere&" and ((a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S')))"
						Else
							strwhere=strwhere&" and ((a.BillTypeID<>'2' and a.BillStatus='0') or (a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S')) or (a.BillTypeID='2' and a.BillStatus='0'))"
						End If	
					Else
						strwhere=strwhere&" and ((a.BillTypeID<>'2' and a.BillStatus='0') or (a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S')) or (a.BillTypeID='2' and a.BillStatus='0'))"
					End If 
				Else
					strwhere=strwhere&" and ((a.BillTypeID<>'2' and a.BillStatus='0') or (a.BillStatus='1' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='A' and DciReturnStatusID='S')) or (a.BillTypeID='2' and a.BillStatus='0'))"
				End If 
								
			end if
		elseif trim(request("DCIstatus"))="1" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='2' and a.Sn in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='W' and (DciReturnStatusID in ('Y','S','n') or DciReturnStatusID is null)))"
			else
				strwhere=" and a.BillStatus='2' and a.Sn in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='W' and (DciReturnStatusID in ('Y','S','n') or DciReturnStatusID is null)))"
			end if			
		end if
		if trim(request("billtype"))="2" then
			if trim(request("BillUseTool"))="All" then
				strwhere=strwhere
			elseif trim(request("BillUseTool"))="0" then
				strwhere=strwhere&" and a.UseTool<>8"
			else
				strwhere=strwhere&" and a.UseTool=8"
			end if
		end if
		if strwhere<>"" then
			strwhere=strwhere&" and a.RecordStateID=0"
		else
			strwhere=" and a.RecordStateID=0"
		end if

		'是否要判斷一打一驗 1:是 0:否
		if Session("DoubleCheck")="1" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.DoubleCheckStatus=1"
			else
				strwhere=" and a.DoubleCheckStatus=1"
			end if
		end if
		
		if trim(request("billtype"))="1" then
			strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillFillDate,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.papercheck=1 and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.RecordDate"
		else
			strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillFillDate,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.RecordDate"
		end if
	'response.write strSQL
end if

'入案(遇到RecordStateID=-1不做)
if trim(request("kinds"))="BillToDCILog" then
	'批號
	strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theBatchTime=(year(now)-1911)&"W"&trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=nothing
	
	if sys_City="基隆市" Then
		if trim(request("UploadNote"))<>"" then
			strInsP="Insert Into BillCaseInNote(SN,BatchNumber,Note,RecordMemberID,RecordDate) " &_
				" values((select nvl(max(SN),0)+1 from BillCaseInNote),'"&trim(theBatchTime)&"'" &_
				",'"&trim(request("UploadNote"))&"'," & Trim(Session("User_ID")) &_
				",sysdate" &_
				")"
			'response.write strInsP
			'response.end
			conn.execute strInsP
		end if
	End If 

	if sys_City="苗栗縣" Or (sys_City="高雄市" And trim(request("billtype"))="2") Then
		'車籍查尋的批號
		strSN2="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN2=conn.execute(strSN2)
		if not rsSN2.eof then
			theBatchTimeQryCar=(year(now)-1911)&"A"&trim(rsSN2("SN"))
		end if
		rsSN2.close
		set rsSN2=nothing
	End If

	if trim(request("HelpPrint"))="1" then
		strInsP="Insert Into BillPrintJob(BatchNumber) " &_
			" values('"&trim(theBatchTime)&"')"
		conn.execute strInsP
	end if

	strToDCI="select a.papercheck, a.SN,a.IllegalDate,a.BillTypeID,a.BillNo,a.CarNo,a.BillUnitID,a.BillStatus,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.papercheck=1 and a.RecordStateID<>-1 and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.RecordDate"
	set rsToDCI=conn.execute(strToDCI)
	If Not rsToDCI.Bof Then
		rsToDCI.MoveFirst
	else
%>
<script language="JavaScript">
	alert("無可進行入案之舉發單！");
</script>
<%
	end if
	While Not rsToDCI.Eof
		if sys_City="苗栗縣" Or (sys_City="高雄市" And trim(request("billtype"))="2") Then '車籍查尋
			funcCarDataCheck conn,trim(rsToDCI("SN")),"",trim(rsToDCI("BillTypeID")),trim(rsToDCI("CarNo")),trim(rsToDCI("BillUnitID")),trim(rsToDCI("RecordDate")),trim(rsToDCI("RecordMemberID")),theBatchTimeQryCar
		End If

		funcBillToDCICaseIn conn,trim(rsToDCI("SN")),trim(rsToDCI("BillNo")),trim(rsToDCI("BillTypeID")),trim(rsToDCI("CarNo")),trim(rsToDCI("BillUnitID")),trim(rsToDCI("RecordDate")),trim(rsToDCI("RecordMemberID")),theBatchTime,sys_City
	rsToDCI.MoveNext
	Wend
	If Not rsToDCI.Bof Then
%>
<script language="JavaScript">
	alert("入案處理完成，批號：<%=theBatchTime%>");
</script>
<%
	end if
	rsToDCI.close
	set rsToDCI=nothing
end if

'做完車籍查詢及入案等動作後再查詢告發單，讓列表取得的資料為最新
if request("DB_Selt")="Selt" then
'response.write strSQL
'response.end
		set rsfound=conn.execute(strSQL)
		if trim(request("billtype"))="1" then
			strCnt="select count(*) as cnt from BillBase a,MemberData b where a.papercheck=1 and a.RecordMemberID=b.MemberID(+)"&strwhere
		else
			strCnt="select count(*) as cnt from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close
		tmpSQL=strwhere
		'Session.Contents.Remove("BillSQL")
		'Session("BillSQL")=strSQL
		Session.Contents.Remove("PrintCarDataSQL")
		Session("PrintCarDataSQL")=strwhere
		Session.Contents.Remove("BillSQLforReport")
		Session("BillSQLforReport")=strwhere
end if
%>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">舉發單<%
		if trim(request("billtype"))="1" then
			response.write "攔停"
		elseif trim(request("billtype"))="2" then
			response.write "逕舉"
		end if
		%>入案&nbsp;&nbsp;
		<%if trim(request("billtype"))="2" then%>
		<span class="style4"><strong>(逕舉手開單入案，監理站回傳後，請確認自動帶回之應到案處所是否與舉發單上相同)</strong></span>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<input type="hidden" name="billtype" value="<%=trim(request("billtype"))%>">
						<input type="hidden" name="HelpPrint" value="<%=trim(request("HelpPrint"))%>">
						<input type="hidden" name="RecordDateCheck" value="1" >
						建檔日期
						<input name="RecordDate" type="text" value="<%
							if trim(request("DB_Selt"))="" Then
								if sys_City="台東縣" Then
									RecordDateTmp=Year(DateAdd("d",-60,now))-1911&Right("00" & Month(DateAdd("d",-60,now)),2)&Right("00" & Day(DateAdd("d",-60,now)),2)
								else
									RecordDateTmp=Year(DateAdd("d",-15,now))-1911&Right("00" & Month(DateAdd("d",-15,now)),2)&Right("00" & Day(DateAdd("d",-15,now)),2)
								End If 
							else
								RecordDateTmp=trim(request("RecordDate"))
							end if
							response.write RecordDateTmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
						~
						<input name="RecordDate1" type="text" value="<%
							if trim(request("DB_Selt"))="" then
								RecordDate1Tmp=ginitdt(now)
							else
								RecordDate1Tmp=trim(request("RecordDate1"))
							end if
							response.write RecordDate1Tmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
						
						
						<!--時段-->
						<input name="RecordDate_h" type="hidden" value="<%=request("RecordDate_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')"> <!-- 時 ~ -->
						<input name="RecordDate1_h" type="hidden" value="<%=request("RecordDate1_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')"><!-- 時 -->
						
						<img src="space.gif" width="2" height="10">
						DCI作業
						<select name="DCIstatus">
							<option value="0" <%
							if trim(request("DCIstatus"))="0" then response.write "selected"
							%>>入案</option>
							<!-- <option value="1" <%
							if trim(request("DCIstatus"))="1" then response.write "selected"
							%>>入案異常</option> -->
						</select>

						<img src="space.gif" width="8" height="10">
				<%if trim(request("billtype"))="2" then%>
						舉發單類別
						<select name="BillUseTool">
							<option value="0" <%
							if trim(request("BillUseTool"))="0" then response.write "selected"
							%>>逕舉</option>
							<option value="1" <%
							if trim(request("BillUseTool"))="1" then response.write "selected"
							%>>逕舉手開單</option>
							<option value="All" <%
							if trim(request("BillUseTool"))="All" then response.write "selected"
							%>>全部</option>
						</select>
						可選擇針對 "逕舉手開單" 入案
				<%end if%>
						
						<br>
						建檔單位
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						建檔人
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						
						<img src="space.gif" width="8" height="10">
				<%
				if trim(request("billtype"))="1" then
				%>
						<input type="hidden" name="sys_BatcuNumber" size="8" value="<%=trim(request("sys_BatcuNumber"))%>">
				<%
				elseif trim(request("billtype"))="2" then
				%>
						<font color="red"><strong>車籍查詢批號</strong></font>
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							strSQL1="select distinct TO_char(ExchangeDate,'YYYY/MM/DD') ExchangeDate,BatchNumber from DCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",-5, date)&" 00:00"&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID='A' order by ExchangeDate DESC"
		
							set rs=conn.execute(strSQL1)
							cut=0
							while Not rs.eof
								ExchangeDate=gInitDT(trim(rs("ExchangeDate")))
								response.write "<option value="""&trim(rs("BatchNumber"))&""">"
								response.write ExchangeDate& " - "&cut&"　"&trim(rs("BatchNumber"))
								response.write "</option>"
								cut=cut+1
								rs.movenext
							wend
							rs.close
						%>
						</select>
						<input type="text" name="sys_BatcuNumber" size="8" value="<%=trim(request("sys_BatcuNumber"))%>" onkeyup="value=value.toUpperCase()">
						車號
						<input type="text" name="Sys_CarNo" size="8" maxlength="9" value="<%=trim(request("Sys_CarNo"))%>" onkeyup="value=value.toUpperCase()">
				<%
				end if
				%>
						
						<img src="space.gif" width="8" height="10">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" <%
					if trim(request("billtype"))="2" then
						if CheckPermission(255,1)=false then
							response.write "disabled"
						end if
					else
						if CheckPermission(250,1)=false then
							response.write "disabled"
						end if
					end if
						%>>
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseQryInCase.asp?billtype=<%=trim(request("billtype"))%>'"> 
<%
if sys_City<>"高雄市" And sys_City<>ApconfigureCityName then
	if trim(request("billtype"))="2" then
		if trim(Session("SpecUser"))="1" then
			if sys_City<>"花蓮縣" then
	%>
				<input type="button" name="cancel" value="入案前特殊車輛比對" onClick="funChkVIP();"> 
	<%
			else
	%>
				<input type="button" name="cancel" value="入案前特殊車輛比對" onClick="funChkVIP_HL();"> 
	<%
			end if
		end if
	end if
end if
%>
					<br /><span class="style4"><strong>※上傳入案兩小時後，請確認案件是否都有正常入案</strong></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF" class="style3">
			舉發單紀錄列表
			<img src="space.gif" width="56" height="8">
			每頁 
			<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>
			筆 <font color="#F90000"><strong>(共 <%=DBsum%> 筆)</strong></font>
			&nbsp;&nbsp;<%
		if sys_City="高雄縣" or sys_City="花蓮縣" or sys_City="高雄市" or sys_City="彰化縣" Or sys_City=ApconfigureCityName then
			if trim(request("billtype"))="2" then
				if sys_City="高雄市" then
					'thirdNo=right(year(date)-1911,1)
					thirdNo="H"
					thirdNoSubUnit="D"
					'response.write thirdNo
					if trim(Session("Credit_ID"))="T220933992" then
						UserStartNo="BD"&thirdNo&"0"
						UserSeq="bill0807BD0"
					elseif trim(Session("Credit_ID"))="E223625931" then
						UserStartNo="BD"&thirdNo&"3"
						UserSeq="bill0807BD3"
					elseif trim(Session("Credit_ID"))="E121011955" then
						UserStartNo="BD"&thirdNo&"5"
						UserSeq="bill0807BD5"
					elseif trim(Session("Credit_ID"))="E120003931" then
						UserStartNo="BB"&thirdNo&"0"
						UserSeq="bill0807BB0"
					elseif trim(Session("Credit_ID"))="T220359567" then
						UserStartNo="BB"&thirdNo&"3"
						UserSeq="bill0807BB3"
					elseif trim(Session("Credit_ID"))="E220912204" then
						UserStartNo="BB"&thirdNo&"5"
						UserSeq="bill0807BB5"					
					elseif trim(Session("Credit_ID"))="T220988040" then
						UserStartNo="BB"&thirdNo&"7"
						UserSeq="bill0807BB7"
					elseif trim(Session("Credit_ID"))="S220060233" then
						UserStartNo="BB"&thirdNo&"9"
						UserSeq="bill0807BB9"
					elseif trim(Session("Credit_ID"))="E220182233" then
						UserStartNo="BC"&thirdNo&"0"
						UserSeq="bill0807BC0"
					elseif trim(Session("Credit_ID"))="T121457177" then
						UserStartNo="BC"&thirdNo&"3"
						UserSeq="bill0807BC3"
					elseif trim(Session("Credit_ID"))="R120228634" Or trim(Session("Credit_ID"))="E221201933" Or trim(Session("Credit_ID"))="S221552347" then
						UserStartNo="BC"&thirdNo&"5"
						UserSeq="bill0807BC5"
					else
						if trim(Session("Unit_ID"))="0807" then
							sysUserUnit="0807"
						else
							strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUserUnit=conn.execute(strUserUnit)
							if not rsUserUnit.eof then
								sysUserUnit=trim(rsUserUnit("UnitTypeID"))
							end if
							rsUserUnit.close
							set rsUserUnit=nothing
						end if
						'還有進入入案鈕的disabled
						UserStartNo=""
						UserSeq=""
						'抓 起始碼 跟 Seq Name
						strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
						set rsSeq=conn.execute(strSeq)
						if not rsSeq.eof Then
							If Len(trim(rsSeq("BillStartVocab")))=3 then
								UserStartNo=trim(rsSeq("BillStartVocab"))
							Else
								UserStartNo=trim(rsSeq("BillStartVocab"))&thirdNoSubUnit
							End If 
							UserSeq=trim(rsSeq("SeqNoName"))
						end if
						rsSeq.close
						set rsSeq=nothing	
					end If
				elseif sys_City="花蓮縣" Then
					UserStartNo=""
					UserSeq=""
					if trim(Session("Credit_ID"))="A06" Or trim(Session("Credit_ID"))="A07" Then
						UserStartNo="PB"
						UserSeq="billA06PB"
					Else
						strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
						set rsUserUnit=conn.execute(strUserUnit)
						if not rsUserUnit.eof then
							sysUserUnit=trim(rsUserUnit("UnitTypeID"))
						end if
						rsUserUnit.close
						set rsUserUnit=Nothing

						
						'抓 起始碼 跟 Seq Name
						strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
						set rsSeq=conn.execute(strSeq)
						if not rsSeq.eof then
							UserStartNo=trim(rsSeq("BillStartVocab"))
							UserSeq=trim(rsSeq("SeqNoName"))
						end if
						rsSeq.close
						set rsSeq=nothing	
					End if
				else
					if sys_City="高雄縣" then
						if Session("Unit_ID")="8H00" then
							sysUserUnit="8H00"
						else
							strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUserUnit=conn.execute(strUserUnit)
							if not rsUserUnit.eof then
								sysUserUnit=trim(rsUserUnit("UnitTypeID"))
							end if
							rsUserUnit.close
							set rsUserUnit=nothing
						end if
					ElseIf sys_City="彰化縣" then
						strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
						set rsUserUnit=conn.execute(strUserUnit)
						if not rsUserUnit.eof then
							sysUserUnit=trim(rsUserUnit("UnitTypeID"))
						end if
						rsUserUnit.close
						set rsUserUnit=Nothing
					elseif sys_City=ApconfigureCityName then
	
							strUserUnit="select UnitID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUserUnit=conn.execute(strUserUnit)
							if not rsUserUnit.eof then
								sysUserUnit=trim(rsUserUnit("UnitID"))
							end if
							rsUserUnit.close
							set rsUserUnit=nothing

					end if
					'還有進入入案鈕的disabled
					UserStartNo=""
					UserSeq=""
					'抓 起始碼 跟 Seq Name
					strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
					set rsSeq=conn.execute(strSeq)
					if not rsSeq.eof then
						UserStartNo=trim(rsSeq("BillStartVocab"))
						UserSeq=trim(rsSeq("SeqNoName"))
					end if
					rsSeq.close
					set rsSeq=nothing	
				end if
				
				stringAlert=""
				if sys_City="高雄市" Then
					If Len(UserStartNo)=4 Then
						stringAlert="  單號後五碼用到 90000 時，請提早通知工程師處理，以免重複取號"
					ElseIf Len(UserStartNo)=3 Then
						stringAlert="  單號後六碼用到 900000 時，請提早通知工程師處理，以免重複取號"
					End If 
				End If 

				response.write "<font color=""#0066FF""><strong><font size=""5"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;舉發單開頭碼&nbsp;"&UserStartNo&stringAlert&"</font></strong></font>"

			
			end if
		end if
			%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th width="8%">違規日期</th>
					<th width="8%">舉發員警</th>
				<%if trim(request("billtype"))="1" then%>
					<th width="6%">舉發單號</th>
				<%end if%>
					<th width="5%">車號</th>
					<th width="6%">車種</th>
					<th width="4%">類別</th>
				<%if trim(request("billtype"))="1" then%>
					<th width="6%">駕駛人</th>
				<%end if%>
					<th width="10%">法條</th>
					<th width="8%">DCI</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
				chkCaseInDelayFlag=0
				CaseInDelayBillNo=""
				if request("DB_Selt")="Selt" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						chname="":chRule="":ForFeit=""
						if rsfound("BillMem1")<>"" then	chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then	chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then	chname=chname&"/"&rsfound("BillMem3")
						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='30'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
						response.write "<td width='8%'>"&chname&"</td>"
'					if trim(rsfound("BillTypeID"))="2" then
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("BillNo")&"</a></td>"
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					else
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("BillNo")&"</a></td>"
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					end if
					if trim(request("billtype"))="1" then
						response.write "<td width='6%'>"&rsfound("BillNo")&"</td>"
					end if
						response.write "<td width='6%'>"&rsfound("CarNo")&"</td>"
						response.write "<td width='5%'>"
							if trim(rsfound("CarSimpleID"))="1" then
								response.write "汽車"
							elseif trim(rsfound("CarSimpleID"))="2" then
								response.write "拖車"
							elseif trim(rsfound("CarSimpleID"))="3" then
								response.write "重機"
							elseif trim(rsfound("CarSimpleID"))="4" then
								response.write "輕機"
							elseif trim(rsfound("CarSimpleID"))="5" then
								response.write "動力機械"
							elseif trim(rsfound("CarSimpleID"))="6" then
								response.write "臨時車牌"
							end If
						
						response.write "</td>"
						response.write "<td width='4%'>"
					strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
					set rsBTypeVal=conn.execute(strBTypeVal)
					if not rsBTypeVal.eof then
						response.write rsBTypeVal("Content")
					end if
					rsBTypeVal.close
					set rsBTypeVal=nothing
						response.write "</td>"
					if trim(request("billtype"))="1" then
						response.write "<td width='6%'>"&rsfound("Driver")&"</td>"
					end if
						response.write "<td width='10%'>"&chRule&"</td>"
						response.write "<td width='8%'>"
						if trim(rsfound("BillStatus"))="0" then
							response.write "<font color='#999999'>未處理</font>"
						elseif trim(rsfound("BillStatus"))="1" then
							response.write "<font color='#FF66CC'>車籍查詢</font>"
						elseif trim(rsfound("BillStatus"))="2" then
							response.write "<font color='#009900'>入案</font>"
						elseif trim(rsfound("BillStatus"))="3" then
							response.write "<font color='#0000FF'>退件</font>"
						elseif trim(rsfound("BillStatus"))="4" then
							response.write "<font color='#0000FF'>寄存</font>"
						elseif trim(rsfound("BillStatus"))="5" then
							response.write "<font color='#0000FF'>公示</font>"
						elseif trim(rsfound("BillStatus"))="6" then
							response.write "<font color='#FF0000'>刪除</font>"
						end if
						response.write "</td>"
						response.write "</tr>"
						If sys_City="基隆市" then
							If DateDiff("d",trim(rsfound("BillFillDate")),now)>=4 Then
								chkCaseInDelayFlag=1
								If CaseInDelayBillNo<>"" Then
									CaseInDelayBillNo=CaseInDelayBillNo&","&trim(rsfound("BillNo"))
								Else
									CaseInDelayBillNo=CaseInDelayBillNo&trim(rsfound("BillNo"))
								End If 
								
							end If 
						End If 
						rsfound.movenext
					next
				end if
				%>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="Submit4242" value="進行入案" <%
			if request("billtype")<>"2" then%>
				onclick="BillToDCILog()"
			<%else%>
				onclick="BillFillDate_Update()"
			<%end if
			if ((sys_City="高雄縣" and sysUserUnit<>"8J00") or sys_City="花蓮縣") and UserStartNo="" And Trim(request("BillUseTool"))<>"1" and trim(request("billtype"))="2" then
				response.write " disabled"
			end if
			%>>
			<!-- <input type="button" name="b12" value="自然人憑證進行入案" onclick="SignICCard()">-->
			<span class="style3"><img src="space.gif" width="5" height="8"></span>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="hidden" name="DelReason" value="">
		<%if sys_City="基隆市" Then%>
			<%if request("billtype")<>"2" then%>
			<br>
			備註 <input type="text" name="UploadNote" value="<%=CaseInDelayBillNo%>" style="width:800px;"> 
			<%End If %>
		<%End If %>
		</td>
	</tr>
	<tr>
		<td>
			<p align="center">&nbsp;</p>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="PKICarchk" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&trim(request("Sys_RecordMemberID"))&"');"%>
	function funChkVIP(){
		var Billsum="<%=DBsum%>";
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲入案的舉發單！");
		}else if (Billsum=="0"){
			alert("查無可入案之舉發單！");
		}else{
			window.open("ChkSpecCar.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
		}
	}
	function funChkVIP_HL(){
		var Billsum="<%=DBsum%>";
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲入案的舉發單！");
		}else if (Billsum=="0"){
			alert("查無可入案之舉發單！");
		}else{
			window.open("ChkSpecCar_HL.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
		}
	}
	function funSelt(){
		var error=0;
		var errorString="";
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate.value.substr(0,1)=="9" && myForm.RecordDate.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate.value.substr(0,1)=="1" && myForm.RecordDate.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate1.value.substr(0,1)=="9" && myForm.RecordDate1.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}else if( myForm.RecordDate1.value.substr(0,1)=="1" && myForm.RecordDate1.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate_h.value!="" || myForm.RecordDate1_h.value!=""){
			if(myForm.RecordDate_h.value=="" || myForm.RecordDate1_h.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔時段輸入不完整!!";
			}
		}
		if (error>0){
			alert(errorString);
		}else{
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}

	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
		win.focus();
		return win;
	}
	function repage(){
		myForm.DB_Move.value=0;
		myForm.submit();
	}
	function funchgExecel(){
		UrlStr="BillBaseQry_Execel.asp?WorkType=1";
		newWin(UrlStr,"inputWin",790,550,50,10,"yes","yes","yes","no");
	}
	//列印車籍清冊
	function funchgCarDataList(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲列印車籍清冊的舉發單！");
		}else{
			UrlStr="PrintCarDataList.asp?dcitype=<%=trim(request("dcitype"))%>";
			newWin(UrlStr,"CarListWin",790,575,50,10,"yes","no","yes","no");
			myForm.action="PrintCarDataList.asp";
			myForm.target="CarListWin";
			myForm.submit();
			myForm.action="";
			myForm.target="";
		}
	}
	function funDbMove(MoveCnt){
		if (eval(MoveCnt)>0){
			if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
				myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
				myForm.submit();
			}
		}else{
			if (eval(myForm.DB_Move.value)>0){
				myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
				myForm.submit();
			}
		}
	}
	//入案
	function BillToDCILog(){
		var Billsum="<%=DBsum%>";
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲入案的舉發單！");
		}else if (Billsum=="0"){
			alert("查無可入案之舉發單！");
	<%If sys_City="基隆市" then%>
		<%if chkCaseInDelayFlag=1 then%>
			<%if request("billtype")<>"2" then%>
		}else if (myForm.UploadNote.value==""){
			alert("此批案件中，有填單日距上傳入案日超過 4 天之案件，請先於備註欄位輸入原因後，再進行入案！");
			<%End If%> 
		<%End If%> 
	<%End If%> 
		}else{
			if (myForm.Sys_RecordMemberID.value==""){
				if(confirm('您選擇將所有建檔人的舉發單入案，是否確定要入案？')){
					myForm.kinds.value="BillToDCILog";
					myForm.submit();
				}
			}else{
				if(confirm('確定要入案到監理所？')){
					myForm.kinds.value="BillToDCILog";
					myForm.submit();
				}
			}
		}
	}
	//逕舉入案
	function BillFillDate_Update(){
<%	'檢查是否有逕舉手開單
	if trim(request("billtype"))="2" then
		if trim(request("DB_Selt"))<>"" then
			strChk="select count(*) as cnt from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+) and a.UseTool=8 "&strwhere
			set rsChk=conn.execute(strChk)
			if not rsChk.eof then
				if trim(rsChk("cnt"))>0 then
					chkUseTool=1
				else
					chkUseTool=0
				end if
			end if
			response.write chkUseTool
			rsChk.close
			set rsChk=nothing
		else
			chkUseTool=0
		end if 
	else
		chkUseTool=0
	end if
%>
		var Billsum="<%=DBsum%>";
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲入案的舉發單！");
		}else if (Billsum=="0"){
			alert("查無可入案之舉發單！");
		}else{
			if (myForm.Sys_RecordMemberID.value==""){
			if(confirm('您選擇將所有建檔人的舉發單入案，是否確定要入案？')){
		<%if chkUseTool=0 then
			if trim(request("HelpPrint"))="1" then%>
				window.open("BillFillDate_Update.asp?HelpPrint=1&DciLogSQLforReport=<%=replace(strwhere,"%","@!@")%>","Report_CaseIn","width=520,height=200,left=300,top=150,scrollbars=yes,menubar=no,resizable=no,status=yes");
		<%	else%>
				window.open("BillFillDate_Update.asp?DciLogSQLforReport=<%=replace(strwhere,"%","@!@")%>","Report_CaseIn","width=520,height=200,left=300,top=150,scrollbars=yes,menubar=no,resizable=no,status=yes");
		<%	end if
		else%>
				myForm.kinds.value="BillToDCILog";
				myForm.submit();
		<%end if%>
			}
			}else{
		<%if chkUseTool=0 then
			if trim(request("HelpPrint"))="1" then	%>
				window.open("BillFillDate_Update.asp?HelpPrint=1&DciLogSQLforReport=<%=replace(strwhere,"%","@!@")%>","Report_CaseIn","width=520,height=200,left=300,top=150,scrollbars=yes,menubar=no,resizable=no,status=yes");
		<%	else %>
				window.open("BillFillDate_Update.asp?DciLogSQLforReport=<%=replace(strwhere,"%","@!@")%>","Report_CaseIn","width=520,height=200,left=300,top=150,scrollbars=yes,menubar=no,resizable=no,status=yes");
		<%	end if
		else%>
				BillToDCILog();
		<%end if%>
			}
		}
	}
	function chkPKI(){
		runServerScript("chkPKI.asp?PKICarchk="+myForm.PKICarchk.value);
	}
	function fnBatchNumber(){
		myForm.sys_BatcuNumber.value=myForm.Selt_BatchNumber.value;
	}

	function KeyDown(){ 

		if (event.keyCode==116){	//F5鎖死
			event.keyCode=0;   
			event.returnValue=false;   
		}
	}
<%if trim(request("DB_Selt"))="" then%>
	//funSelt();
<%end if%>
</script>
<script language="VBScript">
'Sub SignICCard()
'	Set atxCert = createobject("AresPKIAtx.AtxCertificate")
'	Set atxUtility = createobject("AresPKIAtx.AtxUtility")
'	Set cms = createobject("AresPKIAtx.AtxCmsSignedData")
'	AresPKIClient.setLicense "9a6d220031dad1702592b900e497c95df5c864d0848b0e26659fc6721d4979b62373ccfb46ed64a7e7ebfa6f80e9a498d1f70268e58d39042bb0282861b991ac8ad0a331a241f450b1cc0f8c270335a1e97f115a834ac5ba455095e38cb318dffa6e1db9662e22406ec1dc7aa3770d4adb091798170f1dc380fc3d7783de375a","聯宏科技股份有限公司"
'    ticket = "<%=strAuthority%>"
'    nRet = AresPKIClient.Init()    
'    if nRet <> 0 then
'		msgbox(AresPKIClient.GetErrorMessage())
'		AresPKIClient.Finalize()
'		exit sub
'    end if
'	nRet = AresPKIClient.EncodeP7SignedData(ticket,"",pSignedData,"")
'    if nRet <> 0 then
'		msgbox(AresPKIClient.GetErrorMessage())
'		AresPKIClient.Finalize()
'		exit sub
'    end if     
'    AresPKIClient.Finalize()
'        
'    If pSignedData <> "" Then
'        pSignedData = AresPKIClient.HexStringToB64(pSignedData)
'	    'msgbox pSignedData
'    End If
'
'	encodeData = atxUtility.BSTR_B64ToBin(atxUtility.BSTR_WideCharToMultiByte(pSignedData))
'	result = cms.InitDecode(encodeData,"")
'	binary=cms.Decode()
'	'For j = lbound(binary) To ubound(binary)
'		'msgbox(binary(j))
'	'Next
'
'	'取得原始資料
'	'msgbox "GetDecodeContent()=================================="
'	'msgbox(atxUtility.BSTR_MultiByteToWideChar(cms.GetDecodeContent()))
'
'	'取得憑證
'	certs=cms.GetDecodeCertificates()
'	cert = certs(0)
'	cms.FinalDecode()
'
'	'驗證憑證
'	atxCert.BinaryCert = cert
'	'msgbox("---有效日期自(double byte string)：")
'	'msgbox(atxCert.FromDate)
'	'msgbox("---有效日期自(long)：")
'	'msgbox(atxCert.FromDateBinary)
'	'msgbox("---有效日期至(double byte string)：")
'	'msgbox(atxCert.ToDate)
'	'msgbox("---有效日期至(long)：")
'	'msgbox(atxCert.ToDateBinary)
'	if now>atxCert.ToDate then
'		msgbox("此卡片已過期")
'		exit sub
'	end if
'
'	nRet = AresPKIClient.Init("aetpkss1.dll")
'	nRet = AresPKIClient.GetCertificate(0,cert)
'	AresPKIClient.Finalize()
'	nRet = AresPKIClient.DecodeCertificate(cert)
'	certSN = AresPKIClient.GetCertSubjectSN()
'	certSN = certSN + AresPKIClient.GetCertIssuerName()
'	certHex = AresPKIClient. B64ToHexString(AresPKIClient.B64Encode(certSN))
'	myForm.PKICarchk.value=certHex
'	chkPKI()
'End Sub

</script>
<%
conn.close
set conn=nothing
%>