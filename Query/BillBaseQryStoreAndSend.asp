<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單寄存送達</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->
<!-- #include file="../Common/Banner.asp"-->
<% Server.ScriptTimeout = 8800 %>
<%
'抓縣市
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
'權限
'AuthorityCheck(253)
RecordDate=split(gInitDT(date),"-")
'組成查詢SQL字串
if request("DB_Selt")="Selt" then
		strwhere=""
		if trim(request("ReturnRecordDateCheck"))="1" then
			if request("ReturnRecordDate")<>"" and request("ReturnRecordDate1")<>""then
				ReturnRecordDate1=gOutDT(request("ReturnRecordDate"))&" 0:0:0"
				ReturnRecordDate2=gOutDT(request("ReturnRecordDate1"))&" 23:59:59"
				if strwhere<>"" then
					strwhere=strwhere&" and c.UserMarkDate between TO_DATE('"&ReturnRecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ReturnRecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				else
					strwhere=" and c.UserMarkDate between TO_DATE('"&ReturnRecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ReturnRecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				end if
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
			strwhere=strwhere&" and c.UserMarkMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
		end if
		if request("Sys_RecordMemberID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and c.UserMarkMemberID="&request("Sys_RecordMemberID")
			else
				strwhere=" and c.UserMarkMemberID="&request("Sys_RecordMemberID")
			end if
		end if
		if request("Sys_BillTypeID")<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
			else
				strwhere=" and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
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
		if request("Sys_BatchNo")<>"" then
			strBatch=" and BatchNumber in ('"&replace(request("Sys_BatchNo"),",","','")&"')"
		else
			strBatch=""
		end if
		
		if request("DCIstatus")="0" then
			if sys_City="基隆市" or sys_City="台中市" or sys_City="雲林縣" or sys_City="高雄縣" or sys_City="台東縣" or sys_City="澎湖縣" or sys_City="彰化縣" or sys_City="高雄市" or sys_City="南投縣" Or sys_City=ApconfigureCityName then
				if request("Sys_BatchNo")<>"" then
					strBatch2=" and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N'" &_
						" and (DCIRETURNSTATUSID='S' or DCIRETURNSTATUSID<>'S' or DCIRETURNSTATUSID is null) and ReturnMarkType='3'" &_
						" and BatchNumber in ('"&replace(request("Sys_BatchNo"),",","','")&"'))"
				else
					strBatch2=""
				end if
				strwhere=strwhere&" and a.BillStatus='3' "&strBatch2&""
			'elseif sys_City="南投縣" then
			'	strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' and DCIRETURNSTATUSID<>'n' "&strBatch&") and c.StoreAndSendSendDate is null"

			elseif sys_City="台南市" or (sys_City="金門縣" and trim(Session("Unit_ID"))="0400") then
				strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' and DCIRETURNSTATUSID<>'n' "&strBatch&")"

			elseif sys_City="宜蘭縣" or sys_City="嘉義市" or sys_City="嘉義縣" then
				if request("Sys_BatchNo")<>"" Then
					If sys_City="宜蘭縣" Then
						strwhere=strwhere&" and a.BillStatus in ('3','9') and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' "&strBatch&")"
					else
						strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' and DCIRETURNSTATUSID<>'n' "&strBatch&")"
					End if
				else
					strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select BillSN from BillMailHistory where StoreAndSendReCordDate between to_date('"& Date() &" 00:00:00','YYYY/MM/DD/HH24/MI/SS') and to_date('"& Date() &" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and MailTypeID=2)"

				end if
				

			else
				strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' "&strBatch&")"
			end if
		elseif request("DCIstatus")="1" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='4' and a.Sn in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='N' and (DciReturnStatusID='S' or DciReturnStatusID='n' or DciReturnStatusID='h' or DciReturnStatusID is null) and ReturnMarkType='4' "&strBatch&"))"
			else
				strwhere=" and a.BillStatus='4' and a.Sn in (select distinct(BillSN) from DciLog where BillSN not in (select Billsn from DciLog where exchangeTypeID='N' and (DciReturnStatusID='S' or DciReturnStatusID='n' or DciReturnStatusID='h' or DciReturnStatusID is null) and ReturnMarkType='4' "&strBatch&"))"
			end if
		end if
		
		if strwhere<>"" then
			strwhere=strwhere&" and a.RecordStateID=0"
		else
			strwhere=" and a.RecordStateID=0"
		end if
		'第一次或第二次做寄存送達
		if StoreAndSendMode=2 or sys_City="高雄市" Or sys_City=ApconfigureCityName or sys_City="澎湖縣" then
			if strwhere<>"" then
				strwhere=strwhere&" and c.MailTypeID=2"
			else
				strwhere=" and c.MailTypeID=2"
			end if
		end if

		if OpenAndGovMode=2 then
			if strwhere<>"" then
				strwhere=strwhere&" and (c.MailTypeID<>6 or c.MailTypeID is null)"
			else
				strwhere=" and (c.MailTypeID<>6 or c.MailTypeID is null)"
			end if
		end if

		'是否要判斷一打一驗 1:是 0:否
		if Session("DoubleCheck")="1" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.DoubleCheckStatus=1"
			else
				strwhere=" and a.DoubleCheckStatus=1"
			end if
		end if

		CancelBillNo=""
		strchk="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID,c.UserMarkDate,c.StoreAndSendMailReturnDate,c.StoreAndSendEffectDate from BillBase a,MemberData b,BillMailHistory c where a.RecordMemberID=b.MemberID and c.BillSN=a.SN and c.UserMarkResonID in ('5','6','7','T')"&strwhere&" and exists (select Billsn from dcilog where ExchangeTypeID='N' and ReturnMarkType='Y' and DciReturnStatusID is null and billsn=a.sn) order by c.UserMarkDate"
		set rschk=conn.execute(strchk)
		if not rschk.eof then
			rschk.MoveFirst 
			While Not rschk.Eof
				if CancelBillNo="" then
					CancelBillNo=trim(rschk("BillNo"))
				else
					CancelBillNo=CancelBillNo&"、"&trim(rschk("BillNo"))
				end if
				rschk.MoveNext
			Wend
		end if
		rschk.close
		set rschk=nothing
		if CancelBillNo<>"" then
%>
		<script language="JavaScript">
			alert("<%=CancelBillNo%> 撤銷送達監理站尚未處理，請等撤銷送達回傳後，再進行寄存送達，避免資料發生錯誤!!");
		</script>
<%
		end if

		strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID,c.UserMarkDate,c.StoreAndSendMailReturnDate,c.StoreAndSendEffectDate from BillBase a,MemberData b,BillMailHistory c where c.BillSN=a.SN and c.UserMarkResonID in ('5','6','7','T') and a.RecordMemberID=b.MemberID "&strwhere&" order by c.UserMarkDate"
end if


'寄存(遇到RecordStateID=-1不做)
if trim(request("kinds"))="SafeKeeping" then

	strSafe="select a.SN,a.IllegalDate,a.BillTypeID,a.BillNo,a.CarNo,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID,c.UserMarkDate,c.StoreAndSendMailReturnDate from BillBase a,MemberData b,BillMailHistory c where c.BillSN=a.SN and c.UserMarkResonID in ('5','6','7','T') and a.RecordMemberID=b.MemberID "&strwhere&"  order by c.UserMarkDate"
	set rsSafe=conn.execute(strSafe)
	If Not rsSafe.Bof Then
		rsSafe.MoveFirst 
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing
	else
%>
<script language="JavaScript">
	alert("無可進行寄存送達之舉發單！");
</script>
<%
	end if
	While Not rsSafe.Eof
		funcSafeKeep conn,trim(rsSafe("SN")),trim(rsSafe("BillNo")),trim(rsSafe("BillTypeID")),trim(rsSafe("CarNo")),trim(rsSafe("BillUnitID")),trim(rsSafe("RecordDate")),trim(rsSafe("RecordMemberID")),theBatchTime
	rsSafe.MoveNext
	Wend
	If Not rsSafe.Bof then
%>
<script language="JavaScript">
	alert("監理站寄存送達註記完成，批號：<%=theBatchTime%>");
</script>
<%
	end if
	rsSafe.close
	set rsSafe=nothing

end if


'做完車籍查詢及入案等動作後再查詢告發單，讓列表取得的資料為最新
if request("DB_Selt")="Selt" then
'response.write strSQL
'response.end
		set rsfound=conn.execute(strSQL)
		strCnt="select count(*) as cnt from BillBase a,MemberData b,BillMailHistory c where a.RecordMemberID=b.MemberID and c.BillSN=a.SN and c.UserMarkResonID in ('5','6','7','T')"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close
		tmpSQL=strwhere
		Session.Contents.Remove("BillSQLforStoreAndSendUpload")
		Session("BillSQLforStoreAndSendUpload")=strSQL
		Session.Contents.Remove("PrintCarDataSQL")
		Session("PrintCarDataSQL")=strwhere
end if

%>
<html>

</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">舉發單寄存送達</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<input type="checkbox" name="ReturnRecordDateCheck" value="1" <%
						DateChk=trim(request("ReturnRecordDateCheck"))
						if DateChk="1" then
							response.write "checked"
						end if
						%>>
						送達註記日期
						<input name="ReturnRecordDate" type="text" value="<%
						if trim(request("DB_Selt"))="" then
							RecordDateTmp=ginitdt(now)
						else
							RecordDateTmp=trim(request("ReturnRecordDate"))
						end if
						response.write RecordDateTmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ReturnRecordDate');">
						~
						<input name="ReturnRecordDate1" type="text" value="<%
						if trim(request("DB_Selt"))="" then
							RecordDate1Tmp=ginitdt(now)
						else
							RecordDate1Tmp=trim(request("ReturnRecordDate1"))
						end if
						response.write RecordDate1Tmp
						%>" size="8" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('ReturnRecordDate1');">
						<img src="space.gif" width="8" height="10">
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						送達註記人員
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<br>
						<img src="space.gif" width="8" height="10">
						DCI作業
						<select name="DCIstatus">
							<option value="0" <%
							if trim(request("DCIstatus"))="0" then response.write "selected"
							%>>寄存送達</option>
							<option value="1" <%
							if trim(request("DCIstatus"))="1" then response.write "selected"
							%>>寄存送達失敗</option>
						</select>
						<img src="space.gif" width="8" height="10">
						舉發單號
						<input name="Sys_BillNo" type="text" value="<%=request("Sys_BillNo")%>" size="10" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="8" height="10">
						<strong>單退批號</strong>
						<input name="Sys_BatchNo" type="text" value="<%=request("Sys_BatchNo")%>" size="20" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="8" height="10">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" <%
						if CheckPermission(253,1)=false then
							response.write "disabled"
						end if
						%>>
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseQryStoreAndSend.asp'"> 
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
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th width="5%">違規日期</th>
					<th width="8%">舉發員警</th>
					<th width="6%">舉發單號</th>
					<th width="5%">車號</th>
					<th width="4%">類別</th>
					<th width="10%">法條</th>
					<!-- <th width="8%">DCI</th> -->
					<th width="20%"><%if sys_City="高雄縣" or sys_City="高雄市" or sys_City="台中市" Or sys_City=ApconfigureCityName then%>送達日期，<%end if%>送達原因，註記日期</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
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

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='35'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
						response.write "<td>"&chname&"</td>"
'					if trim(rsfound("BillTypeID"))="2" then
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("BillNo")&"</a></td>"
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					else
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("BillNo")&"</a></td>"
'						response.write "<td width='6%'><a href='../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&rsfound("CarNo")&"</a></td>"
'					end if
						response.write "<td width='6%'>"&rsfound("BillNo")&"</td>"
						response.write "<td width='6%'>"&rsfound("CarNo")&"</td>"

						response.write "<td>"
					strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
					set rsBTypeVal=conn.execute(strBTypeVal)
					if not rsBTypeVal.eof then
						response.write rsBTypeVal("Content")
					end if
					rsBTypeVal.close
					set rsBTypeVal=nothing
						response.write "</td>"
						response.write "<td>"&chRule&"</td>"

'						response.write "<td>"
'						if trim(rsfound("BillStatus"))="0" then
'							response.write "<font color='#999999'>未處理</font>"
'						elseif trim(rsfound("BillStatus"))="1" then
'							response.write "<font color='#FF66CC'>車籍查詢</font>"
'						elseif trim(rsfound("BillStatus"))="2" then
'							response.write "<font color='#009900'>入案</font>"
'						elseif trim(rsfound("BillStatus"))="3" then
'							response.write "<font color='#0000FF'>單退</font>"
'						elseif trim(rsfound("BillStatus"))="4" then
'							response.write "<font color='#0000FF'>寄存</font>"
'						elseif trim(rsfound("BillStatus"))="5" then
'							response.write "<font color='#0000FF'>公示</font>"
'						elseif trim(rsfound("BillStatus"))="6" then
'							response.write "<font color='#FF0000'>刪除</font>"
'						end if
'						response.write "</td>"
						response.write "<td>"
						if sys_City="高雄縣" or sys_City="高雄市" or sys_City="台中市" Or sys_City=ApconfigureCityName then
							if not isnull(rsfound("StoreAndSendEffectDate")) and rsfound("StoreAndSendEffectDate")<>"" then
								response.write gInitDT(rsfound("StoreAndSendEffectDate"))&","
							end if
						end if
						if trim(rsfound("UserMarkResonID"))<>"" and not isnull(rsfound("UserMarkResonID")) then
							strMDCode1="select Content from DCICode where TypeID=7 and ID='"&trim(rsfound("UserMarkResonID"))&"'"
							set rsMDCode1=conn.execute(strMDCode1)
							if not rsMDCode1.eof then	
								response.write trim(rsMDCode1("Content"))&","&gInitDT(rsfound("UserMarkDate"))
							end if
							rsMDCode1.close
							set rsMDCode1=nothing
						end if
						response.write "</td>"

						response.write "</tr>"
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
		<%
		if sys_City="高雄縣X" then
		%>
			<input type="button" name="b1" value="進行寄存送達" onclick="StoreAndSend2()">
		<%
		'送達證書回來才做寄存送達的話，那就直接作上傳監理站
		elseif StoreAndSendMode=2 or sys_City="彰化縣" or sys_City="南投縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName or (sys_City="金門縣" and trim(Session("Unit_ID"))="0400") then
		%>
			<input type="button" name="b1" value="進行寄存送達" onclick="SafeKeeping()">
		<%else%>
			<input type="button" name="b1" value="進行寄存送達" onclick="StoreAndSend()">
		<%end if%>
			<span class="style3"><img src="space.gif" width="5" height="8"></span>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="hidden" name="DelReason" value="">
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
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&trim(request("Sys_RecordMemberID"))&"');"%>
	function funSelt(){
		var error=0;
		var errorString="";
		if(myForm.ReturnRecordDate.value!=""){
			if(!dateCheck(myForm.ReturnRecordDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate.value.substr(0,1)=="9" && myForm.ReturnRecordDate.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate.value.substr(0,1)=="1" && myForm.ReturnRecordDate.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
			}
		}
		if(myForm.ReturnRecordDate1.value!=""){
			if(!dateCheck(myForm.ReturnRecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate1.value.substr(0,1)=="9" && myForm.ReturnRecordDate1.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate1.value.substr(0,1)=="1" && myForm.ReturnRecordDate1.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：送達註記日期輸入不正確!!";
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
		UrlStr="BillBaseQry_Execel.asp?WorkType=3";
		newWin(UrlStr,"inputWin",790,550,50,10,"yes","yes","yes","no");
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
	//寄存(StoreAndSendMode=2)
	function SafeKeeping(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲寄存的舉發單！");
		}else{
			myForm.kinds.value="SafeKeeping";
			myForm.submit();
		}
	}
	//寄存(StoreAndSendMode=1)
	function StoreAndSend(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲寄存的舉發單！");
		}else{
			window.open("../BillReturn/BillBaseStoreAndSendDCI.asp","WebPage_Del_Bill","left=0,top=0,location=0,width=600,height=400,resizable=yes,scrollbars=yes")
		}
	}
	function StoreAndSend2(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲寄存的舉發單！");
		}else{
			window.open("../BillReturn/BillBaseStoreAndSendDCI_KA.asp","WebPage_Del_Bill","left=0,top=0,location=0,width=600,height=400,resizable=yes,scrollbars=yes")
		}
	}
</script>
<%
conn.close
set conn=nothing
%>