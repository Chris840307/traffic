<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單單退</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->
<!-- #include file="../Common/Banner.asp"-->
<% Server.ScriptTimeout = 6800 %>
<%
'抓縣市
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
'權限
'AuthorityCheck(252)
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
		if trim(request("ReturnRecordDate_h"))<>"" or trim(request("ReturnRecordDate1_h"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and to_char(c.UserMarkDate,'hh') between "&trim(request("ReturnRecordDate_h"))&" and "&trim(request("ReturnRecordDate1_h"))
			else
				strwhere=" and to_char(c.UserMarkDate,'hh') between "&trim(request("ReturnRecordDate_h"))&" and "&trim(request("ReturnRecordDate1_h"))
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
				strwhere=strwhere&" and c.UserMarkMemberID="&request("Sys_RecordMemberID")
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
		if trim(request("DCIstatus"))="0" then
			if sys_City="澎湖縣" then
					'已結案不傳
					'抓註記人最後一次上傳
					UpdateTimeTmp=year(DateAdd("d",-1,now))&"/"&month(DateAdd("d",-1,now))&"/"&day(DateAdd("d",-1,now))&" "&hour(now)&":"&minute(now)&":"&second(now)
					if trim(request("Sys_RecordMemberID"))<>"" then
						strTime="select Max(ExchangeDate) as ExchangeDate from Dcilog where ExchangeTypeID='N' and ReturnMarkType='3' and RecordMemberID="&trim(request("Sys_RecordMemberID"))
					else
						strTime="select Max(ExchangeDate) as ExchangeDate from Dcilog where ExchangeTypeID='N' and ReturnMarkType='3' "
					end if
					set rsTime=conn.execute(strTime)
					if not rsTime.eof then
						if not isnull(rsTime("ExchangeDate")) or trim(rsTime("ExchangeDate"))<>"" then
							UpdateTimeTmp=year(rsTime("ExchangeDate"))&"/"&month(rsTime("ExchangeDate"))&"/"&day(rsTime("ExchangeDate"))&" "&hour(rsTime("ExchangeDate"))&":"&minute(rsTime("ExchangeDate"))&":"&second(rsTime("ExchangeDate"))
						end if
					end if
					rsTime.close
					set rsTime=nothing
					strwhere=strwhere&" and a.BillStatus='3' and not exists " &_
					"(select BillSN from DciLog" &_
					" where  a.SN =DciLog.billsn and ExchangeTypeID='N' and (DCIRETURNSTATUSID='n' or DCIRETURNSTATUSID is null)" &_
					" and ReturnMarkType='3') and c.UserMarkDate >  TO_DATE('"&UpdateTimeTmp&"','YYYY/MM/DD/HH24/MI/SS')" 
			elseif sys_City="基隆市" or sys_City="南投縣" or sys_City="苗栗縣" or (sys_City="高雄市" and Trim(Session("Unit_ID"))<>"0807") Or sys_City=ApconfigureCityName then
					'已結案還要再傳
					'抓註記人最後一次上傳
					UpdateTimeTmp=year(DateAdd("d",-1,now))&"/"&month(DateAdd("d",-1,now))&"/"&day(DateAdd("d",-1,now))&" "&hour(now)&":"&minute(now)&":"&second(now)
					if trim(request("Sys_RecordMemberID"))<>"" then
						strTime="select Max(ExchangeDate) as ExchangeDate from Dcilog where ExchangeTypeID='N' and ReturnMarkType='3' and RecordMemberID="&trim(request("Sys_RecordMemberID"))
					else
						strTime="select Max(ExchangeDate) as ExchangeDate from Dcilog where ExchangeTypeID='N' and ReturnMarkType='3' "
					end if
					set rsTime=conn.execute(strTime)
					if not rsTime.eof then
						if not isnull(rsTime("ExchangeDate")) or trim(rsTime("ExchangeDate"))<>"" then
							UpdateTimeTmp=year(DateAdd("n",1,rsTime("ExchangeDate")))&"/"&month(DateAdd("n",1,rsTime("ExchangeDate")))&"/"&day(DateAdd("n",1,rsTime("ExchangeDate")))&" "&hour(DateAdd("n",1,rsTime("ExchangeDate")))&":"&minute(DateAdd("n",1,rsTime("ExchangeDate")))&":"&second(rsTime("ExchangeDate"))
							'response.write UpdateTimeTmp
						end if
					end if
					rsTime.close
					set rsTime=nothing
					strwhere=strwhere&" and a.BillStatus='3' and not exists " &_
					"(select BillSN from DciLog" &_
					" where a.SN =DciLog.BillSN and ExchangeTypeID='N' and (DCIRETURNSTATUSID is null)" &_
					" and ReturnMarkType='3') and c.UserMarkDate >  TO_DATE('"&UpdateTimeTmp&"','YYYY/MM/DD/HH24/MI/SS')" 
			else
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillStatus='3' and not exists (select BillSN from DciLog where a.SN=DciLog.billsn and ExchangeTypeID='N' and (DCIRETURNSTATUSID='S' or DCIRETURNSTATUSID<>'S' or DCIRETURNSTATUSID is null) and ReturnMarkType='3')"
				else
					strwhere=" and a.BillStatus='3' and not exists (select BillSN from DciLog where a.SN=DciLog.billsn and ExchangeTypeID='N' and (DCIRETURNSTATUSID='S' or DCIRETURNSTATUSID<>'S' or DCIRETURNSTATUSID is null) and ReturnMarkType='3')"
				end if
			end if
		elseif trim(request("DCIstatus"))="1" then	'要抓最後依次
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='3' and a.Sn in (select x.BillSN from DCILog x,(select BillSN,Max(EXCHANGEDATE)as EXCHANGEDATE  from DCILog group BY BillSN) y where x.BIllSN=y.BillSN and x.EXCHANGEDATE=y.EXCHANGEDATE and x.DCIRETURNSTATUSID not in ('S','n') and x.EXCHANGETYPEID='N'and x.DCIRETURNSTATUSID is not null)"
			else
				strwhere=" and a.BillStatus='3' and a.Sn in (select x.BillSN from DCILog x,(select BillSN,Max(EXCHANGEDATE)as EXCHANGEDATE  from DCILog group BY BillSN) y where x.BIllSN=y.BillSN and x.EXCHANGEDATE=y.EXCHANGEDATE and x.DCIRETURNSTATUSID not in ('S','n') and x.EXCHANGETYPEID='N'and x.DCIRETURNSTATUSID is not null)"
			end if
		elseif trim(request("DCIstatus"))="2" then	
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and DciReturnStatusID='S' and ReturnMarkType='3') and a.SN not in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' and DciReturnStatusID is null) and c.StoreAndSendReturnResonID is not null"
			else
				strwhere=" and a.BillStatus='3' and a.SN in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and DciReturnStatusID='S' and ReturnMarkType='3') and a.SN not in (select distinct(BillSN) from DciLog where ExchangeTypeID='N' and ReturnMarkType='3' and DciReturnStatusID is null) and c.StoreAndSendReturnResonID is not null"
			end if
		else
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='3' and a.SN in (select x.BillSN from DCILog x,(select BillSN,Max(EXCHANGEDATE)as EXCHANGEDATE  from DCILog group BY BillSN) y where x.BIllSN=y.BillSN and x.EXCHANGEDATE=y.EXCHANGEDATE and x.DCIRETURNSTATUSID is not null and x.EXCHANGETYPEID='N' and ReturnMarkType='Y')"
			else
				strwhere=" and a.BillStatus='3' and a.SN in (select x.BillSN from DCILog x,(select BillSN,Max(EXCHANGEDATE)as EXCHANGEDATE  from DCILog group BY BillSN) y where x.BIllSN=y.BillSN and x.EXCHANGEDATE=y.EXCHANGEDATE and x.DCIRETURNSTATUSID is not null and x.EXCHANGETYPEID='N' and ReturnMarkType='Y')"
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

		'第一次或第二次做寄存送達
		if StoreAndSendMode=2 or sys_City="高雄市" Or sys_City=ApconfigureCityName then
			if strwhere<>"" then
				strwhere=strwhere&" and c.MailTypeID is null"
			else
				strwhere=" and c.MailTypeID is null"
			end if
		end if

		strSQL="select a.SN,a.IllegalDate,a.BillMem1,a.BillMem2,a.BillMem3,a.BillTypeID" &_
			",a.BillNo,a.CarNo,a.Driver,a.DriverID,a.Rule1,a.Rule2,a.Rule3" &_
			",a.Rule4,a.BillUnitID,a.BillStatus" &_
			",a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a" &_
			",BillMailHistory c where c.BillSN=a.SN "&strwhere &_
			" order by c.UserMarkDate"
end if

'response.write strSQL
'撤銷送達(遇到RecordStateID=-1不做)
if trim(request("kinds"))="BillCancel" Then
	strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=Nothing
	
	strReturn="select a.SN,a.IllegalDate,a.BillTypeID" &_
		",a.BillNo,a.CarNo" &_
		",a.BillUnitID,a.BillStatus,a.RecordDate" &_
		",a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID from BillBase a" &_
		",BillMailHistory c where a.RecordStateID<>-1" &_
		" c.BillSN=a.SN "&strwhere&" order by c.UserMarkDate"
	set rsReturn=conn.execute(strReturn)
	If Not rsReturn.Bof Then
		rsReturn.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("無可進行撤銷送達之舉發單！");
</script>
<%
	end if
	While Not rsReturn.Eof
		funcStoreAndSendToGov conn,trim(rsReturn("SN")),trim(rsReturn("BillNo")),trim(rsReturn("BillTypeID")),trim(rsReturn("CarNo")),trim(rsReturn("BillUnitID")),trim(rsReturn("RecordDate")),trim(Session("User_ID")),theBatchTime
	rsReturn.MoveNext
	Wend
	If Not rsReturn.Bof Then
%>
<script language="JavaScript">
	alert("撤銷送達處理完成，批號：<%=theBatchTime%>");
</script>
<%
	end if
	rsReturn.close
	set rsReturn=nothing

End If 

'退件(遇到RecordStateID=-1不做)
if trim(request("kinds"))="BillReturn" then
	if sys_City="台中市" then
	'台中市單退時要多做車籍查詢
		strQCSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsQCSN=conn.execute(strQCSN)
		if not rsQCSN.eof then
			theBatchTimeQryCar=(year(now)-1911)&"A"&trim(rsQCSN("SN"))
		end if
		rsQCSN.close
		set rsQCSN=nothing
	end if

	strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=nothing

	strReturn="select a.SN,a.IllegalDate,a.BillTypeID" &_
		",a.BillNo,a.CarNo" &_
		",a.BillUnitID,a.BillStatus,a.RecordDate" &_
		",a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID from BillBase a" &_
		",BillMailHistory c where " &_
		" c.BillSN=a.SN "&strwhere&" order by c.UserMarkDate"
	set rsReturn=conn.execute(strReturn)
	If Not rsReturn.Bof Then
		rsReturn.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("無可進行單退之舉發單！");
</script>
<%
	end if
	While Not rsReturn.Eof
		if (trim(rsReturn("UserMarkResonID"))="5" or trim(rsReturn("UserMarkResonID"))="6" or trim(rsReturn("UserMarkResonID"))="7" or trim(rsReturn("UserMarkResonID"))="T") and trim(rsReturn("StoreAndSendReturnResonID"))<>"" and not isnull(rsReturn("StoreAndSendReturnResonID")) then
			'寄存改公示
			'funcStoreAndSendToGov conn,trim(rsReturn("SN")),trim(rsReturn("BillNo")),trim(rsReturn("BillTypeID")),trim(rsReturn("CarNo")),trim(rsReturn("BillUnitID")),trim(rsReturn("RecordDate")),trim(rsReturn("RecordMemberID")),theBatchTime
		end if
		if sys_City="台中市" then
			'車籍查詢
			funcCarDataCheck conn,trim(rsReturn("SN")),trim(rsReturn("BillNo")),trim(rsReturn("BillTypeID")),trim(rsReturn("CarNo")),trim(rsReturn("BillUnitID")),trim(rsReturn("RecordDate")),trim(rsReturn("RecordMemberID")),theBatchTimeQryCar
		end if
		funcBillReturn conn,trim(rsReturn("SN")),trim(rsReturn("BillNo")),trim(rsReturn("BillTypeID")),trim(rsReturn("CarNo")),trim(rsReturn("BillUnitID")),trim(rsReturn("RecordDate")),trim(rsReturn("RecordMemberID")),theBatchTime
	rsReturn.MoveNext
	Wend
	If Not rsReturn.Bof Then
%>
<script language="JavaScript">
	alert("單退註記處理完成，批號：<%=theBatchTime%>");
</script>
<%
	end if
	rsReturn.close
	set rsReturn=nothing

end if

'做完車籍查詢及入案等動作後再查詢告發單，讓列表取得的資料為最新
if request("DB_Selt")="Selt" then
'response.write strSQL
'response.end
		set rsfound=conn.execute(strSQL)
		strCnt="select count(*) as cnt from BillBase a,BillMailHistory c" &_
			" where c.BillSN=a.SN "&strwhere
		set Dbrs=conn.execute(strCnt)
		if not Dbrs.eof then
			DBsum=Dbrs("cnt")
		end if
		Dbrs.close
		tmpSQL=strwhere
		'Session.Contents.Remove("BillSQL")
		'Session("BillSQL")=strSQL
		Session.Contents.Remove("PrintCarDataSQL")
		Session("PrintCarDataSQL")=strwhere
end if

%>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">舉發單單退</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<input type="hidden" name="ReturnRecordDateCheck" value="1" >
						退件註記日期
						<input name="ReturnRecordDate" type="text" value="<%
						if trim(request("DB_Selt"))="" then
							RecordDateTmp=ginitdt(DateAdd("d",-5,now))
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
						<!-- 時段
						<input name="ReturnRecordDate_h" type="text" value=" --><%'=request("ReturnRecordDate_h")%><!-- " size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時 ~ 
						<input name="ReturnRecordDate1_h" type="text" value=" --><%'=request("ReturnRecordDate1_h")%><!-- " size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時
						<img src="space.gif" width="8" height="10"> -->
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						退件註記人
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<br>
						DCI作業
						<select name="DCIstatus">
							<option value="0" <%
							if trim(request("DCIstatus"))="0" then response.write "selected"
							%>>進行單退</option>
							<option value="1" <%
							if trim(request("DCIstatus"))="1" then response.write "selected"
							%>>單退失敗</option>
							<option value="2" <%
							if trim(request("DCIstatus"))="2" then response.write "selected"
							%>>進行第二次單退</option>
							<option value="3" <%
							if trim(request("DCIstatus"))="3" then response.write "selected"
							%>>註銷送達後再次單退</option>
						</select>
						<img src="space.gif" width="8" height="10">
						舉發單號
						<input name="Sys_BillNo" type="text" value="<%=request("Sys_BillNo")%>" size="10" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="8" height="10">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" <%
						'if CheckPermission(252,1)=false then
						'	response.write "disabled"
						'end if
						%>>

						<input type="button" name="cancel" value="清除" onClick="location='BillBaseQryReturn.asp'"> 
						<br>
					<%if sys_City="基隆市" then%>
						<font color="red"><b>兩批案件處理間隔，請勿低於五分鐘，如查詢缺少註記案件，請於五分鐘後再重新註記一次即可</b></font>
						<br />
					<%end if %>
						<font color="red">如屬  撤銷送達後 再次單退  "DCI作業" 請選擇 "註銷送達後再次單退"</font>
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
					<th width="8%">DCI</th>
					<th width="19%"><%if sys_City="高雄市" Or sys_City=ApconfigureCityName then%>單退日期，<%end if%>單退原因，註記日期</th>
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
						response.write "<td>"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
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
						response.write "<td>"
						if trim(rsfound("BillStatus"))="0" then
							response.write "<font color='#999999'>未處理</font>"
						elseif trim(rsfound("BillStatus"))="1" then
							response.write "<font color='#FF66CC'>車籍查詢</font>"
						elseif trim(rsfound("BillStatus"))="2" then
							response.write "<font color='#009900'>入案</font>"
						elseif trim(rsfound("BillStatus"))="3" then
							response.write "<font color='#0000FF'>單退</font>"
						elseif trim(rsfound("BillStatus"))="4" then
							response.write "<font color='#0000FF'>寄存</font>"
						elseif trim(rsfound("BillStatus"))="5" then
							response.write "<font color='#0000FF'>公示</font>"
						elseif trim(rsfound("BillStatus"))="6" then
							response.write "<font color='#FF0000'>刪除</font>"
						end if
						response.write "</td>"
						response.write "<td>"
						strMail1="select MailReturnDate,UserMarkResonID,UserMarkDate,OpenGovMailReturnDate from BillMailHistory where BillSN="&trim(rsfound("SN"))
						set rsMail1=conn.execute(strMail1)
						if not rsMail1.eof then
							if trim(rsMail1("UserMarkResonID"))<>"" and not isnull(rsMail1("UserMarkResonID")) then
								strMDCode1="select Content from DCICode where TypeID=7 and ID='"&trim(rsMail1("UserMarkResonID"))&"'"
								set rsMDCode1=conn.execute(strMDCode1)
								if not rsMDCode1.eof then	
									if sys_City="高雄市" Or sys_City=ApconfigureCityName  then
										if (trim(rsMail1("UserMarkResonID"))="5" or trim(rsMail1("UserMarkResonID"))="6" or trim(rsMail1("UserMarkResonID"))="7" or trim(rsMail1("UserMarkResonID"))="T") then
										
											if trim(rsMail1("MailReturnDate"))<>"" and not isnull(rsMail1("MailReturnDate")) then
												response.write gInitDT(rsMail1("MailReturnDate"))
											end if
											response.write ","
										else
											if trim(rsMail1("OpenGovMailReturnDate"))<>"" and not isnull(rsMail1("OpenGovMailReturnDate")) then
												response.write gInitDT(rsMail1("OpenGovMailReturnDate"))
											end if
											response.write ","
										end if
									end if
									response.write trim(rsMDCode1("Content"))&","&gInitDT(rsMail1("UserMarkDate"))
									
								end if
								rsMDCode1.close
								set rsMDCode1=nothing
							end if
						end if
						rsMail1.close
						set rsMail1=nothing
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

			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
		<%If sys_City="苗栗縣" then%>
			<input type="button" name="Submit4244" value="撤銷送達" onclick="if(confirm('確定要做撤銷送達嗎？')){BillCancel()}">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="Submit4244" value="撤銷送達回傳後資料處理" onclick="BillCloseUpdate();">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
		<%End If %>
			<input type="button" name="Submit4244" value="單退註記" onclick="if(confirm('確定要做監理所單退嗎？')){BillReturn()}">
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
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate.value.substr(0,1)=="9" && myForm.ReturnRecordDate.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate.value.substr(0,1)=="1" && myForm.ReturnRecordDate.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}
		}
		if(myForm.ReturnRecordDate1.value!=""){
			if(!dateCheck(myForm.ReturnRecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate1.value.substr(0,1)=="9" && myForm.ReturnRecordDate1.value.length==7 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}else if( myForm.ReturnRecordDate1.value.substr(0,1)=="1" && myForm.ReturnRecordDate1.value.length==6 ){
				error=error+1;
				errorString=errorString+"\n"+error+"：單退註記日期輸入不正確!!";
			}
		}
		/*
		if(myForm.ReturnRecordDate_h.value!="" || myForm.ReturnRecordDate1_h.value!=""){
			if(myForm.ReturnRecordDate_h.value=="" || myForm.ReturnRecordDate1_h.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔時段輸入不完整!!";
			}
		}
		*/
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
	
	function BillCloseUpdate(){
		UrlStr="BillCloseUpdate.asp";
		newWin(UrlStr,"BillCloseUpdate",590,450,50,10,"yes","yes","yes","no");
	}

	function funchgExecel(){
		UrlStr="BillBaseQry_Execel.asp?WorkType=2";
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
	//退件
	function BillReturn(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲單退註記的舉發單！");
		}else{
			myForm.kinds.value="BillReturn";
			myForm.submit();
		}
	}
	//撤銷送達
	function BillCancel(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲單退註記的舉發單！");
		}else{
			myForm.kinds.value="BillCancel";
			myForm.submit();
		}

	}
</script>
<%
conn.close
set conn=nothing
%>