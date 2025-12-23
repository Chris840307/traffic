<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催繳單 / 各式清冊 列印</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
'檢查是否可進入本系統
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
RecordDate=split(gInitDT(date),"-")

if request("DB_Selt")="BatchSelt" then
	strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
	if UCase(request("Sys_BatchNumber"))<>"" then
		tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
		for i=0 to Ubound(tmp_BatchNumber)
			if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
			if i=0 then
				Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(tmp_BatchNumber(i))
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(tmp_BatchNumber(i))
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
		next
		strwhere=" and b.BatchNumber in ('"&Sys_BatchNumber&"')"
	end if

	if trim(request("Sys_ImageFileNameB1"))<>"" and trim(request("Sys_ImageFileNameB2"))<>"" then
		Sys_BillNo1=right("00000000000000000"&trim(request("Sys_ImageFileNameB1")),16)
		Sys_BillNo2=right("00000000000000000"&trim(request("Sys_ImageFileNameB2")),16)

		strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo1&"' and '"&Sys_BillNo2&"'"

	elseif trim(request("Sys_ImageFileNameB1"))<>"" then
		Sys_BillNo1=right("00000000000000000"&trim(request("Sys_ImageFileNameB1")),16)

		strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo1&"' and '"&Sys_BillNo1&"'"

	elseif trim(request("Sys_ImageFileNameB2"))<>"" then
		Sys_BillNo2=right("00000000000000000"&trim(request("Sys_ImageFileNameB2")),16)

		strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo2&"' and '"&Sys_BillNo2&"'"

	end if
	if request("Sys_IllegalDate1")<>"" and request("Sys_IllegalDate2")<>""then
		IllegalDate1=gOutDT(request("Sys_IllegalDate1"))&" 0:0:0"
		IllegalDate2=gOutDT(request("Sys_IllegalDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.IllegalDate between TO_DATE('"&IllegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&IllegalDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	DB_Display=request("DB_Display")
end if
if DB_Display="show" then
	If strwhere="" Then DB_Display=""

	if trim(strwhere)<>"" or (trim(request("Sys_UserMarkDate1"))<>"" and trim(request("Sys_UserMarkDate2"))<>"") or (trim(request("Sys_SendMarkDate1"))<>"" and trim(request("Sys_SendMarkDate2"))<>"") then
		if trim(strwhere)<>"" then
			strSQL="select distinct a.SN,a.CarNo,a.IllegalDate,a.ImageFileNameB from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere&" order by a.SN"

			set rssn=conn.execute(strSQL)
			BillSN="":tempBillSN="":strBillNo=""
			while Not rssn.eof
				If trim(tempBillSN)<>trim(rssn("SN")) Then
					tempBillSN=trim(rssn("SN"))
					if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
					BillSN=BillSN&trim(rssn("SN"))
				end if

				If instr(strBillNo,trim(rssn("ImageFileNameB")))=0 Then
					if trim(strBillNo)<>"" then strBillNo=trim(strBillNo)&","
					strBillNo=strBillNo&trim(rssn("ImageFileNameB"))
				end if

				rssn.movenext
			wend
			rssn.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A') b where a.SN=b.BillSN "&strwhere

			set Dbrs=conn.execute(strSQL)
			DBsum=Cint(Dbrs("cnt"))
			Dbrs.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&strwhere

			set chksuess=conn.execute(strSQL)
			filsuess=Cint(chksuess("cnt"))
			chksuess.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID in('N','E')) b where a.SN=b.BillSN "&strwhere
			set chksuess=conn.execute(strSQL)
			fildel=Cint(chksuess("cnt"))
			chksuess.close

			strSQL="select count(*) cnt from (select * from BillBase where ImagePathName is not null and RecordStateId = -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A') b where a.SN=b.BillSN "&strwhere
			set Dbrs=conn.execute(strSQL)
			deldata=Cint(Dbrs("cnt"))
			Dbrs.close

			strSQL2=strwhere
		end if
		twoSend=""
		if trim(request("Sys_SendMarkDate1"))<>"" and trim(request("Sys_SendMarkDate2"))<>""  then
			
			UserMarkDate1=gOutDT(request("Sys_SendMarkDate1"))&" 0:0:0"
			UserMarkDate2=gOutDT(request("Sys_SendMarkDate2"))&" 23:59:59"

			strSQL="select distinct BillNo from StopCarSendAddress where UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

			set rssn=conn.execute(strSQL)
			strBillNo=""
			while Not rssn.eof
				if trim(strBillNo)<>"" then strBillNo=trim(strBillNo)&","
				strBillNo=strBillNo&trim(rssn("BillNo"))

				rssn.movenext
			wend
			rssn.close

			strSQL="select count(1) cnt from StopCarSendAddress where UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

			set rsTwo=conn.execute(strSQL)
			if not rsTwo.eof then twoSend=cdbl(rsTwo("cnt"))
			rsTwo.close
		end if

		'單退要用註記日查詢
		if trim(request("Sys_UserMarkDate1"))<>"" and trim(request("Sys_UserMarkDate2"))<>""  then
			UserMarkDate1=gOutDT(request("Sys_UserMarkDate1"))&" 0:0:0"
			UserMarkDate2=gOutDT(request("Sys_UserMarkDate2"))&" 23:59:59"

			strwhere=strwhere&" and c.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

			strGet="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('A','B','C')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsGet=conn.execute(strGet)
			if not rsGet.eof then
				Getdata=Cint(rsGet("cnt"))
			end if
			rsGet.close
			set rsGet=nothing

			strROpen="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('1','2','3','4','8','K','L','M','O','P','Q')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsROpen=conn.execute(strROpen)
			if not rsROpen.eof then
				Opendata=Cint(rsROpen("cnt"))
			end if
			rsROpen.close
			set rsROpen=nothing

			strRStore="select count(*) as cnt from Billbase a,StopBillMailHistory b" &_
				" where a.Sn=b.BillSn and a.RecordStateID=0 and b.UserMarkResonID in ('5','6','7','T')" &_
				" and b.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsRStore=conn.execute(strRStore)
			if not rsRStore.eof then
				Storedata=Cint(rsRStore("cnt"))
			end if
			rsRStore.close
			set rsRStore=nothing
		end if
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
tmpSQL=strwhere
%>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr height="30">
		<td bgcolor="#FFCC33"><span class="style3">催繳單 / 各式清冊 列印</span><img src="space.gif" width="60" height="1"> <strong>請勿升級 Internet Explorer 7 . 避免套印舉發單出現異常</strong></img></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							strSQL="select distinct TO_char(ExchangeDate,'YYYY/MM/DD') ExchangeDate,BatchNumber from DCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",-5, date)&" 00:00"&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID='A' and DCIReturnStatusID='S' order by ExchangeDate DESC"
		
							set rs=conn.execute(strSQL)
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
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" size="29" onkeyup="funShowBillNo()">
						
						(<strong>多個批號同時處理</strong>，各批號請用,隔開。如：95A361,95A382,95A486）						
						<br>
						催繳單號
						<input name="Sys_ImageFileNameB1" type="text" class="btn1" value="<%=UCase(request("Sys_ImageFileNameB1"))%>" size="16" maxlength="16">
						~
						<input name="Sys_ImageFileNameB2" type="text" class="btn1" value="<%=UCase(request("Sys_ImageFileNameB2"))%>" size="16" maxlength="16"> ( 列印 <strong>單筆</strong> 或 特定範圍 催繳單才需填寫)
						<br>
						二次時間
						<input name="Sys_SendMarkDate1" type="text" class="btn1" value="<%=request("Sys_SendMarkDate1")%>" size="8" maxlength="7">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendMarkDate1');">
						~
						<input name="Sys_SendMarkDate2" type="text" class="btn1" value="<%=request("Sys_SendMarkDate2")%>" size="8" maxlength="7">
						<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_SendMarkDate2');">
						( 列印 <strong>二次催繳</strong>才需填寫)
						<br>
						註記時間
						<input name="Sys_UserMarkDate1" type="text" class="btn1" value="<%=request("Sys_UserMarkDate1")%>" size="8" maxlength="7">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_UserMarkDate1');">
						~
						<input name="Sys_UserMarkDate2" type="text" class="btn1" value="<%=request("Sys_UserMarkDate2")%>" size="8" maxlength="7">
						<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_UserMarkDate2');">
						( 列印 <strong>收受清冊</strong> 或 <strong>單退清冊</strong> 才需填寫)
						<br>
						停車時間
						<input name="Sys_IllegalDate1" type="text" class="btn1" value="<%=request("Sys_IllegalDate1")%>" size="8" maxlength="7">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_IllegalDate1');">
						~
						<input name="Sys_IllegalDate2" type="text" class="btn1" value="<%=request("Sys_IllegalDate2")%>" size="8" maxlength="7">
						<input type="button" name="datestr2" value="..." onclick="OpenWindow('Sys_IllegalDate2');">
							( 產生 <strong>公示檔</strong> 才需填寫)
						<br>
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt('BatchSelt');">
						<input type="button" name="cancel" value="清除" onClick="location='StopAllBillPrint.asp'"><br>
						<!--<input type="button" name="btnOK" value="匯入第二次催繳檔" onclick="funStopSelt();">-->
						<img src="space.gif" width="35" height="1"></img><strong ID="strCount">( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功 ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-filsuess-fildel-deldata%>筆未處理, <%=Getdata%> 筆收受, <%=Storedata%> 筆單退_寄存, <%=Opendata%> 筆單退_公示, <%=twoSend%> 筆二次郵寄. )</strong>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr>
		<td height="35" bgcolor="#FFDD77" align="left">
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;本批資料繳費期限
				<input name="Sys_DeallineDate" type="text" class="btn1" value="<%
					if ifnull(request("Sys_DeallineDate")) then
						If sys_City="台東縣" Then
							response.write gInitDT(DateAdd("d", 20,date()))
						else
							response.write gInitDT(DateAdd("d", 10,date()))
						end if
					else
						response.write request("Sys_DeallineDate")
					end if
				%>" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
				<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_DeallineDate');">
				&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="fun_DeallineDate();">
				&nbsp;&nbsp;<font color="red"><B><span id="showBillNoA""></span>&nbsp;&nbsp;<span id="showBillNoB"></span></B></font>
				<br><br>
				<!--<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 催繳單" onclick="funStopBillPrints_HuaLien()">
				<img src="space.gif" width="37" height="1"></img>-->
				<img src="space.gif" width="25" height="1"></img>
				
				
				<%If sys_City="台東縣" Then%>
					<img src="space.gif" width="8" height="1"></img>
					<!--
					<input type="button" name="btnprint" value="列印 第一次催繳單(新版)" onclick="funStopBillPrints_HuaLien3()">
					<img src="space.gif" width="8" height="1"></img>
					-->
					<input type="button" name="btnprint" value="列印 第一次催繳單(105年新版)" onclick="funStopBillPrints_HuaLien4()">
					<img src="space.gif" width="8" height="1"></img>
					<input type="button" name="btnprint" value="列印 第一次催繳單(107年新版)" onclick="funStopBillPrints_HuaLien5()">
					<br>
					<img src="space.gif" width="8" height="1"></img>
					<input type="button" name="btnprint" value="列印 第二次催繳單(105年新版)" onclick="funStopBillPrints_Two_TaiTung2()">
					<br>
					
					<img src="space.gif" width="37" height="1"></img>
					<input type="button" name="btnprint" value="匯出 停管寄存檔 " onclick="funExportStoreMarkTxt()">

				<%elseIf sys_City="花蓮縣" Then%>
					<input type="button" name="btnprint" value="列印 補印催繳單" onclick="funStopBillPrints_HuaLien_Mend()">
					<img src="space.gif" width="37" height="1"></img>

					<img src="space.gif" width="8" height="1"></img>
					<input type="button" name="btnprint" value="列印 第一次催繳單" onclick="funStopBillPrints_HuaLien2()">

				<%end if%>
				<img src="space.gif" width="23" height="1"></img>

				<!--<input type="button" name="btnprint" value="列印 補印催繳單(新版)" onclick="funStopBillPrints_HuaLien_Mend2()">
				<img src="space.gif" width="37" height="1"></img>
				-->


				<br>
				<img src="space.gif" width="37" height="1"></img>
				<input type="button" name="btnprint" value="匯出 第一次停管催繳檔 " onclick="funExportTxt_HL()">

				<%If sys_City="花蓮縣" Then%>
					<img src="space.gif" width="35" height="1"></img>
					<input type="button" name="btnprint" value="匯出 第一次停管催繳檔(新版)" onclick="funExportTxt_1020227_HL()">
				<%end if%>

				<br>
				<img src="space.gif" width="37" height="1"></img>

				<!--<input type="button" name="btnprint" value="匯出 停管第二次催繳檔 " onclick="funExportTxt_HLTwo()">
				<img src="space.gif" width="37" height="1"></img>-->

				<input type="button" name="btnprint" value="匯出 停管公示檔 " onclick="funExportOpenGovTxt()">

				<%If sys_City="花蓮縣" Then%>
					<img src="space.gif" width="35" height="1"></img>
					<input type="button" name="btnprint" value="匯出 停管公示檔(新版)" onclick="funExportOpenGovTxt_1020227()">
					
					<br>
					<img src="space.gif" width="37" height="1"></img>
					<input type="button" name="btnprint" value="列印 第二次催繳單(新版)" onclick="funStopBillPrints_TwoHuaLien2()">
				<%end if%>

				<!--<img src="space.gif" width="37" height="1"></img>
				<input type="button" name="btnprint" value="列印 第二次催繳單(新版)" onclick="funStopBillPrints_HuaLien2_Two()">-->

				<img src="space.gif" width="37" height="1"></img>
				<input type="button" name="btnprint" value="匯出 第二次停管催繳檔 " onclick="funExportTxt_Two_HL()">

				<%If sys_City="花蓮縣" Then%>
					<img src="space.gif" width="35" height="1"></img>
					<input type="button" name="btnprint" value="匯出 第二次停管催繳檔(新版)" onclick="funExportTxt_Two_1020227_HL()">
				<%end if%>
			<hr>
			<!--<span class="style3">
			DCI檔案名稱
			<input name="textfield42324" type="text" value="" size="14" maxlength="13">
			</span>-->
	
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList_HL()">

			<span class="style3"><img src="space.gif" width="22" height="8"></span>
			<input type="button" name="Submit4234" value="催繳清冊" onclick="funReportSendList_HL()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>

			<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit3f32" value="第二次催繳清冊" onclick="funReportSendList_HL_Second()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit3f32" value="第二次交寄大宗函件" onclick="funMailListSecond()">
		<br>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList_HL()">

			<span class="style3"><img src="space.gif" width="20" height="8"></span>
			<input type="button" name="Submit488423" value="退件清冊_寄存(全部)" onclick="funReturnSendList_Store_All()">
			<!--<span class="style3"><img src="space.gif" width="22" height="8"></span>
			<input type="button" name="Submit4234" value="第二次催繳清冊" onclick="funReportSendList_HL2()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit3f32" value="第二次交寄大宗函件" onclick="funMailList3()">-->

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="退件清冊_公示(全部)" onclick="funReturnSendList_Gov_All()">
			<%If sys_City="台東縣" Then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">

				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit423423" value="未退回清冊" onclick="funBillReturnList()">
				<br>
				公告日期：
				<input name="Sys_opengovDate" type="text" class="btn1" value="" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
				<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_opengovDate');">
				&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="fun_OpengovDate();">
				<br>
				本批資料一次郵寄日期
				<input name="Sys_BillBaseMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
				<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BillBaseMailDate');">
				&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSys_MailDate();">
			<%end if%>
			
				
	    <Br>
		<span class="style3"><img src="space.gif" width="10" height="8"></span>
		<input type="button" name="Submit3f32" value="郵寄未退回清冊" onclick="funMailNotReturnList()">
			
	
		<br>
		<br>
		<!--<HR>
		本批資料發文監理站日期
		<input name="Sys_SendOpenGovDocToStationDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendOpenGovDocToStationDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSendOpenGovDocToStationDate();">		
		<Br>
		本批資料二次郵寄日期
		&nbsp;&nbsp;&nbsp;&nbsp;<input name="Sys_StoreAndSendMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendMailDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendMailDate();">
		<br><br>-->
	</td>
  </tr>
  <tr>
    <td><p align="center">&nbsp;</p>    </td></tr>
<tr>
<td>
<b>催繳單列印設定</b>  <br>
印表機&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;OKI<br>
紙張格式 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;LEGAL 8.5 x 14 <br>
紙張來源 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;進紙夾1  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;(催繳單放最下方進紙夾,背面空白朝上.送達證書區域朝印表機內) <br>
上下左右邊界 : &nbsp;&nbsp; 0.166
</td>
</tr>
</table>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="BillSN" value="<%=BillSN%>">
<input type="Hidden" name="SQLstr" value="<%=strSQL2%>">
<input type="Hidden" name="BillPrintKind" value="<%=trim(request("BillPrintKind"))%>">
<input type="Hidden" name="PBillNo" value="<%=strBillNo%>">
<input type="Hidden" name="PCarNo" value="">
<input type="Hidden" name="PBillNo2" value="">
<input type="Hidden" name="PCarNo2" value="">
<input type="Hidden" name="newCode" value="0">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
funShowBillNo();
function fnBatchNumber(){
	myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
	funShowBillNo();
}

function funStopSelt(){
	newWin("","Address",700,200,50,10,"yes","yes","yes","no");
	UrlStr="StopSendStyle.asp";
	myForm.action=UrlStr;
	myForm.target="Address";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funShowBillNo(){
	if(myForm.Sys_BatchNumber.value.length>=5){
		runServerScript("StopchkShoBillNo.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value);
	}
}

function fun_DeallineDate(){
	sys_City='<%=sys_City%>';
	if(myForm.BillPrintKind.value=='2'){
		if (myForm.Sys_DeallineDate.value!=''){
			UrlStr="BillBaseDeadLineDateTwo.asp";
			myForm.action=UrlStr;
			myForm.target="BillBaseDeadLineDate";
			myForm.submit();
			myForm.action="";
			myForm.target="";
		}
	}else if (myForm.DB_Display.value!=""){
		if (myForm.Sys_DeallineDate.value!=''){
			UrlStr="BillBaseDeadLineDate.asp";
			if (sys_City=='花蓮縣'){
				UrlStr="BillBaseDeadLineDate_HuaLien_011225.asp";
			}

			myForm.action=UrlStr;
			myForm.target="BillBaseDeadLineDate";
			myForm.submit();
			myForm.action="";
			myForm.target="";
		}
	}
}

function fun_OpengovDate(){
	if (myForm.Sys_opengovDate.value!=''){
		UrlStr="StopOpengovDate.asp?SQLstr=<%=tmpSQL%>&Sys_opengovDate="+myForm.Sys_opengovDate.value;
		newWin(UrlStr,"StopOpengovDate",800,700,0,0,"yes","yes","yes","no");
	}
}

function funSys_MailDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_BillBaseMailDate.value!=''){
			
			var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			xmlhttp.Open("post","StopBillBaseMailDate.asp",false);	
			xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");
			xmlhttp.send("MailDate="+myForm.Sys_BillBaseMailDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_ImageFileNameB1.value+"&Sys_BillNo2="+myForm.Sys_ImageFileNameB2.value);
			alert("儲存完成!!");
		}
	}
}

function funStopBillPrints_HuaLien_Mend(){
	sys_City='<%=sys_City%>';
	<%if sys_City="花蓮縣" then
		Response.Write "myForm.newCode.value='0';"
		'Response.Write "if(confirm('要使用新廠商繳費代碼(245)嗎?')){"
		'Response.Write "myForm.newCode.value='1';}"
	End if
	%>
	UrlStr="StopBillPrints_HuaLien_Mend.asp";
	if (sys_City=='花蓮縣'){
		UrlStr="StopBillPrints_HuaLien_011225.asp";
	}
	myForm.action=UrlStr;
	myForm.target="StopBillPrints_Mend";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien_Mend2(){
	UrlStr="StopBillPrints_HuaLien_Mend2.asp";
	myForm.action=UrlStr;
	myForm.target="StopBillPrints_Mend2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien(){
	//UrlStr="StopBillPrints_HuaLien.asp";
	UrlStr="StopBillPrintsType_HuaLien.asp";
	myForm.action=UrlStr;
	myForm.target="StopBillPrints";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_TwoHuaLien2(){
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次日期查詢欲列印第二次催繳單！");
	}else{
		<%if sys_City="花蓮縣" then
			Response.Write "myForm.newCode.value='0';"
			'Response.Write "if(confirm('要使用新廠商繳費代碼(245)嗎?')){"
			'Response.Write "myForm.newCode.value='1';}"
		End if
		%>
		UrlStr="StopBillPrints_HuaLien2_Two.asp";
		myForm.BillPrintKind.value=2;
		myForm.action=UrlStr;
		myForm.target="StopBillPrints3";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funStopBillPrints_Two_TaiTung2(){
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次日期查詢欲列印第二次催繳單！");
	}else{
		UrlStr="StopBillPrints_HuaLien_TaiTung2_1050930.asp";
		myForm.BillPrintKind.value=2;
		myForm.action=UrlStr;
		myForm.target="StopBillPrints3";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funStopBillPrints_HuaLien2(){
	//UrlStr="StopBillPrints_HuaLien.asp";
	<%If sys_City="台東縣" Then
		Response.Write "UrlStr=""StopBillPrints_HuaLien2_TaiTung.asp"";"
	elseif sys_City="花蓮縣" then
		Response.Write "UrlStr=""StopBillPrints_HuaLien_011225.asp"";"
		Response.Write "myForm.newCode.value='0';"
		'Response.Write "if(confirm('要使用新廠商繳費代碼(245)嗎?')){"
		'Response.Write "myForm.newCode.value='1';}"
	End if
	%>
	myForm.BillPrintKind.value=1;
	myForm.action=UrlStr;
	myForm.target="StopBillPrints2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien3(){
	UrlStr="StopBillPrints_HuaLien3_TaiTung.asp";
	
	myForm.BillPrintKind.value=1;
	myForm.action=UrlStr;
	myForm.target="StopBillPrints2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien4(){
	UrlStr="StopBillPrints_HuaLien_TaiTung_1050603.asp";
	
	myForm.BillPrintKind.value=1;
	myForm.action=UrlStr;
	myForm.target="StopBillPrints2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien5(){
	UrlStr="StopBillPrints_HuaLien_TaiTung_1071116.asp";
	
	myForm.BillPrintKind.value=1;
	myForm.action=UrlStr;
	myForm.target="StopBillPrints2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStopBillPrints_HuaLien2_Two(){
	//UrlStr="StopBillPrints_HuaLien.asp";
	if(myForm.PBillNo.value!=''){
		UrlStr="StopBillPrints_HuaLien2.asp";
		myForm.BillPrintKind.value=2;
		myForm.action=UrlStr;
		myForm.target="StopBillPrints2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("目前無資料!!");
	}
}
function funSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==''&&myForm.Sys_ImageFileNameB1.value==''&&myForm.Sys_ImageFileNameB2.value==''&&myForm.Sys_UserMarkDate1.value==''&&myForm.Sys_UserMarkDate2.value==''&&myForm.Sys_IllegalDate1.value==''&&myForm.Sys_IllegalDate2.value==''&&myForm.Sys_SendMarkDate1.value==''&&myForm.Sys_SendMarkDate2.value==''){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(error==0){
			myForm.BillPrintKind.value=1;
			if(myForm.Sys_SendMarkDate1.value!=''&&myForm.Sys_SendMarkDate2.value!=''){
				myForm.BillPrintKind.value=2;
			}
			myForm.BillSN.value="";
			myForm.DB_Selt.value=DBKind;
			myForm.DB_Display.value='show';
			myForm.submit();
		}
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}

function funchgCarDataList_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
		alert("請先輸入作業批號或單號查詢欲列印車籍資料清冊的舉發單！");
	}else{
		UrlStr="StopDciPrintCarDataList.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","yes","yes","no");
	}
}

function funReportSendList_HL_Second(){
	var SqlTmp="<%=tmpSQL%>";
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次時間查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopReportSendList_New_Second_Excel.asp?Sys_SendMarkDate1="+myForm.Sys_SendMarkDate1.value+"&Sys_SendMarkDate2="+myForm.Sys_SendMarkDate2.value;
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}

function funReportSendList_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopReportSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}

function funReportSendList_HL2(){
	if(myForm.PBillNo.value!=''){
		//UrlStr="StopBillPrints_HuaLien.asp";
		UrlStr="StopReportSendList_Second_Excel.asp";
		myForm.action=UrlStr;
		myForm.target="StopReportSendList_Second_Excel";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("目前無資料!!");
	}
}

function funExportOpenGovTxt(){
	UrlStr="StopExportOpenGov_txt.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","yes","yes","no");
}

function funExportOpenGovTxt_1020227(){
	UrlStr="StopExportOpenGov_txt1020227.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","yes","yes","no");
}


function funExportStoreMarkTxt(){
	UrlStr="StopExportStoreMark_txt.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"DciStoreMark",790,575,50,10,"yes","yes","yes","no");
}

function funMailNotReturnList(){
	UrlStr="StopMailNotReturn.asp";
	newWin(UrlStr,"MailNotReturn",790,575,50,10,"yes","yes","yes","no");
}

function funExportTxt_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopExport_txt.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",10,10,0,0,"no","no","no","no");
	}
}

function funExportTxt_1020227_HL(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印催繳資料清冊的舉發單！");
	}else{
		UrlStr="StopExport_txt1020227.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",10,10,0,0,"no","no","no","no");
	}
}

function funExportTxt_Two_HL(){
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次時間查詢欲列印催繳資料清冊的舉發單！");
	}else{
		//winsub=newWin(UrlStr,"inputWin2",10,10,0,0,"no","no","no","no");
		var UrlStr="StopExport_two_txt.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}


function funExportTxt_Two_1020227_HL(){
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次時間查詢欲列印催繳資料清冊的舉發單！");
	}else{
		//winsub=newWin(UrlStr,"inputWin2",10,10,0,0,"no","no","no","no");
		var UrlStr="StopExport_two_txt1020227.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}


function funExportTxt_HLTwo(){
	if(myForm.PBillNo.value!=''){
		//newWin("","inputWin2",10,10,0,0,"no","no","no","no");
		UrlStr="StopExportTwo_txt.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("目前無資料!!");
	}
}

function funBillReturnList(){
	UrlStr="StopBillReturnSelect.asp";
	newWin(UrlStr,"BillReturnSelect",400,220,350,200,"no","no","no","no");
}

function funMailMoneyList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印郵費單的舉發單！");
	}else{
		UrlStr="StopMailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList",300,220,350,200,"no","no","no","no");
	}
}


function funMailList2(){
	var SqlTmp="<%=tmpSQL%>";
	if (SqlTmp==""){
			alert("請先輸入作業批號或單號查詢欲列印交寄大宗函件的舉發單！");
	}else{
		UrlStr="StopMailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>&MailSendType=S";
		newWin(UrlStr,"MailReportList",300,220,350,200,"no","no","no","no");
	}
}

function funMailListSecond(){
	if (myForm.Sys_SendMarkDate1.value=="" || myForm.Sys_SendMarkDate2.value==""){
			alert("請先輸入二次日期查詢欲列印第二次交寄大宗函件的舉發單！");
	}else{
		UrlStr="StopMailMoneyList_New_Second_Select.asp?Sys_SendMarkDate1="+myForm.Sys_SendMarkDate1.value+"&Sys_SendMarkDate2="+myForm.Sys_SendMarkDate2.value+"&MailSendType=S";
		newWin(UrlStr,"MailReportList",300,220,350,200,"no","no","no","no");
	}
}


function funMailList3(){
	if(myForm.PBillNo.value!=''){
		//UrlStr="StopBillPrints_HuaLien.asp";
		UrlStr="StopMailMoneyList_Second_Select.asp";
		myForm.action=UrlStr;
		myForm.target="StopMailMoneyList_Second";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("目前無資料!!");
	}
	
}

function funReturnSendList_Store_All(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印退件清冊的舉發單！");
	}else{

		UrlStr="StopReturnSendList_Excel_A3_Store_All.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin6a";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funGetSendList_HL(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印收受清冊的舉發單！");
	}else{

		UrlStr="StopGetSendList_Excel_A3.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin7a";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funReturnSendList_Gov_All(){
	if (myForm.Sys_UserMarkDate1.value=="" || myForm.Sys_UserMarkDate2.value==""){
			alert("請先輸入註記日期查詢欲列印退件清冊的舉發單！");
	}else{

		UrlStr="StopReturnSendList_Excel_A3_Gov_All.asp";
		myForm.action=UrlStr;
		myForm.target="inputWin8a";
		myForm.submit();
		myForm.action="";
		myForm.target="";

	}
}
</script>
<%conn.close%>