<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-size: 13px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>

<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<%	
	CaseInDate=""
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQLTemp=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp=" where BillNO='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("CarNo"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and CarNo like '%"&trim(request("CarNo"))&"%'"
		else
			strSQLTemp=" where CarNo Like '%"&trim(request("CarNo"))&"%'"
		end if
	end if
	if trim(request("illFID"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and DriverID='"&trim(request("illFID"))&"'"
		else
			strSQLTemp=" where DriverID='"&trim(request("illFID"))&"'"
		end if
	end if
	if trim(request("illName"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and Driver='"&trim(request("illName"))&"'"
		else
			strSQLTemp=" where Driver='"&trim(request("illName"))&"'"
		end if
	end if
	if trim(request("IllegalDate"))<>"" and trim(request("IllegalDate1"))<>"" then
		RecordDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
		RecordDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strSQLTemp=" where IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
	end if
	if trim(request("BillSn"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and SN='"&trim(request("BillSn"))&"'"
		else
			strSQLTemp=" where SN='"&trim(request("BillSn"))&"'"
		end if
	end if
	if trim(request("MailNo"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and MailNumber='"&trim(request("MailNo"))&"'"
		else
			strSQLTemp=" where MailNumber='"&trim(request("MailNo"))&"'"
		end if
		strSQL="select a.* from BillBase a,BillMailHistory b"&strSQLTemp&" and a.SN=b.BillSN"
	else
		strSQL="select * from BillBase"&strSQLTemp
	end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("查無資料！");
	window.close();
</script>	
<%
	end if
	While Not rs1.Eof
%>
	<table width='100%' border='1' cellpadding="2">
		<tr bgcolor="#1BF5FF">
			<td colspan="6"><strong>舉發單詳細資料</strong></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" width="20%"><strong>單號</strong></td>
			<td align="center" width="20%"><strong>車號</strong></td>
			<td align="center" width="20%"><strong>違規法條</strong></td>
			<td align="center" width="20%"><strong>舉發單狀態</strong></td>
			<td align="center" width="20%"><strong>舉發員警</strong></td>
		</tr>
		<tr>
			<td align="center"><%
			'單號
			if trim(rs1("BillNo"))="" or isnull(rs1("BillNo")) then
				response.write "&nbsp;"
			else
				response.write trim(rs1("BillNo"))
			end if
			%></td>
			<td align="center"><%
			'車號
			if trim(rs1("CarNo"))="" or isnull(rs1("CarNo")) then
				response.write "&nbsp;"
			else
				response.write trim(rs1("CarNo"))
			end if
			%></td>
			<td align="center"><%
			'違規法條
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				chRule=rs1("Rule1")
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				chRule=chRule&"/"&rs1("Rule2")
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				chRule=chRule&"/"&rs1("Rule3")
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then 
				chRule=chRule&"/"&rs1("Rule4")
			end if
			response.write chRule
			%></td>
			<td align="center"><%
			'舉發單狀態
			if trim(rs1("RecordStateID"))<>"" and not isnull(rs1("RecordStateID")) then
				if trim(rs1("RecordStateID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'舉發人
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				If trim(rs1("BillMemID1"))<>"" then
					strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write "("&trim(rsMem1("LoginID"))&")"
					end if
					rsMem1.close
					set rsMem1=Nothing
				End If 
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
				If trim(rs1("BillMemID2"))<>"" then
					strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
					set rsMem2=conn.execute(strMem2)
					if not rsMem2.eof then
						response.write "("&trim(rsMem2("LoginID"))&")"
					end if
					rsMem2.close
					set rsMem2=Nothing
				End If 
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
				If trim(rs1("BillMemID3"))<>"" then
					strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
					set rsMem3=conn.execute(strMem3)
					if not rsMem3.eof then
						response.write "("&trim(rsMem3("LoginID"))&")"
					end if
					rsMem3.close
					set rsMem3=Nothing
				End If 
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
				If trim(rs1("BillMemID4"))<>"" then
					strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
					set rsMem4=conn.execute(strMem4)
					if not rsMem4.eof then
						response.write "("&trim(rsMem4("LoginID"))&")"
					end if
					rsMem4.close
					set rsMem4=Nothing
				End If 
			end if
			%></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" width="20%"><strong>填單日</strong></td>
			<td align="center" width="20%"><strong>建檔日,建檔人</strong></td>
			<td align="center" width="20%"><strong>違規日</strong></td>
			<td align="center" width="20%"><strong>入案日,入案人</strong></td>
			<td align="center" width="20%"><strong>刪除日,刪除人</strong></td>
		</tr>
		<tr>
			<td align="center"><%
			'填單日期
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'建檔日期
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))&" "
				response.write Right("00"&hour(rs1("RecordDate")),2)&":"
				response.write Right("00"&minute(rs1("RecordDate")),2)
			else
				response.write "&nbsp;"
			end if
			%>&nbsp;,&nbsp;<%
			'建檔人
			if trim(rs1("RecordMemberID"))<>"" and not isnull(rs1("RecordMemberID")) then
				strRMem="select ChName from MemberData where MemberID="&trim(rs1("RecordMemberID"))
				set rsRMem=conn.execute(strRMem)
				if not rsRMem.eof then
					response.write trim(rsRMem("ChName"))
				end if
				rsRMem.close
				set rsRMem=nothing
			else
				response.write "&nbsp;"
			end if
			%>
			
			</td>
			<td align="center"><%
			'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'入案日,入案人
			strCaseIn="select d.ChName,b.DciCaseInDate from DciLog a," &_
			"BillbaseDciReturn b,DCIReturnStatus c,MemberData d" &_
			" where a.BillSn="&trim(rs1("Sn"))&" and a.BillNo=b.BillNo and a.CarNo=b.CarNo" &_
			" and a.ExchangeTypeID='W' and a.ExchangeTypeID=c.DCIActionID " &_
			" and a.DCIReturnStatusID=c.DCIReturn and a.ExchangeTypeID=b.ExchangeTypeID" &_
			" and a.DciReturnStatusID=b.Status and c.DCIreturnStatus=1 and d.MemberID=a.RecordMemberID"
			set rsCaseIn=conn.execute(strCaseIn)
			if not rsCaseIn.eof Then
				If Len(trim(rsCaseIn("DCICASEINDATE")))=6 then
					response.write mid(trim(rsCaseIn("DCICASEINDATE")),1,2)
					response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),3,2)
					response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),5,2)
				Else
					response.write mid(trim(rsCaseIn("DCICASEINDATE")),1,3)
					response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),4,2)
					response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),6,2)
				End If 
				response.write "&nbsp;,&nbsp;"&trim(rsCaseIn("ChName"))
			else
				response.write "&nbsp;"
			end if
			rsCaseIn.close
			set rsCaseIn=nothing
			%></td>
			<td align="center"><%
			'刪除日期
			strDelLog="select ActionDate from Log where TypeID=352 and ActionContent Like '%"&trim(rs1("BillNo"))&"%' and ActionContent Like '%"&trim(rs1("CarNo"))&"%'"
			set rsDelLog=conn.execute(strDelLog)
			if not rsDelLog.eof Then  '20110222 by jafe 原本為西元年修改為民國年
				response.write gArrDT(trim(rsDelLog("ActionDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rsDelLog("ActionDate")),2)&":"
				response.write Right("00"&minute(rsDelLog("ActionDate")),2)
			Else
    			strDelLog="select deldate from BILLDELETEREASON where Billsn="&rs1("sn")
				set rsDelLog2=conn.execute(strDelLog)
				if not rsDelLog2.eof Then  '20110222 by jafe 入案後七天之後系統會把刪除資料清掉以利使用者作業，故抓這個資料表的刪除時間
					response.write gArrDT(trim(rsDelLog2("deldate")))&"&nbsp;&nbsp;"
					response.write Right("00"&hour(rsDelLog2("deldate")),2)&":"
					response.write Right("00"&minute(rsDelLog2("deldate")),2)
				End if
			end if
			rsDelLog.close
			set rsDelLog=nothing
			'刪除人
			if trim(rs1("DelMemberID"))<>"" and not isnull(rs1("DelMemberID")) then
				strDMem="select ChName from MemberData where MemberID="&trim(rs1("DelMemberID"))
				set rsDMem=conn.execute(strDMem)
				if not rsDMem.eof then
					response.write "&nbsp;,&nbsp;"&trim(rsDMem("ChName"))
				end if
				rsDMem.close
				set rsDMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" width="20%"><strong>單退日,單退人</strong></td>
			<td align="center" width="20%"><strong>寄存送達日,寄存人</strong></td>
			<td align="center" width="20%"><strong>公示送達日</strong></td>
			<td align="center" width="20%"><strong>違規影像資料</strong></td>
			<td align="center" width="20%"><strong>寄存送達證書掃描檔</strong></td>
		</tr>
		<tr>
<%
	strMailHistory2="select * from BillMailHistory where BillSN="&trim(rs1("SN"))
	set rsMH2=conn.execute(strMailHistory2)
	if not rsMH2.eof then
			OPENGOVGOVNUMBER=""
			'公示送達文號
			if trim(rsMH2("OPENGOVNUMBER"))<>"" and not isnull(rsMH2("OPENGOVNUMBER")) then

				OPENGOVGOVNUMBER=trim(rsMH2("OPENGOVNUMBER"))
			end if
			'第一次雙掛號寄存郵局
			if trim(rsMH2("MailStation"))<>"" and not isnull(rsMH2("MailStation")) then
				FirstMailStation=trim(rsMH2("MailStation"))
			else
				FirstMailStation="&nbsp;"
			end if
			'代收人
			if trim(rsMH2("SignMan"))<>"" and not isnull(rsMH2("SignMan")) then
				SignMan=trim(rsMH2("SignMan"))
			else
				SignMan="&nbsp;"
			end if
			'移送監理站時間
			if trim(rsMH2("SendOpenGovDocToStationDate"))<>"" and not isnull(rsMH2("SendOpenGovDocToStationDate")) then
				SendStationDate=trim(rsMH2("SendOpenGovDocToStationDate"))
			else
				SendStationDate="&nbsp;"
			end if

%>
			<td align="center"><%
			'檢查是單退還是收受
			strCheck="select count(*) as cnt from Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7'"
			set rsCheck=conn.execute(strCheck)
			'if not rsCheck.eof then
				if rsCheck("cnt")="0" then
					CheckFlag2=0
				else
					CheckFlag2=1
				end if
			'end if
			rsCheck.close
			set rsCheck=nothing
			'退件日期
			if CheckFlag2=0 then
				if trim(rsMH2("MAILRETURNDATE"))<>"" and not isnull(rsMH2("MAILRETURNDATE")) then
					response.write gArrDT(trim(rsMH2("MAILRETURNDATE")))&"&nbsp;,&nbsp;"
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			'退件註記人員
			if CheckFlag2=0 then
				if trim(rsMH2("RETURNRECORDMEMBERID"))<>"" and not isnull(rsMH2("RETURNRECORDMEMBERID")) then
					strReturnRecMem="select chName from MemberData where MemberID="&trim(trim(rsMH2("RETURNRECORDMEMBERID")))
					set rsRRMem=conn.execute(strReturnRecMem)
					if not rsRRMem.eof then
						response.write trim(rsRRMem("chName"))
					end if
					rsRRMem.close
					set rsRRMem=nothing
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'寄存送達單退日
			if trim(rsMH2("STOREANDSENDMAILRETURNDATE"))<>"" and not isnull(rsMH2("STOREANDSENDMAILRETURNDATE")) then
				response.write gArrDT(trim(rsMH2("STOREANDSENDMAILRETURNDATE")))&"&nbsp;,&nbsp;"
			else
				response.write "&nbsp;"
			end if
			'寄存送達紀錄人員
			if trim(rsMH2("STOREANDSENDRECORDMEMBERID"))<>"" and not isnull(rsMH2("STOREANDSENDRECORDMEMBERID")) then
				strSendRecordMem="select chName from MemberData where memberId="&trim(rsMH2("STOREANDSENDRECORDMEMBERID"))
				set rsSRMem=conn.execute(strSendRecordMem)
				if not rsSRMem.eof then
					response.write trim(rsSRMem("chName"))
				end if
				rsSRMem.close
				set rsSRMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'公示退件日期
			if trim(rsMH2("OPENGOVMAILRETURNDATE"))<>"" and not isnull(rsMH2("OPENGOVMAILRETURNDATE")) then
				response.write gArrDT(trim(rsMH2("OPENGOVMAILRETURNDATE")))
			else
				response.write "&nbsp;"
			end if

			%></td>
			
<%
	else
%>
			<td align="center"><%="&nbsp;"%></td>
			<td align="center"><%="&nbsp;"%></td>
			<td align="center"><%="&nbsp;"%></td>
<%
	end if
	rsMH2.close
	set rsMH2=nothing
%>
			<td align="center"><%
		If (sys_City="高雄縣" And (trim(rs1("IllegalAddress"))="鳳山市鳳頂路與田中央路口" Or trim(rs1("IllegalAddress"))="大寮鄉鳳屏路高屏大橋下橋處(往高雄)")) Or sys_City<>"高雄縣" Then
			'違規影像資料
			strImage="select FileName,SN from ProsecutionImageDetail where BillSn="&rs1("SN")
			set rsImage=conn.execute(strImage)
			if not rsImage.eof then
				ImgFile=trim(rsImage("FileName"))
				ImgSn=trim(rsImage("SN"))
%>
			<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>','<%=ImgSn%>')" <%lightbarstyle 1 %>><u><%=ImgFile%></u></a>
<%
			Else
				If sys_City="高雄縣" then
					strImage2="select b.FileName,b.Sn from ProsecutionImage a,ProsecutionImageDetail b" &_
						" where " &_
						" a.FileName=b.FileName" &_
						" and ProsecutionTime between TO_DATE('"&year(rs1("IllegalDate"))&"/"&Month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&Hour(rs1("IllegalDate"))&":"&Minute(rs1("IllegalDate"))&":00','YYYY/MM/DD/HH24/MI/SS') " &_
						" and TO_DATE('"&year(rs1("IllegalDate"))&"/"&Month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&Hour(rs1("IllegalDate"))&":"&Minute(rs1("IllegalDate"))&":59','YYYY/MM/DD/HH24/MI/SS')"
					'Location='"&trim(rs1("IllegalAddress"))&"'
					'response.write strImage2
					set rsImage2=conn.execute(strImage2)
					if not rsImage2.eof then
						While Not rsImage2.Eof
							ImgFile=trim(rsImage2("FileName"))
							ImgSn=trim(rsImage2("SN"))
	%>
				<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>','<%=ImgSn%>')" <%lightbarstyle 1 %>><u><%=ImgFile%></u></a>
	<%					
						rsImage2.MoveNext
						Wend
					else
						response.write "&nbsp;"
					end if
					rsImage2.close
					set rsImage2=Nothing
				End if
			end if
			rsImage.close
			set rsImage=nothing
		End If
			'違規影像掃描
			strScan2="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID=1 and Recordstateid=0"
			set rsScan2=conn.execute(strScan2)
			While Not rsScan2.Eof
%>
			<a title="開啟違規影像掃描檔.." href="<%=trim(rsScan2("FileName"))%>" target="_blank" <%lightbarstyle 1 %>><br><u>開啟違規影像掃描檔</u></a>
<%
			rsScan2.MoveNext
			Wend
			rsScan2.close
			set rsScan2=nothing
					
			%></td>
			<td align="center"><%
			'寄存送達證書掃描檔
			strScan="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID=0 and Recordstateid=0"
			set rsScan=conn.execute(strScan)
			if rsScan.eof then response.write "&nbsp;"
			While Not rsScan.Eof
%>
			<a title="開啟相關文件掃描檔.." href="<%=replace(trim(rsScan("FileName")),"/img/","/scanimg/")%>" target="_blank" <%lightbarstyle 1 %>><u>啟寄存送達證書掃描檔</u></a><br>
			
<%
			
			rsScan.MoveNext
			Wend
			rsScan.close
			set rsScan=nothing
			%></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" width="20%"><strong>回執聯掃描檔</strong></td>
			<td align="center" ><strong>監理站日期</strong></td>
			<td align="center"><strong>移送聯掃描檔</strong></td>
			<td align="center"><strong>第一次雙掛號寄存郵局</strong></td>
			<td align="center" ><strong>代收人</strong></td>
			<!-- <td align="center" width="20%"><strong></strong></td>
			<td align="center" width="20%"><strong></strong></td> -->
		</tr>
		<tr>
			<td align="center"><%
			'回執聯掃瞄
			strScan3="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID=2 and Recordstateid=0"
			set rsScan3=conn.execute(strScan3)
			if rsScan3.eof then response.write "&nbsp;"
			While Not rsScan3.Eof
%>
			<a title="開啟回執聯掃描檔.." href="<%=trim(rsScan3("FileName"))%>" target="_blank" <%lightbarstyle 1 %>><br><u>開啟回執聯掃描檔</u></a>
<%
			rsScan3.MoveNext
			Wend
			rsScan3.close
			set rsScan3=nothing
			%></td>
			<td align="center" ><%=SendStationDate%></td>
			<td align="center"><%
			'移送聯掃瞄檔
			strScan3="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID=3 and Recordstateid=0"
			set rsScan3=conn.execute(strScan3)
			if rsScan3.eof then response.write "&nbsp;"
			While Not rsScan3.Eof
%>
			<a title="移送聯掃描檔.." href="<%=trim(rsScan3("FileName"))%>" target="_blank" <%lightbarstyle 1 %>><br><u>開啟移送聯掃描檔</u></a>
<%
			rsScan3.MoveNext
			Wend
			rsScan3.close
			set rsScan3=nothing
			%></td>
			<td align="center"><%=FirstMailStation%></td>
			<td align="center" ><%=SignMan%></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" colspan="5"><strong>相關文件掃描檔</strong></td>
		<tr>
			<td align="center" colspan="5"><%
			'相關文件
			OpenGovBatchNumber=""
			strOpenGov="select BatchNumber from Dcilog where billno='"&trim(rs1("BillNo"))&"' and ExchangeTypeID='N' and ReturnMarkType='5'"
			Set rsOpenGov=conn.execute(strOpenGov)
			If Not rsOpenGov.eof Then
				OpenGovBatchNumber=trim(rsOpenGov("BatchNumber"))
			End If
			rsOpenGov.close
			Set rsOpenGov=nothing

			strScanGov="select * from BillAttatchImage where BillNo='"&trim(OpenGovBatchNumber)&"' and TypeID in (0,1,4) and Recordstateid=0"
			set rsScanGov=conn.execute(strScanGov)
			while Not rsScanGov.eof
			%>
				<a title="開啟相關文件掃描檔.." href="<%=replace(trim(rsScanGov("FileName")),"/img/","/scanimg/")%>" target="_blank" <%lightbarstyle 1 %>><u>開啟相關文件掃描檔</u></a><br>
				<%
			rsScanGov.movenext
			wend
			rsScanGov.close
			set rsScanGov=Nothing
			
			strScan2="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID in (0,1,4) and Recordstateid=0"
			set rsScan2=conn.execute(strScan2)
			while Not rsScan2.eof
			%>
				<a title="開啟相關文件掃描檔.." href="<%=replace(trim(rsScan2("FileName")),"/img/","/scanimg/")%>" target="_blank" <%lightbarstyle 1 %>><u>開啟相關文件掃描檔</u></a><br>
				<%
			rsScan2.movenext
			wend
			rsScan2.close
			set rsScan2=nothing
			%></td>
		</tr>
<%
		strDSupd="select * from DCISTATUSUPDATE where Billsn="&Trim(rs1("Sn"))
		Set rsDSupd=conn.execute(strDSupd)
		If Not rsDSupd.eof Then
		%>
				<tr>
				<td align="center" colspan="5"><strong>
					強制入案前狀態：<%
				strDS1="select * from Dcireturnstatus where DciActionID='W' " &_
					" and DciReturn='"&Trim(rsDSupd("StatUS"))&"'"
				Set rsDS1=conn.execute(strDS1)
				If Not rsDS1.eof Then
					response.write rsDS1("StatusContent")
				End If
				rsDS1.close
				Set rsDS1=Nothing
				strDS2="select * from Dcireturnstatus where DciActionID='WE' " &_
					" and DciReturn='"&Trim(rsDSupd("DciErrorCarData"))&"'"
				Set rsDS2=conn.execute(strDS2)
				If Not rsDS2.eof Then
					response.write " "&rsDS2("StatusContent")
				End If
				rsDS2.close
				Set rsDS2=Nothing
				response.write " "&rsDSupd("RecordDate")
					%></strong>
				</td>
				</tr>
		<%
		End If
		rsDSupd.close
		Set rsDSupd=nothing
		%>			
	</table>
	<br>
	<table width='100%' border='1' cellpadding="2">
	
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>1"></a>
				<strong>舉發單基本資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>2">監理所回傳資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>6">舉發單處理紀錄</a>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFFF99" width="13%" align="right"><strong>單號</strong></td>
			<td align="left" width="20%"><%
			'單號
			if trim(rs1("BillNo"))<>"" and not isnull(rs1("BillNo")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" width="13%" align="right"><strong>舉發類別</strong></td>
			<td align="left" width="20%"><%
			'舉發類別
			if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
				strBillType="select * from DciCode where TypeID=2 and ID='"&trim(rs1("BillTypeID"))&"'"
				set rsBillType=conn.execute(strBillType)
				if not rsBillType.eof then
					response.write trim(rsBillType("Content"))
				end if
				rsBillType.close
				set rsBillType=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" width="13%" align="right"><strong>車號</strong></td>
			<td align="left" width="20%"><%
			'車號
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		</tr>
			<td bgcolor="#FFFF99" align="right"><strong>違規人姓名</strong></td>
			<td align="left"><%
			'違規人姓名
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write trim(rs1("Driver"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>違規人身份證</strong></td>
			<td align="left"><%
			'違規人身分証
			if trim(rs1("DriverID"))<>"" and not isnull(rs1("DriverID")) then
				response.write trim(rs1("DriverID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>違規人生日</strong></td>
			<td align="left"><%
			'違規人生日
			if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
				response.write gArrDT(trim(rs1("DriverBirth")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>			
			<td bgcolor="#FFFF99" align="right"><strong>違規人性別</strong></td>
			<td align="left"><%
			'違規人性別
			if trim(rs1("DriverSex"))<>"" and not isnull(rs1("DriverSex")) then
				if trim(rs1("DriverSex"))="1" then
					response.write "男"
				elseif trim(rs1("DriverSex"))="2" then
					response.write "女"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>違規人地址</strong></td>
			<td align="left" colspan="3"><%
			'違規人地址
			if trim(rs1("DriverZip"))<>"" and not isnull(rs1("DriverZip")) then
				response.write trim(rs1("DriverZip"))&"&nbsp;"
			end if
			if trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) then
				response.write trim(rs1("DriverAddress"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>簡示車種</strong></td>
			<td align="left"><%
			'簡式車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車"
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				elseif trim(rs1("CarSimpleID"))="5" then
					response.write "動力機械"
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "臨時車牌"
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "試車牌"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>輔助車種</strong></td>
			<td align="left"><%
			'輔助車種
			if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
				if trim(rs1("CarAddID"))="1" then
					response.write "大貨車"
				elseif trim(rs1("CarAddID"))="2" then
					response.write "大客車"
				elseif trim(rs1("CarAddID"))="3" then
					response.write "砂石車"
				elseif trim(rs1("CarAddID"))="4" then
					response.write "土方車"
				elseif trim(rs1("CarAddID"))="5" then
					response.write "動力機"
				elseif trim(rs1("CarAddID"))="6" then
					response.write "貨櫃"
				elseif trim(rs1("CarAddID"))="7" then
					response.write "大型重機"
				elseif trim(rs1("CarAddID"))="8" then
					response.write "拖吊"
				elseif trim(rs1("CarAddID"))="9" then
					response.write "(550cc)重機"
				elseif trim(rs1("CarAddID"))="10" then
					response.write "計程車"
				elseif trim(rs1("CarAddID"))="11" then
					response.write "危險物品"
				elseif trim(rs1("CarAddID"))="12" then
					response.write "幼兒車(課輔車)"
				elseif trim(rs1("CarAddID"))="0" then
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>違規日期</strong></td>
			<td align="left"><%
			'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>違規地點</strong></td>
			<td align="left" colspan="5"><%
			'違規地點
			if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
				response.write trim(rs1("IllegalAddressID"))&" "
			end if
			if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
				response.write trim(rs1("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>違規事實</strong></td>
			<td align="left" colspan="5"><%
			'違規事實
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				if left(trim(rs1("Rule1")),4)="2110" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end If
				Elseif left(trim(rs1("Rule1")),4)="2210" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif (trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4) And trim(rs1("CarAddID"))="0" then
						strCarImple=" and CarSimpleID in ('3','0')"
					elseif (trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4) And trim(rs1("CarAddID"))<>"0" then
						strCarImple=" and CarSimpleID in ('5','0')"
					else
						strCarImple=""
					end If
				end if
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					response.write trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))="2" then
					response.write "("&trim(rs1("Rule4"))&")"
				end if
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if left(trim(rs1("Rule2")),4)="2110" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end If
				Elseif left(trim(rs1("Rule2")),4)="2210" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif (trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4) And trim(rs1("CarAddID"))="0" then
						strCarImple2=" and CarSimpleID in ('3','0')"
					elseif (trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4) And trim(rs1("CarAddID"))<>"0" then
						strCarImple2=" and CarSimpleID in ('5','0')"
					else
						strCarImple2=""
					end If
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write "<br>"&trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if left(trim(rs1("Rule3")),4)="2110" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write "<br>"&trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))="1" then
				if left(trim(rs1("Rule4")),4)="2110" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write "<br>"&trim(rs1("Rule4"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if
			%></td>
		</tr>
		<tr>

			<td align="right" bgcolor="#FFFF99"><strong>限速、限重</strong></td>
			<td align="left"><%
			'限速、限重
			if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
				response.write trim(rs1("RuleSpeed"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>車速</strong></td>
			<td align="left"><%
			'車速
			if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
				response.write trim(rs1("IllegalSpeed"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>填單日期</strong></td>
			<td align="left"><%
			'填單日期
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>應到案日期</strong></td>
			<td align="left"><%
			'應到案日期
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>應到案處所</strong></td>
			<td align="left"><%
			'應到案處所
			if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
				strMStation="select DCIStationName from Station where StationID='"&trim(rs1("MemberStation"))&"'"
				set rsMStation=conn.execute(strMStation)
				if not rsMStation.eof then
					response.write trim(rsMStation("DCIStationName"))
				end if
				rsMStation.close
				set rsMStation=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>舉發單位</strong></td>
			<td align="left"><%
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUName=conn.execute(strUName)
				if not rsUName.eof then
					response.write trim(rsUName("UnitName"))
				end if
				rsUName.close
				set rsUName=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>舉發人</strong></td>
			<td align="left"><%
			'舉發人
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				If trim(rs1("BillMemID1"))<>"" then
					strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write "("&trim(rsMem1("LoginID"))&")"
					end if
					rsMem1.close
					set rsMem1=Nothing
				End if
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
				If trim(rs1("BillMemID2"))<>"" then
					strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
					set rsMem2=conn.execute(strMem2)
					if not rsMem2.eof then
						response.write "("&trim(rsMem2("LoginID"))&")"
					end if
					rsMem2.close
					set rsMem2=Nothing
				End If 
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
				If trim(rs1("BillMemID3"))<>"" then
					strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
					set rsMem3=conn.execute(strMem3)
					if not rsMem3.eof then
						response.write "("&trim(rsMem3("LoginID"))&")"
					end if
					rsMem3.close
					set rsMem3=Nothing
				End If 
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
				If trim(rs1("BillMemID4"))<>"" then
					strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
					set rsMem4=conn.execute(strMem4)
					if not rsMem4.eof then
						response.write "("&trim(rsMem4("LoginID"))&")"
					end if
					rsMem4.close
					set rsMem4=Nothing
				End if
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>代保管物</strong></td>
			<td align="left"><%
			'代保管物
			FastenerDetail=""
			strFas="select b.Content from BillFastenerDetail a,DCICode b where BillSN="&trim(rs1("SN"))&" and b.TypeID=6 and a.FastenerTypeID=b.ID"
			set rsFas=conn.execute(strFas)
			If Not rsFas.Bof Then
				rsFas.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsFas.Eof
				if FastenerDetail="" then
					FastenerDetail=trim(rsFas("Content"))
				else
					FastenerDetail=FastenerDetail&"、"&trim(rsFas("Content"))
				end if
			rsFas.MoveNext
			Wend
			rsFas.close
			set rsFas=nothing
				response.write FastenerDetail
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>專案</strong></td>
			<td align="left"><%
			'專案
			if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
				strProj="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"
				set rsProj=conn.execute(strProj)
				if not rsProj.eof then
					response.write trim(rsProj("Name"))
				end if
				rsProj.close
				set rsProj=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>第三責任險</strong></td>
			<td align="left"><%
			'第三責任險(0:有出示/1:未出示/2:肇事且未出示/3:逾期或未保險/4:肇事且逾期或未保險) *欄停才顯示
			if trim(rs1("Insurance"))<>"" and not isnull(rs1("Insurance")) and rs1("BillTypeID")="1" then
				if trim(rs1("Insurance"))="0" then
					response.write "有出示"
				elseif trim(rs1("Insurance"))="1" then
					response.write "未出示"
				elseif trim(rs1("Insurance"))="2" then
					response.write "肇事且未出示"
				elseif trim(rs1("Insurance"))="3" then
					response.write "逾期或未保險"
				elseif trim(rs1("Insurance"))="4" then
					response.write "肇事且逾期或未保險"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>採証工具</strong></td>
			<td align="left"><%
			'採証工具 (空:無/1:固定桿/2:雷達測速[三腳架]/3:儀器[相機]) 
			if trim(rs1("UseTool"))<>"" and not isnull(rs1("UseTool")) and trim(rs1("UseTool"))<>"0" then
				if trim(rs1("UseTool"))="1" then
					response.write "固定桿"
				elseif trim(rs1("UseTool"))="2" then
					response.write "雷達測速[三腳架]"
				elseif trim(rs1("UseTool"))="3" then
					response.write "儀器[相機]"
				elseif trim(rs1("UseTool"))="4" And sys_City="台南市" then
					response.write "車載攝影機"
				elseif trim(rs1("UseTool"))="4" And sys_City="基隆市" then
					response.write "雷射測速鎗"
				elseif trim(rs1("UseTool"))="8" then
					response.write "逕舉手開單"
				end if
			elseif trim(rs1("UseTool"))="" or trim(rs1("UseTool"))="0" or isnull(rs1("UseTool")) then
				response.write "無"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>是否郵寄</strong></td>
			<td align="left"><%
			'是否郵寄
			if trim(rs1("EquipmentID"))<>"" and not isnull(rs1("EquipmentID")) then
				if trim(rs1("EquipmentID"))="1" then
					response.write "是"
				else
					response.write "否"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>交通事故案號</strong></td>
			<td align="left"><%
			'交通事故案號
			if trim(rs1("TrafficAccidentNo"))<>"" and not isnull(rs1("TrafficAccidentNo")) then
				response.write trim(rs1("TrafficAccidentNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>交通事故種類</strong></td>
			<td align="left"><%
			'交通事故種類
			if trim(rs1("TrafficAccidentType"))<>"" and not isnull(rs1("TrafficAccidentType")) then
				response.write "A"&trim(rs1("TrafficAccidentType"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>備註</strong></td>
			<td align="left"><%
			'備註
			if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
				response.write trim(rs1("Note"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>填單人</strong></td>
			<td align="left"><%
			'填單人
			if trim(rs1("BillFiller"))<>"" and not isnull(rs1("BillFiller")) then
				response.write trim(rs1("BillFiller"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>建檔人</strong></td>
			<td align="left"><%
			'建檔人
			if trim(rs1("RecordMemberID"))<>"" and not isnull(rs1("RecordMemberID")) then
				strRMem="select ChName from MemberData where MemberID="&trim(rs1("RecordMemberID"))
				set rsRMem=conn.execute(strRMem)
				if not rsRMem.eof then
					response.write trim(rsRMem("ChName"))
				end if
				rsRMem.close
				set rsRMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>建檔日期</strong></td>
			<td align="left"><%
			'建檔日期
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))&" "
				response.write Right("00"&hour(rs1("RecordDate")),2)&":"
				response.write Right("00"&minute(rs1("RecordDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>DCI作業階段</strong></td>
			<td align="left"><%
			'DCI作業階段
			if trim(rs1("BillStatus"))<>"" and not isnull(rs1("BillStatus")) then
				if trim(rs1("BillStatus"))="0" then
					response.write "未處理"
				elseif trim(rs1("BillStatus"))="1" then
					response.write "車籍查詢"
				elseif trim(rs1("BillStatus"))="2" then
					response.write "入案"
				elseif trim(rs1("BillStatus"))="3" then
					response.write "單退"
				elseif trim(rs1("BillStatus"))="4" then
					response.write "寄存送達"
				elseif trim(rs1("BillStatus"))="5" then
					response.write "公示送達"
				elseif trim(rs1("BillStatus"))="6" then
					response.write "刪除"
				elseif trim(rs1("BillStatus"))="7" then
					response.write "收受註記"
				elseif trim(rs1("BillStatus"))="9" then
					response.write "結案"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>舉發單狀態</strong></td>	
			<td align="left"><%
			'舉發單狀態
			if trim(rs1("RecordStateID"))<>"" and not isnull(rs1("RecordStateID")) then
				if trim(rs1("RecordStateID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法條版本</strong></td>
			<td align="left"><%
			'法條版本
			if trim(rs1("RuleVer"))<>"" and not isnull(rs1("RuleVer")) then
				response.write trim(rs1("RuleVer"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>刪除原因</strong></td>
			<td align="left"><%
			'刪除原因
			strDelRea="select b.Content from BillDeleteReason a,DciCode b where a.BillSn="&trim(rs1("SN"))&" and b.TypeID=3 and a.DelReason=b.ID"
			set rsDelRea=conn.execute(strDelRea)
			if not rsDelRea.eof then
				response.write trim(rsDelRea("Content"))
			else
				response.write "&nbsp;"
			end if
			rsDelRea.close
			set rsDelRea=nothing
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>刪除人</strong></td>
			<td align="left"><%
			'刪除人
			if trim(rs1("DelMemberID"))<>"" and not isnull(rs1("DelMemberID")) then
				strDMem="select ChName from MemberData where MemberID="&trim(rs1("DelMemberID"))
				set rsDMem=conn.execute(strDMem)
				if not rsDMem.eof then
					response.write trim(rsDMem("ChName"))
				end if
				rsDMem.close
				set rsDMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>違規影像資料</strong></td>
			<td align="left"><%
		If (sys_City="高雄縣" And (trim(rs1("IllegalAddress"))="鳳山市鳳頂路與田中央路口" Or trim(rs1("IllegalAddress"))="大寮鄉鳳屏路高屏大橋下橋處(往高雄)")) Or sys_City<>"高雄縣" Then
			'違規影像資料
			strImage="select FileName,SN from ProsecutionImageDetail where BillSn="&rs1("SN")
			set rsImage=conn.execute(strImage)
			if not rsImage.eof then
				ImgFile=trim(rsImage("FileName"))
				ImgSn=trim(rsImage("SN"))
%>
			<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>','<%=ImgSn%>')" <%lightbarstyle 1 %>><u><%=ImgFile%></u></a>
<%
			Else
				If sys_City="高雄縣" then
					strImage2="select b.FileName,b.Sn from ProsecutionImage a,ProsecutionImageDetail b" &_
						" where " &_
						" a.FileName=b.FileName" &_
						" and ProsecutionTime between TO_DATE('"&year(rs1("IllegalDate"))&"/"&Month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&Hour(rs1("IllegalDate"))&":"&Minute(rs1("IllegalDate"))&":00','YYYY/MM/DD/HH24/MI/SS') " &_
						" and TO_DATE('"&year(rs1("IllegalDate"))&"/"&Month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&Hour(rs1("IllegalDate"))&":"&Minute(rs1("IllegalDate"))&":59','YYYY/MM/DD/HH24/MI/SS')"
					'Location='"&trim(rs1("IllegalAddress"))&"'
					'response.write strImage2
					set rsImage2=conn.execute(strImage2)
					if not rsImage2.eof then
						While Not rsImage2.Eof
							ImgFile=trim(rsImage2("FileName"))
							ImgSn=trim(rsImage2("SN"))
	%>
				<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>','<%=ImgSn%>')" <%lightbarstyle 1 %>><u><%=ImgFile%></u></a>
	<%					
						rsImage2.MoveNext
						Wend
					else
						response.write "&nbsp;"
					end if
					rsImage2.close
					set rsImage2=Nothing
				End If 
			end if
			rsImage.close
			set rsImage=nothing
		End if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>簽收狀況</strong></td>
			<td align="left" colspan="5"><%
			'可參考google doc "攔停 簽收 狀況 "
				if trim(rs1("SignType"))<>"" and not isnull(rs1("SignType")) then
					if rs1("SignType")="A" then response.write "簽收"
					if rs1("SignType")="U" then 
						strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="2" then response.write "拒簽已收"
							if rsR2("SignStateID")="3" then response.write "已簽拒收"							
						else 
							response.write "拒簽收"
						end if
						rsR2.close
						set rsR2=nothing																
					end if				
				else
						strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="5" then response.write "補開單"
						end if
						rsR2.close
						set rsR2=nothing															
				end if

'			if trim(rs1("SignType"))="A" then
'				response.write "簽收"
'			elseif trim(rs1("SignType"))="U" then
'				response.write "拒收"
'			else
'				response.write "&nbsp;"
'			end if
			%></td>
		</tr>
	</table>
	
	
	<div class="PageNext">&nbsp;</div>
	<table width='100%' border='1' cellpadding="2">
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>2"></a>
				<strong>監理所回傳資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>6">舉發單處理紀錄</a>
			</td>
		</tr>
<%
	strDciReturnPlus=""
	If sys_City="花蓮縣" Then
		strchkA="select * from dcilog where billsn="&trim(rs1("SN"))&" and exchangetypeid='A'"
		Set rsChkA=conn.execute(strchkA)
		If rsChkA.eof Then
			strDciReturnPlus=" and exchangetypeid<>'A'"
		End If 
		rsChkA.close
		Set rsChkA=Nothing 
	End If 
	strReturn="select * from BillBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"'" &_
		" and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) "&strDciReturnPlus&" order by DCICASEINDATE desc"
	i=0
	set rsReturn=conn.execute(strReturn)
	If Not rsReturn.Bof Then rsReturn.MoveFirst 
	While Not rsReturn.Eof
		if i=0 then
			i=i+1
			TRcolor="#FFFF99"
		else
			i=i-1
			TRcolor="#AAF2A2"
		end if
%>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsReturn("BillNo"))<>"" and not isnull(rsReturn("BillNo")) then
				response.write trim(rsReturn("BillNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車號</strong></td>
			<td align="left"><%
			'車號
			if trim(rsReturn("CarNo"))<>"" and not isnull(rsReturn("CarNo")) then
				response.write trim(rsReturn("CarNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">違反牌照稅註記</span></strong></td>
			<td align="left"><%
			'違反牌照稅註記
			if trim(rsReturn("ILLEGALLICENSEID"))<>"" and not isnull(rsReturn("ILLEGALLICENSEID")) then
				if trim(rsReturn("ILLEGALLICENSEID"))="0" then
					response.write "正常"
				else
					response.write "違反牌照稅法"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>DCI傳回車種</strong></td>
			<td align="left"><%
			'DCI傳回車種
			if trim(rsReturn("DCIRETURNCARTYPE"))<>"" and not isnull(rsReturn("DCIRETURNCARTYPE")) then
				strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rsReturn("DCIReturnCarType"))&"'"
				set rsCType=conn.execute(strCType)
				if not rsCType.eof then
					response.write trim(rsCType("Content"))
				end if
				rsCType.close
				set rsCType=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">DCI傳回應到案處所</span></strong></td>
			<td align="left"><%
			'DCI傳回應到案處所
			if trim(rsReturn("DCIRETURNSTATION"))<>"" and not isnull(rsReturn("DCIRETURNSTATION")) then
				strDciStation="select DCIStationName from Station where DCIStationID='"&trim(rsReturn("DCIRETURNSTATION"))&"'"
				set rsDciStation=conn.execute(strDciStation)
				if not rsDciStation.eof then
					response.write trim(rsDciStation("DCIStationName"))
				end if
				rsDciStation.close
				set rsDciStation=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>顏色</strong></td>
			<td align="left"><%
			'顏色
			if trim(rsReturn("DCIReturnCarColor"))<>"" and not isnull(rsReturn("DCIReturnCarColor")) then
				ColorLen=cint(Len(rsReturn("DCIReturnCarColor")))
				for Clen=1 to ColorLen
					colorID=mid(rsReturn("DCIReturnCarColor"),Clen,1)
					strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
					set rsColor=conn.execute(strColor)
					if not rsColor.eof then
						response.write trim(rsColor("Content"))
					else
						response.write "&nbsp;"
					end if
					rsColor.close
					set rsColor=nothing
				next
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車籍現況</strong></td>
			<td align="left"><%
			'車籍現況
			if trim(rsReturn("DCIRETURNCARSTATUS"))<>"" and not isnull(rsReturn("DCIRETURNCARSTATUS")) then
				strCstatus="select Content from DCIcode where TypeID=10 and ID='"&trim(rsReturn("DCIReturnCarStatus"))&"'"
				set rsCS=conn.execute(strCstatus)
				if not rsCS.eof then
					response.write trim(rsCS("Content"))
				end if 
				rsCS.close
				set rsCS=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>入案日期</strong></td>
			<td align="left"><%
			'入案日期
			if trim(rsReturn("DCICASEINDATE"))<>"" and not isnull(rsReturn("DCICASEINDATE")) then
				if len(trim(rsReturn("DCICASEINDATE")))=6 then
					response.write mid(trim(rsReturn("DCICASEINDATE")),1,2)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),3,2)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),5,2)
				elseif len(trim(rsReturn("DCICASEINDATE")))=7 then
					response.write mid(trim(rsReturn("DCICASEINDATE")),1,3)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),4,2)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),6,2)
				end If
				If trim(rsReturn("ExchangeTypeID"))="W" then
					CaseInDate=trim(rsReturn("DCICASEINDATE"))
				End If 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車籍錯誤</strong></td>
			<td align="left"><%
			'車籍錯誤
			if trim(rsReturn("DCIERRORCARDATA"))<>"" and not isnull(rsReturn("DCIERRORCARDATA")) then
				strCarDateErr="select StatusContent from DciReturnStatus where DciActionID='WE' and DciReturn='"&trim(rsReturn("DCIERRORCARDATA"))&"'"
				set rsCDErr=conn.execute(strCarDateErr)
				if not rsCDErr.eof then
					response.write trim(rsCDErr("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCDErr.close
				set rsCDErr=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>駕籍錯誤</strong></td>
			<td align="left"><%
			'駕籍錯誤
			if trim(rsReturn("DCIERRORIDDATA"))<>"" and not isnull(rsReturn("DCIERRORIDDATA")) then
				strIDDateErr="select StatusContent from DciReturnStatus where DciActionID='WE' and DciReturn='"&trim(rsReturn("DCIERRORIDDATA"))&"'"
				set rsIDErr=conn.execute(strIDDateErr)
				if not rsIDErr.eof then
					response.write trim(rsIDErr("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsIDErr.close
				set rsIDErr=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>入案碼</strong></td>
			<td align="left"><%
			'入案碼
			if trim(rsReturn("DCICOUNTERID"))<>"" and not isnull(rsReturn("DCICOUNTERID")) then
				if trim(rsReturn("DCICOUNTERID"))="00" then
					response.write "未寫入資料庫"
				elseif trim(rsReturn("DCICOUNTERID"))="Y" then
					response.write "有寫入資料庫"
				elseif trim(rsReturn("DCICOUNTERID"))="N" then
					response.write "未寫入資料庫"
				elseif trim(rsReturn("DCICOUNTERID"))="S" then
					response.write "違規人已先繳結案"
				elseif trim(rsReturn("DCICOUNTERID"))="L" then
					response.write "已入案過"
				elseif trim(rsReturn("DCICOUNTERID"))="n" then
					response.write "監理單位已入案"
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>結案註記</strong></td>
			<td align="left"><%
			'結案註記
			if trim(rsReturn("BILLCLOSEID"))<>"" and not isnull(rsReturn("BILLCLOSEID")) then
				response.write trim(rsReturn("BILLCLOSEID"))

				strClose="select * from DciCode where TypeID=9 and ID='"&trim(rsReturn("BILLCLOSEID"))&"'"
				set rsClose=conn.execute(strClose)
				if not rsClose.eof then
					response.write "&nbsp;"&trim(rsClose("Content"))
				end if
				rsClose.close
				set rsClose=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>是否有保險證</strong></td>
			<td align="left" colspan="5"><%
			'是否有保險證
			if trim(rsReturn("INSURE"))<>"" and not isnull(rsReturn("INSURE")) then
				if trim(rsReturn("INSURE"))="0" then
					response.write "有正常保險證"
				elseif trim(rsReturn("INSURE"))="1" then
					response.write "未帶保險證"
				elseif trim(rsReturn("INSURE"))="2" then
					response.write "肇事且未帶保險證"
				elseif trim(rsReturn("INSURE"))="3" then
					response.write "保險證過期或未保險"
				elseif trim(rsReturn("INSURE"))="4" then
					response.write "肇事且保險證過期或未保險"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>駕駛人</strong></td>
			<td align="left"><%
			'駕駛人
			if trim(rsReturn("Driver"))<>"" and not isnull(rsReturn("Driver")) then
				response.write funcCheckFont(rsReturn("Driver"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">駕駛人出生年月日</span></strong></td>
			<td align="left"><%
			'駕駛人出生年月日
			if trim(rsReturn("DRIVERBIRTHDAY"))<>"" and not isnull(rsReturn("DRIVERBIRTHDAY")) then
				if len(trim(rsReturn("DRIVERBIRTHDAY")))=6 then
					response.write mid(trim(rsReturn("DRIVERBIRTHDAY")),1,2)
					response.write "-"&mid(trim(rsReturn("DRIVERBIRTHDAY")),3,2)
					response.write "-"&mid(trim(rsReturn("DRIVERBIRTHDAY")),5,2)
				elseif len(trim(rsReturn("DRIVERBIRTHDAY")))=6 then
					response.write mid(trim(rsReturn("DRIVERBIRTHDAY")),1,3)
					response.write "-"&mid(trim(rsReturn("DRIVERBIRTHDAY")),4,2)
					response.write "-"&mid(trim(rsReturn("DRIVERBIRTHDAY")),6,2)
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">駕駛人身分證號</span></strong></td>
			<td align="left"><%
			'駕駛人身分證號
			if trim(rsReturn("DRIVERID"))<>"" and not isnull(rsReturn("DRIVERID")) then
				response.write trim(rsReturn("DRIVERID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>駕駛人地址</strong></td>
			<td align="left" colspan="5"><%
			'駕駛人地址
			DeiverZipName=""
			if trim(rsReturn("DRIVERHomeZip"))<>"" and not isnull(rsReturn("DRIVERHomeZip")) then
				if trim(rsReturn("ExchangeTypeID"))<>"A" then
					strDZip="select ZipName from Zip where ZipID='"&trim(rsReturn("DRIVERHomeZip"))&"'"
					set rsDZip=conn.execute(strDZip)
					if not rsDZip.eof then
						DeiverZipName=trim(rsDZip("ZipName"))
					end if
					rsDZip.close
					set rsDZip=nothing
				end if
				response.write trim(rsReturn("DRIVERHomeZip"))&" "
			end if
			if trim(rsReturn("DRIVERHomeAddress"))<>"" and not isnull(rsReturn("DRIVERHomeAddress")) then
				response.write DeiverZipName&funcCheckFont(rsReturn("DRIVERHomeAddress"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車主</strong></td>
			<td align="left" colspan="3"><%
			'車主
			if trim(rsReturn("OWNER"))<>"" and not isnull(rsReturn("OWNER")) then
				response.write funcCheckFont(rsReturn("OWNER"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車主身分證</strong></td>
			<td align="left"><%
			'車主身分證
			if trim(rsReturn("OWNERID"))<>"" and not isnull(rsReturn("OWNERID")) then
				response.write trim(rsReturn("OWNERID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車主地址</strong></td>
			<td align="left" colspan="5"><%
			'車主地址
			ZipName=""
			if trim(rsReturn("OWNERZIP"))<>"" and not isnull(rsReturn("OWNERZIP")) then
				if trim(rsReturn("ExchangeTypeID"))<>"A" then
					strZip="select ZipName from Zip where ZipID='"&trim(rsReturn("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing
				end if
				response.write trim(rsReturn("OWNERZIP"))&" "
			end if
			if trim(rsReturn("OWNERADDRESS"))<>"" and not isnull(rsReturn("OWNERADDRESS")) then
				response.write ZipName&funcCheckFont(rsReturn("OWNERADDRESS"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>原車主</strong></td>
			<td align="left" colspan="3"><%
			'原車主
			if trim(rsReturn("NWNER"))<>"" and not isnull(rsReturn("NWNER")) then
				response.write funcCheckFont(rsReturn("NWNER"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>原車主身分證</strong></td>
			<td align="left"><%
			'原車主身分證
			if trim(rsReturn("NWNERID"))<>"" and not isnull(rsReturn("NWNERID")) then
				response.write trim(rsReturn("NWNERID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>原車主地址</strong></td>
			<td align="left" colspan="5"><%
			'原車主地址
			NwnZipName=""
			if trim(rsReturn("NWNERZIP"))<>"" and not isnull(rsReturn("NWNERZIP")) then
				if trim(rsReturn("ExchangeTypeID"))<>"A" then
					strNwnZip="select ZipName from Zip where ZipID='"&trim(rsReturn("NWNERZIP"))&"'"
					set rsNwnZip=conn.execute(strNwnZip)
					if not rsNwnZip.eof then
						NwnZipName=trim(rsNwnZip("ZipName"))
					end if
					rsNwnZip.close
					set rsNwnZip=nothing
				end if
				response.write trim(rsReturn("NWNERZIP"))&" "
			end if
			if trim(rsReturn("NWNERADDRESS"))<>"" and not isnull(rsReturn("NWNERADDRESS")) then
				response.write NwnZipName&funcCheckFont(rsReturn("NWNERADDRESS"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%
	If sys_City="花蓮縣" Or sys_City="高雄市" Then
		if trim(rsReturn("ExchangeTypeID"))="A" then
%>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>通訊地址</strong></td>
			<td align="left" colspan="5"><%
			'通訊地址
			NwnZipName=""
			if trim(rsReturn("OwnerNotifyAddress"))<>"" and not isnull(rsReturn("OwnerNotifyAddress")) then
				response.write funcCheckFont(rsReturn("OwnerNotifyAddress"),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%		End If 
	End If %>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車子廠牌</strong></td>
			<td align="left"><%
			'車子廠牌
			if trim(rsReturn("A_NAME"))<>"" and not isnull(rsReturn("A_NAME")) then
				response.write funcCheckFont(trim(rsReturn("A_NAME")),20,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>法條</strong></td>
			<td align="left"><%
			'法條
			IllegalRULE=""
			IllegalFORFEIT=""
			if trim(rsReturn("RULE1"))<>"" and not isnull(rsReturn("RULE1")) then
				IllegalRULE=trim(rsReturn("RULE1"))
				IllegalFORFEIT=trim(rsReturn("FORFEIT1"))
			end if
			if trim(rsReturn("RULE2"))<>"" and not isnull(rsReturn("RULE2")) then
				if IllegalRULE="" then
					IllegalRULE=trim(rsReturn("RULE2"))
					IllegalFORFEIT=trim(rsReturn("FORFEIT2"))
				else
					IllegalRULE=IllegalRULE&"/"&trim(rsReturn("RULE2"))
					IllegalFORFEIT=IllegalFORFEIT&"/"&trim(rsReturn("FORFEIT2"))
				end if
			end if
			if trim(rsReturn("RULE3"))<>"" and not isnull(rsReturn("RULE3")) then
				if IllegalRULE="" then
					IllegalRULE=trim(rsReturn("RULE3"))
					IllegalFORFEIT=trim(rsReturn("FORFEIT3"))
				else
					IllegalRULE=IllegalRULE&"/"&trim(rsReturn("RULE3"))
					IllegalFORFEIT=IllegalFORFEIT&"/"&trim(rsReturn("FORFEIT3"))
				end if
			end if
			if trim(rsReturn("RULE4"))<>"" and not isnull(rsReturn("RULE4")) then
				if IllegalRULE="" then
					IllegalRULE=trim(rsReturn("RULE4"))
					IllegalFORFEIT=trim(rsReturn("FORFEIT4"))
				else
					IllegalRULE=IllegalRULE&"/"&trim(rsReturn("RULE4"))
					IllegalFORFEIT=IllegalFORFEIT&"/"&trim(rsReturn("FORFEIT4"))
				end if
			end if
			if IllegalRULE<>"" then
				response.write IllegalRULE
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>金額</strong></td>
			<td align="left"><%
			if IllegalFORFEIT<>"" then
				response.write IllegalFORFEIT
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">DCI資料交換類型</span></strong></td>
			<td align="left"><%
			'DCI資料交換類型
			if trim(rsReturn("EXCHANGETYPEID"))<>"" and not isnull(rsReturn("EXCHANGETYPEID")) then
				if trim(rsReturn("EXCHANGETYPEID"))="A" then
					response.write "車籍查詢"
				elseif trim(rsReturn("EXCHANGETYPEID"))="W" then
					response.write "入案"
				elseif trim(rsReturn("EXCHANGETYPEID"))="N" then
					response.write "單退/寄存送達/公示送達/收受"
				elseif trim(rsReturn("EXCHANGETYPEID"))="E" then
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>回傳狀態</strong></td>
			<td align="left"><%
			'回傳狀態
			if trim(rsReturn("STATUS"))<>"" and not isnull(rsReturn("STATUS")) then
				strStuts="select StatusContent from DciReturnStatus where DciActionID='"&trim(rsReturn("EXCHANGETYPEID"))&"' and DciReturn='"&trim(rsReturn("STATUS"))&"'"
				set rsStuts=conn.execute(strStuts)
				if not rsStuts.eof then
					response.write trim(rsStuts("StatusContent"))
				end if
				rsStuts.close
				set rsStuts=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1"><!-- 送達與未送達原因 --></span></strong></td>
			<td align="left"><%
			'送達與未送達原因
			if trim(rsReturn("RCVSTS"))<>"" and not isnull(rsReturn("RCVSTS")) then
				if trim(rsReturn("RCVSTS"))="D" then
					'response.write "公示送達"
				elseif trim(rsReturn("RCVSTS"))="F" then
					'response.write "寄存送達"
				else
					strRC="select Content from DciCode where TypeID=7 and ID='"&trim(rsReturn("RCVSTS"))&"'"
					set rsRC=conn.execute(strRC)
					if not rsRC.eof then
						'response.write trim(rsRC("Content"))
					end if
					rsRC.close
					set rsRC=nothing
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr bgcolor="#FF0000">
			<td bgcolor="#CCFFCC" colspan="5"></td>
		</tr>
<%
	
	rsReturn.MoveNext
	Wend
	rsReturn.close
	set rsReturn=nothing

	DispMailHistory=0
	'smith 修改, 開放郵件歷程   20090724
	if DispMailHistory=0 then
	
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>3"></a>
				<strong>舉發單郵件歷程</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">監理所回傳資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>6">舉發單處理紀錄</a>
			</td>
		</tr>
		
<%
	
	strMailHistory="select * from BillMailHistory where BillSN="&trim(rs1("SN"))
	set rsMH=conn.execute(strMailHistory)
	if not rsMH.eof then
%>
		<tr>
			<td align="right" bgcolor="#FFFF99" width="16%"><strong>舉發單號</strong></td>
			<td align="left" width="16%"><%
			'舉發單號
			if trim(rsMH("BILLNO"))<>"" and not isnull(rsMH("BILLNO")) then
				response.write trim(rsMH("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99" width="16%"><strong>車號</strong></td>
			<td align="left" width="16%"><%
			'車號
			if trim(rsMH("CARNO"))<>"" and not isnull(rsMH("CARNO")) then
				response.write trim(rsMH("CARNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99" width="16%"><strong>大宗掛號貼條碼</strong></td>
			<td align="left" width="16%" ><%
			'大宗掛號貼條碼
			if trim(rsMH("MailNumber"))<>"" and not isnull(rsMH("MailNumber")) then
				response.write trim(rsMH("MailNumber"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>

		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>第一次郵寄日</strong></td>
			<td align="left"><%
			'第一次郵寄日
			If sys_City="苗栗縣" Then
				If CaseInDate<>"" then
					response.write gArrDT(DateAdd("d",2,(left(CaseInDate,len(CaseInDate)-4)+1911)&"/"&mid(CaseInDate,len(CaseInDate)-3,2)&"/"&mid(CaseInDate,len(CaseInDate)-1,2)))
				End if
			else
				if trim(rsMH("MAILDATE"))<>"" and not isnull(rsMH("MAILDATE")) then
					response.write gArrDT(trim(rsMH("MAILDATE")))
				else
					response.write "&nbsp;"
				end if
			End if
			
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>第二次郵寄日</strong></td>
			<td align="left">
<%			ReturnIsClose=0
		if sys_City="南投縣" then	'南投交通隊說單退結案不要顯示退件郵寄日981005
			strRChk1="select * from DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='3'" &_
			" order by ExchangeDate desc"
			set rsRChk1=conn.execute(strRChk1)
			if not rsRChk1.eof then
				if trim(rsRChk1("DciReturnStatusID"))="n" then
					ReturnIsClose=1
				end if
			end if
			rsRChk1.close
			set rsRChk1=nothing
		end if
			'寄存送達郵件日期
		if ReturnIsClose=0 then
			if trim(rsMH("STOREANDSENDSENDDATE"))<>"" and not isnull(rsMH("STOREANDSENDSENDDATE")) then
				response.write gArrDT(trim(rsMH("STOREANDSENDSENDDATE")))
			else
				response.write "&nbsp;"
			end if
		else
			response.write "&nbsp;"
		end if
			%>			
			<%
			'第二次郵寄日  --smith mark掉.  應該是 storeandsendsenddate 才對
			'if trim(rsMH("StoreAndSendMailDate"))<>"" and not isnull(rsMH("StoreAndSendMailDate")) then
			'	response.write gArrDT(trim(rsMH("StoreAndSendMailDate")))
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>最後送達狀態</strong></td>
			<td align="left"><%
			'送達狀態
			if trim(rsMH("USERMARKRESONID"))<>"" and not isnull(rsMH("USERMARKRESONID")) then
				strUserReturn="select * from DciCode where TypeID=7 and ID='"&trim(rsMH("USERMARKRESONID"))&"'"
				set rsUR=conn.execute(strUserReturn)
				if not rsUR.eof then
					response.write trim(rsUR("Content"))
					if trim(rsUR("Content"))<>"" and sys_City="南投縣" then
						if instr(trim(rs1("Note")),"退回原因：")>0 then
							response.write "("&mid(trim(rs1("Note")),instr(trim(rs1("Note")),"退回原因：")+5,4)&")"
						end if
					end if
				end if
				rsUR.close
				set rsUR=nothing
				
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><span class="style1"><strong>第一次雙掛號寄存郵局</strong></span></td>
			<td align="left"><%
			'第一次雙掛號寄存郵局
			if trim(rsMH("MailStation"))<>"" and not isnull(rsMH("MailStation")) then
				response.write trim(rsMH("MailStation"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>代收人</strong></td>
			<td align="left"><%
			'代收人
			if trim(rsMH("SignMan"))<>"" and not isnull(rsMH("SignMan")) then
				response.write trim(rsMH("SignMan"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>移送監理站日期</strong></td>
			<td align="left"><%
			'移送監理站日期
			if trim(rsMH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsMH("SendOpenGovDocToStationDate")) then
				response.write trim(rsMH("SendOpenGovDocToStationDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td colspan="6" bgcolor="#FFCCCC">第一次退件</tr>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>退件日期</strong></td>
			<td align="left"><%
			'檢查是單退還是收受
			strCheck="select count(*) as cnt from Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7'"
			set rsCheck=conn.execute(strCheck)
			if not rsCheck.eof then
				if rsCheck("cnt")="0" then
					CheckFlag=0
				else
					CheckFlag=1
				end if
			end if
			rsCheck.close
			set rsCheck=nothing
			'退件日期
			'if CheckFlag=0 then
				if trim(rsMH("MAILRETURNDATE"))<>"" and not isnull(rsMH("MAILRETURNDATE")) then
					response.write gArrDT(trim(rsMH("MAILRETURNDATE")))
				else
					response.write "&nbsp;"
				end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>退件原因</strong></td>
			<td align="left"><%
			'退件原因
			'if CheckFlag=0 then
				if trim(rsMH("RETURNRESONID"))<>"" and not isnull(rsMH("RETURNRESONID")) then
					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMH("RETURNRESONID"))&"'"
					set rsRR=conn.execute(strReturnReason)
					if not rsRR.eof then
						response.write trim(rsRR("Content"))
					end if
					rsRR.close
					set rsRR=nothing
				else
					response.write "&nbsp;"
				end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>退件註記人員</strong></td>
			<td align="left"><%
			'退件註記人員
			'if CheckFlag=0 then
				if trim(rsMH("RETURNRECORDMEMBERID"))<>"" and not isnull(rsMH("RETURNRECORDMEMBERID")) then
					strReturnRecMem="select chName from MemberData where MemberID="&trim(trim(rsMH("RETURNRECORDMEMBERID")))
					set rsRRMem=conn.execute(strReturnRecMem)
					if not rsRRMem.eof then
						response.write trim(rsRRMem("chName"))
					end if
					rsRRMem.close
					set rsRRMem=nothing
				else
					response.write "&nbsp;"
				end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<!-- <td align="right" bgcolor="#FFFF99"><strong>移送日期</strong></td>
			<td align="left"> --><%
			'移送日期
			'if CheckFlag=0 then
			'	if trim(rsMH("SENDDATE"))<>"" and not isnull(rsMH("SENDDATE")) then
			'		response.write gArrDT(trim(rsMH("SENDDATE")))
			'	else
			'		response.write "&nbsp;"
			'	end if
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
		</tr>
		<!-- <TR>
			<td align="right" bgcolor="#FFFF99"><strong>退件註記日期</strong></td>
			<td align="left" colspan="5"> --><%
			'退件記錄日期
			'if CheckFlag=0 then
			'	if trim(rsMH("RETURNRECORDDATE"))<>"" and not isnull(rsMH("RETURNRECORDDATE")) then
			'		response.write gArrDT(trim(rsMH("RETURNRECORDDATE")))
			'	else
			'		response.write "&nbsp;"
			'	end if
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td>
		</tr> -->
		<tr>
			<td colspan="6" bgcolor="#FFCCCC">寄存送達</tr>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達文號</span></strong></td>
			<td align="left"><%
			'寄存送達文號
			if trim(rsMH("STOREANDSENDGOVNUMBER"))<>"" and not isnull(rsMH("STOREANDSENDGOVNUMBER")) then
				response.write trim(rsMH("STOREANDSENDGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達單退日</span></strong></td>
			<td align="left"><%
			'寄存送達單退日
			if trim(rsMH("STOREANDSENDMAILRETURNDATE"))<>"" and not isnull(rsMH("STOREANDSENDMAILRETURNDATE")) then
				response.write gArrDT(trim(rsMH("STOREANDSENDMAILRETURNDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>寄存送達原因</strong></td>
			<td align="left"><%
			'寄存送達原因
			if trim(rsMH("STOREANDSENDRETURNRESONID"))<>"" and not isnull(rsMH("STOREANDSENDRETURNRESONID")) then
				strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMH("STOREANDSENDRETURNRESONID"))&"'"
				set rsRR=conn.execute(strReturnReason)
				if not rsRR.eof then
					response.write trim(rsRR("Content"))
				end if
				rsRR.close
				set rsRR=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>		
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達投郵日</span></strong></td>
			<td align="left"><%
			'寄存送達投郵日
			if trim(rsMH("STOREANDSENDMAILDATE"))<>"" and not isnull(rsMH("STOREANDSENDMAILDATE")) then
				response.write gArrDT(trim(rsMH("STOREANDSENDMAILDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達郵件日期</span></strong></td>	
			<td align="left"><%
			'寄存送達郵件日期
			if trim(rsMH("STOREANDSENDSENDDATE"))<>"" and not isnull(rsMH("STOREANDSENDSENDDATE")) then
				response.write gArrDT(trim(rsMH("STOREANDSENDSENDDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達生效日</span></strong></td>
			<td align="left"><%
			'寄存送達生效日
			if trim(rsMH("STOREANDSENDMailDate"))<>"" and not isnull(rsMH("STOREANDSENDMailDate")) then
				response.write gArrDT(trim(rsMH("STOREANDSENDMailDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<!-- <td align="right" bgcolor="#FFFF99"><strong>寄存送達紀錄時間</strong></td>
			<td align="left"> --><%
			'寄存送達紀錄時間
			'if trim(rsMH("STOREANDSENDRECORDDATE"))<>"" and not isnull(rsMH("STOREANDSENDRECORDDATE")) then
			'	response.write gArrDT(trim(rsMH("STOREANDSENDRECORDDATE")))&" "
			'	response.write Right("00"&hour(rsMH("STOREANDSENDRECORDDATE")),2)&":"
			'	response.write Right("00"&minute(rsMH("STOREANDSENDRECORDDATE")),2)
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達註記人員</span></strong></td>
			<td align="left" colspan="5"><%
			'寄存送達紀錄人員
			if trim(rsMH("STOREANDSENDRECORDMEMBERID"))<>"" and not isnull(rsMH("STOREANDSENDRECORDMEMBERID")) then
				strSendRecordMem="select chName from MemberData where memberId="&trim(rsMH("STOREANDSENDRECORDMEMBERID"))
				set rsSRMem=conn.execute(strSendRecordMem)
				if not rsSRMem.eof then
					response.write trim(rsSRMem("chName"))
				end if
				rsSRMem.close
				set rsSRMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td colspan="6" bgcolor="#FFCCCC">公示送達</tr>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">公示送達文號</span></strong></td>
			<td align="left"><%

			'公示送達文號
			if trim(rsMH("OPENGOVNUMBER"))<>"" and not isnull(rsMH("OPENGOVNUMBER")) then
				response.write trim(rsMH("OPENGOVNUMBER"))

			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">公示送達單退日</span></strong></td>
			<td align="left"><%
			'公示退件日期
			if trim(rsMH("OPENGOVMAILRETURNDATE"))<>"" and not isnull(rsMH("OPENGOVMAILRETURNDATE")) then
				response.write gArrDT(trim(rsMH("OPENGOVMAILRETURNDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>公示送達原因</strong></td>
			<td align="left"><%
			'公示送達原因
			if trim(rsMH("OPENGOVRESONID"))<>"" and not isnull(rsMH("OPENGOVRESONID")) then
				strGovReturn="select * from DciCode where TypeID=7 and ID='"&trim(rsMH("OPENGOVRESONID"))&"'"
				set rsGR=conn.execute(strGovReturn)
				if not rsGR.eof then
					response.write trim(rsGR("Content"))
				end if
				rsGR.close
				set rsGR=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>公告日期</strong></td>
			<td align="left"><%
			'公告日期
			if trim(rsMH("OPENGOVDATE"))<>"" and not isnull(rsMH("OPENGOVDATE")) then
				response.write gArrDT(trim(rsMH("OPENGOVDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">公示送達生效日期</span></strong></td>
			<td align="left"><%
			'公示送達生效日期
			if trim(rsMH("OPENGOVEFFECTDATE"))<>"" and not isnull(rsMH("OPENGOVEFFECTDATE")) then
				response.write gArrDT(trim(rsMH("OPENGOVEFFECTDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>

			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">公示送達監理所</span></strong></td>
			<td align="left"><%
			'公示送達監理所
			if trim(rsMH("OPENGOVSTATIONID"))<>"" and not isnull(rsMH("OPENGOVSTATIONID")) then
				strGovStation="select DciStationName from Station where DciStationID='"&trim(rsMH("OPENGOVSTATIONID"))&"'"
				set rsGS=conn.execute(strGovStation)
				if not rsGS.eof then
					response.write trim(rsGS("DciStationName"))
				end if
				rsGS.close
				set rsGS=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
		</tr>
		<!-- <tr> -->
			<!-- <td align="right" bgcolor="#FFFF99"><strong>公示刊載日期</strong></td>
			<td align="left"> --><%
			'公示刊載日期
			'if trim(rsMH("OPENGOVSENDDATE"))<>"" and not isnull(rsMH("OPENGOVSENDDATE")) then
			'	response.write gArrDT(trim(rsMH("OPENGOVSENDDATE")))
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
			<!-- <td align="right" bgcolor="#FFFF99"><strong>公示紀錄時間</strong></td>
			<td align="left"> --><%
			'公示紀錄時間
			'if trim(rsMH("OPENGOVDATE"))<>"" and not isnull(rsMH("OPENGOVDATE")) then
			'	response.write gArrDT(trim(rsMH("OPENGOVDATE")))&" "
			'	response.write Right("00"&hour(rsMH("OPENGOVDATE")),2)&":"
			'	response.write Right("00"&minute(rsMH("OPENGOVDATE")),2)
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
			<!-- <td align="right" bgcolor="#FFFF99"><strong>公示送達報表</strong></td>
			<td align="left"> --><%
			'公示送達報表
			'if trim(rsMH("OPENGOVREPORTNUMBER"))<>"" and not isnull(rsMH("OPENGOVREPORTNUMBER")) then
			'	response.write trim(rsMH("OPENGOVREPORTNUMBER"))
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
		<!-- </tr> -->
		<tr>
			<td colspan="6" bgcolor="#FFCCCC">收受註記</tr>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>收受日期</strong></td>
			<td align="left"><%
			'收受日期
			'if CheckFlag=1 then
				if trim(rsMH("SIGNDATE"))<>"" and not isnull(rsMH("SIGNDATE")) then
					response.write gArrDT(trim(rsMH("SIGNDATE")))
				else
					response.write "&nbsp;"
				end If
'				if trim(rsMH("MAILRETURNDATE"))<>"" and not isnull(rsMH("MAILRETURNDATE")) then
'					response.write gArrDT(trim(rsMH("MAILRETURNDATE")))
'				else
'					response.write "&nbsp;"
'				end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>收受原因</strong></td>
			<td align="left"><%
			'收受原因
			'if CheckFlag=1 then
				if trim(rsMH("SIGNRESONID"))<>"" and not isnull(rsMH("SIGNRESONID")) then
					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMH("SIGNRESONID"))&"'"
					set rsRR=conn.execute(strReturnReason)
					if not rsRR.eof then
						response.write trim(rsRR("Content"))
					end if
					rsRR.close
					set rsRR=nothing
				else
					response.write "&nbsp;"
				end If
'				if trim(rsMH("RETURNRESONID"))<>"" and not isnull(rsMH("RETURNRESONID")) then
'					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(rsMH("RETURNRESONID"))&"'"
'					set rsRR=conn.execute(strReturnReason)
'					if not rsRR.eof then
'						response.write trim(rsRR("Content"))
'					end if
'					rsRR.close
'					set rsRR=nothing
'				else
'					response.write "&nbsp;"
'				end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>收受註記人員</strong></td>
			<td align="left"><%
			'收受註記人員
			'if CheckFlag=1 Then
				if trim(rsMH("SIGNRECORDMEMBERID"))<>"" and not isnull(rsMH("SIGNRECORDMEMBERID")) then
					strReturnRecMem="select chName from MemberData where MemberID="&trim(trim(rsMH("SIGNRECORDMEMBERID")))
					set rsRRMem=conn.execute(strReturnRecMem)
					if not rsRRMem.eof then
						response.write trim(rsRRMem("chName"))
					end if
					rsRRMem.close
					set rsRRMem=nothing
				else
					response.write "&nbsp;"
				end if
				'if trim(rsMH("RETURNRECORDMEMBERID"))<>"" and not isnull(rsMH("RETURNRECORDMEMBERID")) then
					'strReturnRecMem="select chName from MemberData where MemberID="&trim(trim(rsMH("RETURNRECORDMEMBERID")))
					'set rsRRMem=conn.execute(strReturnRecMem)
					'if not rsRRMem.eof then
						'response.write trim(rsRRMem("chName"))
					'end if
					'rsRRMem.close
					'set rsRRMem=nothing
				'else
					'response.write "&nbsp;"
				'end if
			'else
			'	response.write "&nbsp;"
			'end if
			%></td>
			<!-- <td align="right" bgcolor="#FFFF99"><strong>收受移送日期</strong></td>
			<td align="left"> --><%
			'收受移送日期
			'if CheckFlag=1 then
			'	if trim(rsMH("SENDDATE"))<>"" and not isnull(rsMH("SENDDATE")) then
			'		response.write gArrDT(trim(rsMH("SENDDATE")))
			'	else
			'		response.write "&nbsp;"
			'	end if
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td> -->
		</tr>
		<!-- <TR>
			<td align="right" bgcolor="#FFFF99"><strong>收受註記日期</strong></td>
			<td align="left" colspan="5"> --><%
			'收受記錄日期
			'if CheckFlag=1 then
			'	if trim(rsMH("RETURNRECORDDATE"))<>"" and not isnull(rsMH("RETURNRECORDDATE")) then
			'		response.write gArrDT(trim(rsMH("RETURNRECORDDATE")))
			'	else
			'		response.write "&nbsp;"
			'	end if
			'else
			'	response.write "&nbsp;"
			'end if
			%><!-- </td>
		</tr> -->


<%		end if
		rsMH.close
		set rsMH=nothing
	
	end if
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>6"></a>
				<strong>舉發單處理紀錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">監理所回傳資料</a>
			</td>
		</tr>
<%
	strDCILog="select * from DciLog where BillSN="&trim(rs1("SN"))&" order by ExchangeDate Desc"
	i=0
	set rsLog=conn.execute(strDCILog)
	If Not rsLog.Bof Then rsLog.MoveFirst 
	While Not rsLog.Eof
		if i=0 then
			i=i+1
			TRcolor="#FFFF99"
		else
			i=i-1
			TRcolor="#AAF2A2"
		end if
%>

		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><span class="style1">DCI資料交換型態</span></td>
			<td align="left"><%
			'DCI資料交換型態
			if trim(rsLog("EXCHANGETYPEID"))<>"" and not isnull(rsLog("EXCHANGETYPEID")) then
				if trim(rsLog("EXCHANGETYPEID"))="A" then
					response.write "<b>車籍查詢</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="W" then
					response.write "<b>入案</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="Y" then
					response.write "<b>撤銷送達</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="3" then
					response.write "<b>單退</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="4" then
					response.write "<b>寄存送達</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="5" then
					response.write "<b>公示送達</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="7" then
					response.write "<b>收受</b>"
				elseif trim(rsLog("EXCHANGETYPEID"))="E" then
					response.write "<b>刪除</b>"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>">交換日期時間</td>
			<td align="left"><%
			'交換日期時間
			if trim(rsLog("EXCHANGEDATE"))<>"" and not isnull(rsLog("EXCHANGEDATE")) then
				response.write gArrDT(trim(rsLog("EXCHANGEDATE")))&" "
				response.write Right("00"&hour(rsLog("EXCHANGEDATE")),2)&":"
				response.write Right("00"&minute(rsLog("EXCHANGEDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>		
			<td align="right" bgcolor="<%=TRcolor%>">上傳人員</td>
			<td align="left"><%
			'建檔人員
			if trim(rsLog("RECORDMEMBERID"))<>"" and not isnull(rsLog("RECORDMEMBERID")) then
				strRecordMem="select chName from MemberData where memberId="&trim(rsLog("RECORDMEMBERID"))
				set rsMem=conn.execute(strRecordMem)
				if not rsMem.eof then
					response.write trim(rsMem("chName"))
				end if
				rsMem.close
				set rsMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>	
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>">作業代碼</td>
			<td align="left"><%
			'異常代碼
			if trim(rsLog("DCIRETURNSTATUSID"))<>"" and not isnull(rsLog("DCIRETURNSTATUSID")) then
				strDCIStatus="select StatusContent from DCIReturnStatus where DciActionID='"&trim(rsLog("EXCHANGETYPEID"))&"' and DciReturn='"&trim(rsLog("DCIRETURNSTATUSID"))&"'"
				set rsDStatus=conn.execute(strDCIStatus)
				if not rsDStatus.eof then
					response.write trim(rsDStatus("StatusContent"))
				end if
				rsDStatus.close
				set rsDStatus=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>">
			<!--
			<strong><%
			'DCI資料交換型態
			if trim(rsLog("EXCHANGETYPEID"))<>"" and not isnull(rsLog("EXCHANGETYPEID")) then
				if trim(rsLog("EXCHANGETYPEID"))="A" then
					response.write "車籍查詢"
				elseif trim(rsLog("EXCHANGETYPEID"))="W" then
					response.write "入案"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" then
					'單退型態
					if trim(rsLog("ReturnMarkType"))<>"" and not isnull(rsLog("ReturnMarkType")) then
						if trim(rsLog("ReturnMarkType"))="3" then
							response.write "單退"
						elseif trim(rsLog("ReturnMarkType"))="4" then
							response.write "寄存送達"
						elseif trim(rsLog("ReturnMarkType"))="5" then
							response.write "公示送達"
						elseif trim(rsLog("ReturnMarkType"))="7" then
							response.write "收受"
						end if
					end if
				elseif trim(rsLog("EXCHANGETYPEID"))="E" then
					response.write "刪除"
				end if
			end if
			%>
			-->
			DCI檔名</td>
			<td align="left"><%
			'DCI檔名
			if trim(rsLog("FILENAME"))<>"" and not isnull(rsLog("FILENAME")) then
				response.write trim(rsLog("FILENAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>">DCI 檔序號</td>
			<td align="left"><%
			'DCI 檔序號
			if trim(rsLog("SEQNO"))<>"" and not isnull(rsLog("SEQNO")) then
				response.write trim(rsLog("SEQNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			
			<td align="right" bgcolor="<%=TRcolor%>">作業批號</td>
			<td align="left" colspan="5"><%
			'作業批號
			if trim(rsLog("BatchNumber"))<>"" and not isnull(rsLog("BatchNumber")) then
				response.write trim(rsLog("BatchNumber"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td></td>
		</tr>
<%
	rsLog.MoveNext
	Wend
	rsLog.close
	set rsLog=nothing
%>
	</table>
	<br>
<%	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>
<%
conn.close
set conn=nothing
%>
<center>
<input type="button" value="列印" onclick="DP();">
<br>
(若無列印鈕，可按下滑鼠右鍵選擇列印功能)
</center>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgFileName,ImgSN){
	urlstr='../ProsecutionImage/ProsecutionImageDetail.asp?FileName='+ImgFileName.replace(/\+/g,'@2@')+'&SN='+ImgSN;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
function DP(){
	window.focus();
	window.print();
}
</script>
</html>
