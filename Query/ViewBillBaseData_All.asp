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
<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<%	'車輛
	strSQLTemp1=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp1=" where BillNO='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("CarNo"))<>"" then
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and CarNo like '%"&trim(request("CarNo"))&"%'"
		else
			strSQLTemp1=" where CarNo like '%"&trim(request("CarNo"))&"%'"
		end if
	end if
	if trim(request("illFID"))<>"" then
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and DriverID='"&trim(request("illFID"))&"'"
		else
			strSQLTemp1=" where DriverID='"&trim(request("illFID"))&"'"
		end if
	end if
	if trim(request("illName"))<>"" then
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and Driver='"&trim(request("illName"))&"'"
		else
			strSQLTemp1=" where Driver='"&trim(request("illName"))&"'"
		end if
	end if
	if trim(request("IllegalDate"))<>"" and trim(request("IllegalDate1"))<>"" then
		RecordDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
		RecordDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strSQLTemp1=" where IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
	end if
	if trim(request("BillSn"))<>"" then
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and SN='"&trim(request("BillSn"))&"'"
		else
			strSQLTemp1=" where SN='"&trim(request("BillSn"))&"'"
		end if
	end if
	if trim(request("MailNo"))<>"" then
		if strSQLTemp1<>"" then
			strSQLTemp1=strSQLTemp1&" and MailNumber='"&trim(request("MailNo"))&"'"
		else
			strSQLTemp1=" where MailNumber='"&trim(request("MailNo"))&"'"
		end if
		strSQL1="select a.* from BillBase a,BillMailHistory b"&strSQLTemp1&" and a.SN=b.BillSN"
	else
		strSQL1="select * from BillBase"&strSQLTemp1
	end if
	'========================================================
	'行人
	strSQLTemp2=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp2=" where BillNO='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("illFID"))<>"" then
		if strSQLTemp2<>"" then
			strSQLTemp2=strSQLTemp2&" and DriverID='"&trim(request("illFID"))&"'"
		else
			strSQLTemp2=" where DriverID='"&trim(request("illFID"))&"'"
		end if
	end if
	if trim(request("illName"))<>"" then
		if strSQLTemp2<>"" then
			strSQLTemp2=strSQLTemp2&" and Driver='"&trim(request("illName"))&"'"
		else
			strSQLTemp2=" where Driver='"&trim(request("illName"))&"'"
		end if
	end if
	if trim(request("BillSn"))<>"" then
		if strSQLTemp2<>"" then
			strSQLTemp2=strSQLTemp2&" and SN='"&trim(request("BillSn"))&"'"
		else
			strSQLTemp2=" where SN='"&trim(request("BillSn"))&"'"
		end if
	end if
	if trim(request("IllegalDate"))<>"" and trim(request("IllegalDate1"))<>"" then
		RecordDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
		RecordDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
		if strSQLTemp2<>"" then
			strSQLTemp2=strSQLTemp2&" and IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strSQLTemp2=" where IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
	end if
	strSQL2="select * from PasserBase"&strSQLTemp2
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	CheckSelectData1="1"
	set rs1=conn.execute(strSQL1)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
		'是否查的到資料
		CheckSelectData1="0"
	end if
	While Not rs1.Eof
%>
	<table width='100%' border='1' cellpadding="2">
		<tr bgcolor="#FFCC33">
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
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
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
			if not rsCaseIn.eof then
				response.write mid(trim(rsCaseIn("DCICASEINDATE")),1,2)
				response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),3,2)
				response.write "-"&mid(trim(rsCaseIn("DCICASEINDATE")),5,2)
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
			if not rsDelLog.eof then
				response.write trim(rsDelLog("ActionDate"))
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
			<td align="center" width="20%"><strong>寄存送達證書掃瞄檔</strong></td>
		</tr>
		<tr>
<%
	strMailHistory2="select * from BillMailHistory where BillSN="&trim(rs1("SN"))
	set rsMH2=conn.execute(strMailHistory2)
	if not rsMH2.eof then
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
			'違規影像資料
			if trim(rs1("ImageFileName"))<>"" and not isnull(rs1("ImageFileName")) then
				if instr(trim(rs1("ImagePathName")),"Type3")<>0 then
					ImgFile=left(right(trim(rs1("ImagePathName")),15),14)&replace(trim(rs1("ImageFileName")),".jpg","")
				else
					if instr(trim(rs1("ImageFileName")),"a.jpg")<>0 then
						ImgFile=replace(trim(rs1("ImageFileName")),"a.jpg","")
					elseif instr(trim(rs1("ImageFileName")),"b.jpg")<>0 then
						ImgFile=replace(trim(rs1("ImageFileName")),"b.jpg","")
					else
						ImgFile=replace(trim(rs1("ImageFileName")),".jpg","")
					end if
				end if
%>
			<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>')" <%lightbarstyle 1 %>><u><%=trim(rs1("ImageFileName"))%></u></a>
<%
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="center"><%
			'寄存送達證書掃瞄檔
			strScan="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"'"
			set rsScan=conn.execute(strScan)
			if not rsScan.eof then
%>
			<a title="開啟寄存送達證書掃瞄檔.." href="<%=trim(rsScan("FileName"))%>" target="_blank" <%lightbarstyle 1 %>><u>開啟寄存送達證書掃瞄檔</u></a>
<%
			else
				response.write "&nbsp;"
			end if
			rsScan.close
			set rsScan=nothing
			%></td>
		</tr>
		<tr bgcolor="#33FFCC">
			<td align="center" ><strong>移送堅理站日期</strong></td>
			<td align="center" colspan="2"><strong>第一次雙掛號寄存郵局</strong></td>
			<td align="center" colspan="2"><strong>代收人</strong></td>
			<!-- <td align="center" width="20%"><strong></strong></td>
			<td align="center" width="20%"><strong></strong></td> -->
		</tr>
		<tr>
			<td align="center" ><%=SendStationDate%></td>
			<td align="center" colspan="2"><%=FirstMailStation%></td>
			<td align="center" colspan="2"><%=SignMan%></td>
		</tr>
	</table>
	<table width='100%' border='1' cellpadding="2">
		<tr bgcolor="#FFCC33">
			<td colspan="6"><strong>舉發單資料</strong>
				<a href="#<%=trim(rs1("SN"))%>1">
				
			</td>
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>1"></a>
				<strong>舉發單基本資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>2">監理所回傳資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">舉發單郵件歷程</a>‧
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
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "臨時車牌"
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
				if left(trim(rs1("Rule1")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end if
				end if
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple
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
				if left(trim(rs1("Rule2")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write "<br>"&trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if left(trim(rs1("Rule3")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
				set rsR2=conn.execute(strR2)
				if not rsR2.eof then 
					response.write "<br>"&trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
				end if
				rsR2.close
				set rsR2=nothing
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))="1" then
				if left(trim(rs1("Rule4")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 or trim(rs1("CarSimpleID"))=6 then
						strCarImple2=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple2=" and CarSimpleID in ('3','0')"
					else
						strCarImple2=""
					end if
				end if
				strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2
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
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
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
			'第三責任險(0:有出示/1:未出示/2:肇事且未出示/3:逾期或未保險/4:肇事且逾期或未保險) 
			if trim(rs1("Insurance"))<>"" and not isnull(rs1("Insurance")) then
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
			'違規影像資料
			if trim(rs1("ImageFileName"))<>"" and not isnull(rs1("ImageFileName")) then
				if instr(trim(rs1("ImagePathName")),"Type3")<>0 then
					ImgFile=left(right(trim(rs1("ImagePathName")),15),14)&replace(trim(rs1("ImageFileName")),".jpg","")
				else
					if instr(trim(rs1("ImageFileName")),"a.jpg")<>0 then
						ImgFile=replace(trim(rs1("ImageFileName")),"a.jpg","")
					elseif instr(trim(rs1("ImageFileName")),"b.jpg")<>0 then
						ImgFile=replace(trim(rs1("ImageFileName")),"b.jpg","")
					else
						ImgFile=replace(trim(rs1("ImageFileName")),".jpg","")
					end if
				end if
%>
			<a title="開啟違規影像資料.." onclick="OpenImageWin('<%=ImgFile%>')" <%lightbarstyle 1 %>><u><%=trim(rs1("ImageFileName"))%></u></a>
<%
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>簽收狀況</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("SignType"))="A" then
				response.write "簽收"
			elseif trim(rs1("SignType"))="U" then
				response.write "拒收"
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>2"></a>
				<strong>監理所回傳資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">舉發單郵件歷程</a>‧
				<a href="#<%=trim(rs1("SN"))%>6">舉發單處理紀錄</a>
			</td>
		</tr>
<%
	strReturn="select * from BillBaseDCIReturn where (BillNo='"&trim(rs1("BillNo"))&"'" &_
		" and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"') order by DCICASEINDATE desc"
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
				elseif len(trim(rsReturn("DCICASEINDATE")))=6 then
					response.write mid(trim(rsReturn("DCICASEINDATE")),1,3)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),4,2)
					response.write "-"&mid(trim(rsReturn("DCICASEINDATE")),6,2)
				end if
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
					response.write "以入案過"
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
				response.write trim(rsReturn("Driver"))
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong>駕駛人身分證號</strong></td>
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
			if trim(rsReturn("DRIVERHomeZip"))<>"" and not isnull(rsReturn("DRIVERHomeZip")) then
				strDZip="select ZipName from Zip where ZipID='"&trim(rsReturn("DRIVERHomeZip"))&"'"
				set rsDZip=conn.execute(strDZip)
				if not rsDZip.eof then
					DeiverZipName=trim(rsDZip("ZipName"))
				end if
				rsDZip.close
				set rsDZip=nothing
				response.write trim(rsReturn("DRIVERHomeZip"))&" "
			end if
			if trim(rsReturn("DRIVERHomeAddress"))<>"" and not isnull(rsReturn("DRIVERHomeAddress")) then
				response.write DeiverZipName&trim(rsReturn("DRIVERHomeAddress"))
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
				response.write trim(rsReturn("OWNER"))
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
			if trim(rsReturn("OWNERZIP"))<>"" and not isnull(rsReturn("OWNERZIP")) then
				strZip="select ZipName from Zip where ZipID='"&trim(rsReturn("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
				response.write trim(rsReturn("OWNERZIP"))&" "
			end if
			if trim(rsReturn("OWNERADDRESS"))<>"" and not isnull(rsReturn("OWNERADDRESS")) then
				response.write ZipName&trim(rsReturn("OWNERADDRESS"))
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
				response.write trim(rsReturn("NWNER"))
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
			if trim(rsReturn("NWNERZIP"))<>"" and not isnull(rsReturn("NWNERZIP")) then
				strNwnZip="select ZipName from Zip where ZipID='"&trim(rsReturn("NWNERZIP"))&"'"
				set rsNwnZip=conn.execute(strNwnZip)
				if not rsNwnZip.eof then
					NwnZipName=trim(rsNwnZip("ZipName"))
				end if
				rsNwnZip.close
				set rsNwnZip=nothing
				response.write trim(rsReturn("NWNERZIP"))&" "
			end if
			if trim(rsReturn("NWNERADDRESS"))<>"" and not isnull(rsReturn("NWNERADDRESS")) then
				response.write NwnZipName&trim(rsReturn("NWNERADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>車子廠牌</strong></td>
			<td align="left"><%
			'車子廠牌
			if trim(rsReturn("A_NAME"))<>"" and not isnull(rsReturn("A_NAME")) then
				response.write trim(rsReturn("A_NAME"))
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">送達與未送達原因</span></strong></td>
			<td align="left"><%
			'送達與未送達原因
			if trim(rsReturn("RCVSTS"))<>"" and not isnull(rsReturn("RCVSTS")) then
				if trim(rsReturn("RCVSTS"))="D" then
					response.write "公示送達"
				elseif trim(rsReturn("RCVSTS"))="F" then
					response.write "寄存送達"
				else
					strRC="select Content from DciCode where TypeID=7 and ID='"&trim(rsReturn("RCVSTS"))&"'"
					set rsRC=conn.execute(strRC)
					if not rsRC.eof then
						response.write trim(rsRC("Content"))
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
			<td align="right" bgcolor="#FFFF99"><strong>舉發單號</strong></td>
			<td align="left"><%
			'舉發單號
			if trim(rsMH("BILLNO"))<>"" and not isnull(rsMH("BILLNO")) then
				response.write trim(rsMH("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>車號</strong></td>
			<td align="left"><%
			'車號
			if trim(rsMH("CARNO"))<>"" and not isnull(rsMH("CARNO")) then
				response.write trim(rsMH("CARNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>大宗掛號貼條碼</strong></td>
			<td align="left" ><%
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
			if trim(rsMH("MAILDATE"))<>"" and not isnull(rsMH("MAILDATE")) then
				response.write gArrDT(trim(rsMH("MAILDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>第二次郵寄日</strong></td>
			<td align="left"><%
			'第二次郵寄日
			if trim(rsMH("StoreAndSendMailDate"))<>"" and not isnull(rsMH("StoreAndSendMailDate")) then
				response.write gArrDT(trim(rsMH("StoreAndSendMailDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>最後送達狀態</strong></td>
			<td align="left"><%
			'送達狀態
			if trim(rsMH("USERMARKRESONID"))<>"" and not isnull(rsMH("USERMARKRESONID")) then
				strUserReturn="select * from DciCode where TypeID=7 and ID='"&trim(rsMH("USERMARKRESONID"))&"'"
				set rsUR=conn.execute(strUserReturn)
				if not rsUR.eof then
					response.write trim(rsUR("Content"))
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
			<td align="right" bgcolor="#FFFF99"><strong>移送堅理站日期</strong></td>
			<td align="left"><%
			'移送堅理站日期
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
			if CheckFlag=0 then
				if trim(rsMH("MAILRETURNDATE"))<>"" and not isnull(rsMH("MAILRETURNDATE")) then
					response.write gArrDT(trim(rsMH("MAILRETURNDATE")))
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>退件原因</strong></td>
			<td align="left"><%
			'退件原因
			if CheckFlag=0 then
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
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>退件註記人員</strong></td>
			<td align="left"><%
			'退件註記人員
			if CheckFlag=0 then
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
			else
				response.write "&nbsp;"
			end if
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
			<td align="right" bgcolor="#FFFF99"><strong><span class="style1">寄存送達原因</span></strong></td>
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
			if CheckFlag=1 then
				if trim(rsMH("MAILRETURNDATE"))<>"" and not isnull(rsMH("MAILRETURNDATE")) then
					response.write gArrDT(trim(rsMH("MAILRETURNDATE")))
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>收受原因</strong></td>
			<td align="left"><%
			'收受原因
			if CheckFlag=1 then
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
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>收受註記人員</strong></td>
			<td align="left"><%
			'收受註記人員
			if CheckFlag=1 then
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
			else
				response.write "&nbsp;"
			end if
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


<%	end if
	rsMH.close
	set rsMH=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>6"></a>
				<strong>舉發單處理紀錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">監理所回傳資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">舉發單郵件歷程</a>
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong>上傳人員</strong></td>
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong>交換日期時間</strong></td>
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong><span class="style1">DCI資料交換型態</span></strong></td>
			<td align="left"><%
			'DCI資料交換型態
			if trim(rsLog("EXCHANGETYPEID"))<>"" and not isnull(rsLog("EXCHANGETYPEID")) then
				if trim(rsLog("EXCHANGETYPEID"))="A" then
					response.write "車籍查詢"
				elseif trim(rsLog("EXCHANGETYPEID"))="W" then
					response.write "入案"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))="Y" then
					response.write "撤銷送達"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" and trim(rsLog("ReturnMarkType"))<>"Y" then
					response.write "單退/寄存送達/公示送達/收受"
				elseif trim(rsLog("EXCHANGETYPEID"))="E" then
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>單退型態</strong></td>
			<td align="left"><%
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
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong><%
			'DCI資料交換型態
			if trim(rsLog("EXCHANGETYPEID"))<>"" and not isnull(rsLog("EXCHANGETYPEID")) then
				if trim(rsLog("EXCHANGETYPEID"))="A" then
					response.write "車籍查詢"
				elseif trim(rsLog("EXCHANGETYPEID"))="W" then
					response.write "入案"
				elseif trim(rsLog("EXCHANGETYPEID"))="N" then
					if trim(rsLog("ReturnMarkType"))<>"" and not isnull(rsLog("ReturnMarkType")) then
						if trim(rsLog("ReturnMarkType"))="3" then
							response.write "單退"
						elseif trim(rsLog("ReturnMarkType"))="4" then
							response.write "寄存送達"
						elseif trim(rsLog("ReturnMarkType"))="5" then
							response.write "公示送達"
						elseif trim(rsLog("ReturnMarkType"))="7" then
							response.write "收受"
						elseif trim(rsLog("ReturnMarkType"))="Y" then
							response.write "撤銷送達"
						end if
					end if
				elseif trim(rsLog("EXCHANGETYPEID"))="E" then
					response.write "刪除"
				end if
			end if
			%>DCI檔名</strong></td>
			<td align="left"><%
			'DCI檔名
			if trim(rsLog("FILENAME"))<>"" and not isnull(rsLog("FILENAME")) then
				response.write trim(rsLog("FILENAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>DCI 檔序號</strong></td>
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
			<td align="right" bgcolor="<%=TRcolor%>"><strong>異常代碼</strong></td>
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
			%></td><td align="right" bgcolor="<%=TRcolor%>"><strong>作業批號</strong></td>
			<td align="left" colspan="3"><%
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

	'====================================================================================
%>
<%	CheckSelectData2="1"
'如果沒有查詢條件則不做任何動作
if strSQLTemp2="" then
	CheckSelectData2="0"
else
	set rs1=conn.execute(strSQL2)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
		'是否查的到資料
		CheckSelectData2="0"
	end if
	While Not rs1.Eof
%>
	<table width='100%' border='1' cellpadding="1">
		<tr bgcolor="#FFCC33">
			<td colspan="6"><strong>舉發單資料</strong>
				<a href="#<%=trim(rs1("SN"))%>1">
				
			</td>
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>1"></a>
				<strong>舉發單基本資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
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
			<td align="left" colspan="3"><%
			'舉發類別
			if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
				if trim(rs1("BillTypeID"))="1" then
					response.write "慢車"
				elseif trim(rs1("BillTypeID"))="2" then
					response.write "行人"
				elseif trim(rs1("BillTypeID"))="3" then
					response.write "道路障礙"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td bgcolor="#FFFF99" width="13%" align="right"><strong>違規人姓名</strong></td>
			<td align="left" width="20%"><%
			'違規人姓名
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write trim(rs1("Driver"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>違規人身份證</strong></td>
			<td align="left" width="20%"><%
			'違規人身分証
			if trim(rs1("DriverID"))<>"" and not isnull(rs1("DriverID")) then
				response.write trim(rs1("DriverID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>違規人生日</strong></td>
			<td align="left" width="20%"><%
			'違規人生日
			if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
				response.write gArrDT(trim(rs1("DriverBirth")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>			
			<td align="right" bgcolor="#FFFF99"><strong>違規人地址</strong></td>
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
			<td align="left" colspan="3"><%
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
			<td align="right" bgcolor="#FFFF99"><strong>違規法條</strong></td>
			<td align="left"><%
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
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>填單日期</strong></td>
			<td align="left"><%
			'填單日期
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
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
				strMStation="select UnitName from UnitInfo where UnitID='"&trim(rs1("MemberStation"))&"'"
				set rsMStation=conn.execute(strMStation)
				if not rsMStation.eof then
					response.write trim(rsMStation("UnitName"))
				end if
				rsMStation.close
				set rsMStation=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
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
			<td align="right" bgcolor="#FFFF99"><strong>舉發人</strong></td>
			<td align="left"><%
			'舉發人
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>代保管物</strong></td>
			<td align="left"><%
			'代保管物
			FastenerDetail=""
			strFas="select Confiscate from PasserConfiscate where BillSN="&trim(rs1("SN"))
			set rsFas=conn.execute(strFas)
			If Not rsFas.Bof Then
				rsFas.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsFas.Eof
				if FastenerDetail="" then
					FastenerDetail=trim(rsFas("Confiscate"))
				else
					FastenerDetail=FastenerDetail&"、"&trim(rsFas("Confiscate"))
				end if
			rsFas.MoveNext
			Wend
			rsFas.close
			set rsFas=nothing
				response.write FastenerDetail
			%></td>
		</tr>
		<tr>
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
			<td align="right" bgcolor="#FFFF99"><strong>備註</strong></td>
			<td align="left" colspan="3"><%
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
			<td align="right" bgcolor="#FFFF99"><strong>是否應聽講習</strong></td>
			<td align="left"><%
			'是否講習
			if trim(rs1("ISLECTURE"))<>"" and not isnull(rs1("ISLECTURE")) then
				if trim(rs1("ISLECTURE"))="0" then
					response.write "否"
				else
					response.write "是"
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
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>刪除人</strong></td>
			<td align="left" colspan="5"><%
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
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>2"></a>
				<strong>行人攤販裁決書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strJude="select * from PasserJude where BillSn="&trim(rs1("SN"))
	set rsJude=conn.execute(strJude)
	if not rsJude.eof then
%>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsJude("BILLNO"))<>"" and not isnull(rsJude("BILLNO")) then
				response.write trim(rsJude("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>發文字號</strong></td>
			<td align="left" colspan="3"><%
			'發文字號
			if trim(rsJude("OPENGOVNUMBER"))<>"" and not isnull(rsJude("OPENGOVNUMBER")) then
				response.write trim(rsJude("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>裁決日期</strong></td>
			<td align="left"><%
			'裁決日期
			if trim(rsJude("JUDEDATE"))<>"" and not isnull(rsJude("JUDEDATE")) then
				response.write gArrDT(trim(rsJude("JUDEDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>應到案處所</strong></td>
			<td align="left"><%
			'應到案處所
			if trim(rsJude("DUTYUNIT"))<>"" and not isnull(rsJude("DUTYUNIT")) then
				response.write trim(rsJude("DUTYUNIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(rsJude("FORFEIT"))<>"" and not isnull(rsJude("FORFEIT")) then
				response.write trim(rsJude("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>處罰主文</strong></td>
			<td align="left" colspan="5"><%
			'處罰主文
			if trim(rsJude("PUNISHMENTMAINBODY"))<>"" and not isnull(rsJude("PUNISHMENTMAINBODY")) then
				response.write trim(rsJude("PUNISHMENTMAINBODY"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td bgcolor="#FFFF99" align="right"><strong>簡要理由</strong></td>
			<td align="left" colspan="5"><%
			'簡要理由
			if trim(rsJude("SIMPLERESON"))<>"" and not isnull(rsJude("SIMPLERESON")) then
				response.write trim(rsJude("SIMPLERESON"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsJude("BIGUNITBOSSNAME"))<>"" and not isnull(rsJude("BIGUNITBOSSNAME")) then
				response.write trim(rsJude("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsJude("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsJude("SUBUNITSECBOSSNAME")) then
				response.write trim(rsJude("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>聯絡電話</strong></td>
			<td align="left"><%
			'聯絡電話
			if trim(rsJude("CONTACTTEL"))<>"" and not isnull(rsJude("CONTACTTEL")) then
				response.write trim(rsJude("CONTACTTEL"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人姓名</strong></td>
			<td align="left"><%
			'法定代理人姓名
			if trim(rsJude("AGENTNAME"))<>"" and not isnull(rsJude("AGENTNAME")) then
				response.write trim(rsJude("AGENTNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人生日</strong></td>
			<td align="left"><%
			'法定代理人生日
			if trim(rsJude("AGENTBIRTH"))<>"" and not isnull(rsJude("AGENTBIRTH")) then
				response.write gArrDT(trim(rsJude("AGENTBIRTH")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人身分證字號</strong></td>
			<td align="left"><%
			'法定代理人身分證字號
			if trim(rsJude("AGENTID"))<>"" and not isnull(rsJude("AGENTID")) then
				response.write trim(rsJude("AGENTID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人性別</strong></td>
			<td align="left"><%
			'法定代理人性別
			if trim(rsJude("AGENTSEX"))<>"" and not isnull(rsJude("AGENTSEX")) then
				if trim(rsJude("AGENTSEX"))="0" then
					response.write "女"
				elseif trim(rsJude("AGENTSEX"))="1" then
					response.write "男"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人住址</strong></td>
			<td align="left" colspan="3"><%
			'法定代理人住址
			if trim(rsJude("AGENTADDRESS"))<>"" and not isnull(rsJude("AGENTADDRESS")) then
				response.write trim(rsJude("AGENTADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsJude("RECORDSTATEID"))<>"" and not isnull(rsJude("RECORDSTATEID")) then
				if trim(rsJude("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsJude("RECORDDATE"))<>"" and not isnull(rsJude("RECORDDATE")) then
				response.write gArrDT(trim(rsJude("RECORDDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsJude("RECORDMEMBERID"))<>"" and not isnull(rsJude("RECORDMEMBERID")) then
				strRecordMem="select chName from MemberData where MemberId="&trim(rsJude("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecordMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>送信地址</strong></td>
			<td align="left" colspan="3"><%
			'送信地址
			if trim(rsJude("SENDADDRESS"))<>"" and not isnull(rsJude("SENDADDRESS")) then
				response.write trim(rsJude("SENDADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsJude("DELMEMBERID"))<>"" and not isnull(rsJude("DELMEMBERID")) then
				strRecordMem="select chName from MemberData where MemberId="&trim(rsJude("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecordMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>備註</strong></td>
			<td align="left" colspan="5"><%
			'備註
			if trim(rsJude("NOTE"))<>"" and not isnull(rsJude("NOTE")) then
				response.write trim(rsJude("NOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%	end if
	rsJude.close
	set rsJude=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>3"></a>
				<strong>行人攤販移送書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strPasserSend="select * from PasserSend where BillSn="&trim(rs1("SN"))
	set rsSend=conn.execute(strPasserSend)
	if not rsSend.eof then
%>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsSend("BILLNO"))<>"" and not isnull(rsSend("BILLNO")) then
				response.write trim(rsSend("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>發文文號</strong></td>
			<td align="left"><%
			'發文文號
			if trim(rsSend("OPENGOVNUMBER"))<>"" and not isnull(rsSend("OPENGOVNUMBER")) then
				response.write trim(rsSend("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>移送字號</strong></td>
			<td align="left"><%
			'移送字號
			if trim(rsSend("SENDNUMBER"))<>"" and not isnull(rsSend("SENDNUMBER")) then
				response.write trim(rsSend("SENDNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>移送日期</strong></td>
			<td align="left"><%
			'移送日期
			if trim(rsSend("SENDDATE"))<>"" and not isnull(rsSend("SENDDATE")) then
				response.write gArrDT(trim(rsSend("SENDDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人</strong></td>
			<td align="left"><%
			'法定代理人
			if trim(rsSend("AGENT"))<>"" and not isnull(rsSend("AGENT")) then
				response.write trim(rsSend("AGENT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人生日</strong></td>
			<td align="left"><%
			'法定代理人生日
			if trim(rsSend("AGENTBIRTHDATE"))<>"" and not isnull(rsSend("AGENTBIRTHDATE")) then
				response.write gArrDT(trim(rsSend("AGENTBIRTHDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人證號</strong></td>
			<td align="left"><%
			'法定代理人證號
			if trim(rsSend("AGENTID"))<>"" and not isnull(rsSend("AGENTID")) then
				response.write trim(rsSend("AGENTID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>法定代理人住址</strong></td>
			<td align="left" colspan="3"><%
			'法定代理人住址
			if trim(rsSend("AGENTADDRESS"))<>"" and not isnull(rsSend("AGENTADDRESS")) then
				response.write trim(rsSend("AGENTADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(rsSend("FORFEIT"))<>"" and not isnull(rsSend("FORFEIT")) then
				response.write trim(rsSend("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsSend("BIGUNITBOSSNAME"))<>"" and not isnull(rsSend("BIGUNITBOSSNAME")) then
				response.write trim(rsSend("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsSend("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsSend("SUBUNITSECBOSSNAME")) then
				response.write trim(rsSend("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>確定日期</strong></td>
			<td align="left"><%
			'確定日期
			if trim(rsSend("MAKESUREDATE"))<>"" and not isnull(rsSend("MAKESUREDATE")) then
				response.write gArrDT(trim(rsSend("MAKESUREDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>限繳日期</strong></td>
			<td align="left"><%
			'現繳日期
			if trim(rsSend("LIMITDATE"))<>"" and not isnull(rsSend("LIMITDATE")) then
				response.write gArrDT(trim(rsSend("LIMITDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>執行處回文日期</strong></td>
			<td align="left"><%
			'執行處回文日期
			if trim(rsSend("EXECUTERETURNDATE"))<>"" and not isnull(rsSend("EXECUTERETURNDATE")) then
				response.write gArrDT(trim(rsSend("EXECUTERETURNDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>執行處回文文號</strong></td>
			<td align="left"><%
			'執行處回文文號
			if trim(rsSend("EXECUTERETURNNUMBER"))<>"" and not isnull(rsSend("EXECUTERETURNNUMBER")) then
				response.write trim(rsSend("EXECUTERETURNNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>執行處回文登記</strong></td>
			<td align="left"><%
			'執行處回文登記
			if trim(rsSend("EXECUTERETURNNOTE"))<>"" and not isnull(rsSend("EXECUTERETURNNOTE")) then
				response.write trim(rsSend("EXECUTERETURNNOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsSend("RECORDSTATEID"))<>"" and not isnull(rsSend("RECORDSTATEID")) then
				if trim(rsSend("RECORDSTATEID"))="0" then
					response.write "正常"
				elseif trim(rsSend("RECORDSTATEID"))="-1" then
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsSend("RECORDDATE"))<>"" and not isnull(rsSend("RECORDDATE")) then
				response.write gArrDT(trim(rsSend("RECORDDATE")))&" "
				response.write Right("00"&hour(rsSend("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsSend("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsSend("RECORDMEMBERID"))<>"" and not isnull(rsSend("RECORDMEMBERID")) then
				strSRecMem="select chName from MemberData where MemberID="&trim(rsSend("RECORDMEMBERID"))
				set rsSRmem=conn.execute(strSRecMem)
				if not rsSRmem.eof then
					response.write trim(rsSRmem("chName"))
				end if
				rsSRmem.close
				set rsSRmem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsSend("DELMEMBERID"))<>"" and not isnull(rsSend("DELMEMBERID")) then
				strSRecMem="select chName from MemberData where MemberID="&trim(rsSend("DELMEMBERID"))
				set rsSRmem=conn.execute(strSRecMem)
				if not rsSRmem.eof then
					response.write trim(rsSRmem("chName"))
				end if
				rsSRmem.close
				set rsSRmem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>保全措施</strong></td>
			<td align="left" colspan="5"><%
			'保全措施
			SafeWay=""
			if trim(rsSend("SAFETOEXIT"))<>"" and not isnull(rsSend("SAFETOEXIT")) then
				if trim(rsSend("SAFETOEXIT"))="1" then
					SafeWay="已限制出境"
				end if
			end if
			if trim(rsSend("SAFEACTION"))<>"" and not isnull(rsSend("SAFEACTION")) then
				if trim(rsSend("SAFEACTION"))="1" then
					if SafeWay="" then
						SafeWay="已禁止處分"
					else
						SafeWay=SafeWay&"、已禁止處分"
					end if
				end if
			end if
			if trim(rsSend("SAFEASSURE"))<>"" and not isnull(rsSend("SAFEASSURE")) then
				if trim(rsSend("SAFEASSURE"))="1" then
					if SafeWay="" then
						SafeWay="已提供擔保"
					else
						SafeWay=SafeWay&"、已提供擔保"
					end if
				end if
			end if
			if trim(rsSend("SAFEDETAIN"))<>"" and not isnull(rsSend("SAFEDETAIN")) then
				if trim(rsSend("SAFEDETAIN"))="1" then
					if SafeWay="" then
						SafeWay="已假扣押"
					else
						SafeWay=SafeWay&"、已假扣押"
					end if
				end if
			end if
			if trim(rsSend("SAFESHUTSHOP"))<>"" and not isnull(rsSend("SAFESHUTSHOP")) then
				if trim(rsSend("SAFESHUTSHOP"))="1" then
					if SafeWay="" then
						SafeWay="已勒令停業"
					else
						SafeWay=SafeWay&"、已勒令停業"
					end if
				end if
			end if
			if SafeWay="" then
				response.write "&nbsp;"
			else
				response.write SafeWay
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>附件</strong></td>
			<td align="left" colspan="5"><%
			'附件
			SubObject=""
			if trim(rsSend("ATTATCHTABLE"))<>"" and not isnull(rsSend("ATTATCHTABLE")) then
				if trim(rsSend("ATTATCHTABLE"))="1" then
					SubObject="附表"
				end if
			end if
			if trim(rsSend("ATTATCHJUDE"))<>"" and not isnull(rsSend("ATTATCHJUDE")) then
				if trim(rsSend("ATTATCHJUDE"))="1" then
					if SubObject="" then
						SubObject="處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件"
					else
						SubObject=SubObject&"、處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHURGE"))<>"" and not isnull(rsSend("ATTATCHURGE")) then
				if trim(rsSend("ATTATCHURGE"))="1" then
					if SubObject="" then
						SubObject="義務人經限期履行而逾期能不履行之證明文件及送達證明文件"
					else
						SubObject=SubObject&"、義務人經限期履行而逾期能不履行之證明文件及送達證明文件"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHFORTUNE"))<>"" and not isnull(rsSend("ATTATCHFORTUNE")) then
				if trim(rsSend("ATTATCHFORTUNE"))="1" then
					if SubObject="" then
						SubObject="義務人之財產目錄 "
					else
						SubObject=SubObject&"、義務人之財產目錄 "
					end if
				end if
			end if
			if trim(rsSend("ATTATCHGROUND"))<>"" and not isnull(rsSend("ATTATCHGROUND")) then
				if trim(rsSend("ATTATCHGROUND"))="1" then
					if SubObject="" then
						SubObject="土地登記部謄本"
					else
						SubObject=SubObject&"、土地登記部謄本"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHREGISTER"))<>"" and not isnull(rsSend("ATTATCHREGISTER")) then
				if trim(rsSend("ATTATCHREGISTER"))="1" then
					if SubObject="" then
						SubObject="義務人之戶籍資料"
					else
						SubObject=SubObject&"、義務人之戶籍資料"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHFILELIST"))<>"" and not isnull(rsSend("ATTATCHFILELIST")) then
				if trim(rsSend("ATTATCHFILELIST"))="1" then
					if SubObject="" then
						SubObject="磁片電子檔清單"
					else
						SubObject=SubObject&"、磁片電子檔清單"
					end if
				end if
			end if
			if trim(rsSend("ATTATPOSTAGE"))<>"" and not isnull(rsSend("ATTATPOSTAGE")) then
				if trim(rsSend("ATTATPOSTAGE"))="1" then
					if SubObject="" then
						SubObject="郵資"
					else
						SubObject=SubObject&"、郵資"
					end if
				end if
			end if
			if SubObject="" then
				response.write "&nbsp;"
			else
				response.write SubObject
			end if
			%></td>
		</tr>
<%	end if
	rsSend.close
	set rsSend=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>4"></a>
				<strong>行人攤販催告書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strPasserUrge="select * from PasserUrge where BillSn="&trim(rs1("Sn"))
	set rsUrge=conn.execute(strPasserUrge)
	if not rsUrge.eof then
%>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>舉發單號</strong></td>
			<td align="left"><%
			'舉發單號
			if trim(rsUrge("BILLNO"))<>"" and not isnull(rsUrge("BILLNO")) then
				response.write trim(rsUrge("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>發文字號</strong></td>
			<td align="left" colspan="3"><%
			'發文字號
			if trim(rsUrge("OPENGOVNUMBER"))<>"" and not isnull(rsUrge("OPENGOVNUMBER")) then
				response.write trim(rsUrge("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>催告日期</strong></td>
			<td align="left"><%
			'催告日期
			if trim(rsUrge("URGEDATE"))<>"" and not isnull(rsUrge("URGEDATE")) then
				response.write gArrDT(trim(rsUrge("URGEDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>催繳方式</strong></td>
			<td align="left"><%
			'催繳方式
			if trim(rsUrge("URGETYPEID"))<>"" and not isnull(rsUrge("URGETYPEID")) then
				if trim(rsUrge("URGETYPEID"))="0" then
					response.write "電話"
				elseif trim(rsUrge("URGETYPEID"))="1" then
					response.write "信函"
				elseif trim(rsUrge("URGETYPEID"))="2" then
					response.write "雙掛號或裁決書"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(rsUrge("FORFEIT"))<>"" and not isnull(rsUrge("FORFEIT")) then
				response.write trim(rsUrge("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>寄送地址</strong></td>
			<td align="left" colspan="3"><%
			'寄送地址
			if trim(rsUrge("SENDADDRESS"))<>"" and not isnull(rsUrge("SENDADDRESS")) then
				response.write trim(rsUrge("SENDADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsUrge("BIGUNITBOSSNAME"))<>"" and not isnull(rsUrge("BIGUNITBOSSNAME")) then
				response.write trim(rsUrge("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsUrge("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsUrge("SUBUNITSECBOSSNAME")) then
				response.write trim(rsUrge("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>聯絡電話</strong></td>
			<td align="left"><%
			'聯絡電話
			if trim(rsUrge("CONTACTTEL"))<>"" and not isnull(rsUrge("CONTACTTEL")) then
				response.write trim(rsUrge("CONTACTTEL"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsUrge("RECORDSTATEID"))<>"" and not isnull(rsUrge("RECORDSTATEID")) then
				if trim(rsUrge("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsUrge("RECORDDATE"))<>"" and not isnull(rsUrge("RECORDDATE")) then
				response.write gArrDT(trim(rsUrge("RECORDDATE")))&" "
				response.write Right("00"&hour(rsUrge("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsUrge("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsUrge("RECORDMEMBERID"))<>"" and not isnull(rsUrge("RECORDMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsUrge("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsUrge("DELMEMBERID"))<>"" and not isnull(rsUrge("DELMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsUrge("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%	end if
	rsUrge.close
	set rsUrge=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF">
				<a name="#<%=trim(rs1("SN"))%>5"></a>
				<strong>行人攤販繳費記錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">舉行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%	i=0
	strPasserPay="select * from PasserPay where BillSn="&trim(rs1("SN"))&" order by RECORDDATE desc"
	set rsPay=conn.execute(strPasserPay)
	If Not rsPay.Bof Then rsPay.MoveFirst 
	While Not rsPay.Eof
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
			if trim(rsPay("BILLNO"))<>"" and not isnull(rsPay("BILLNO")) then
				response.write trim(rsPay("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>收據號碼</strong></td>
			<td align="left"><%
			'收據號碼
			if trim(rsPay("PAYNO"))<>"" and not isnull(rsPay("PAYNO")) then
				response.write trim(rsPay("PAYNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費次數</strong></td>
			<td align="left"><%
			'繳費次數
			if trim(rsPay("PAYTIMES"))<>"" and not isnull(rsPay("PAYTIMES")) then
				response.write trim(rsPay("PAYTIMES"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費方式</strong></td>
			<td align="left"><%
			'繳費方式
			if trim(rsPay("PAYTYPEID"))<>"" and not isnull(rsPay("PAYTYPEID")) then
				if trim(rsPay("PAYTYPEID"))="1" then 
					response.write "窗口"
				else
					response.write "郵撥"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費時間</strong></td>
			<td align="left"><%
			'繳費時間
			if trim(rsPay("PAYDATE"))<>"" and not isnull(rsPay("PAYDATE")) then
				response.write gArrDT(trim(rsPay("PAYDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費人姓名</strong></td>
			<td align="left"><%
			'繳費人姓名
			if trim(rsPay("PAYER"))<>"" and not isnull(rsPay("PAYER")) then
				response.write trim(rsPay("PAYER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>應繳金額</strong></td>
			<td align="left"><%
			'應繳金額
			if trim(rsPay("FORFEIT"))<>"" and not isnull(rsPay("FORFEIT")) then
				response.write trim(rsPay("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費金額</strong></td>
			<td align="left"><%
			'繳費金額
			if trim(rsPay("PAYAMOUNT"))<>"" and not isnull(rsPay("PAYAMOUNT")) then
				response.write trim(rsPay("PAYAMOUNT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>是否繳清結案</strong></td>
			<td align="left"><%
			'是否繳清結案
			if trim(rsPay("CASECLOSE"))<>"" and not isnull(rsPay("CASECLOSE")) then
				if trim(rsPay("CASECLOSE"))="1" then
					response.write "結案"
				else
					response.write "未結案"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsPay("RECORDSTATEID"))<>"" and not isnull(rsPay("RECORDSTATEID")) then
				if trim(rsPay("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsPay("RECORDDATE"))<>"" and not isnull(rsPay("RECORDDATE")) then
				response.write gArrDT(trim(rsPay("RECORDDATE")))&" "
				response.write Right("00"&hour(rsPay("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsPay("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsPay("RECORDMEMBERID"))<>"" and not isnull(rsPay("RECORDMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsPay("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsPay("DELMEMBERID"))<>"" and not isnull(rsPay("DELMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsPay("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>備註</strong></td>
			<td align="left"><%
			'備註
			if trim(rsPay("NOTE"))<>"" and not isnull(rsPay("NOTE")) then
				response.write trim(rsPay("NOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>是否遲繳</strong></td>
			<td align="left"><%
			'是否遲繳
			if trim(rsPay("ISLATE"))<>"" and not isnull(rsPay("ISLATE")) then
				if trim(rsPay("ISLATE"))="0" then
					response.write "如期繳納"
				else
					response.write "逾期繳納"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td></td>
		</tr>
<%	rsPay.MoveNext
	Wend
	rsPay.close
	set rsPay=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>7"></a>
				<strong>行人攤販繳費記錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">舉行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>
			</td>
		</tr>
<%	i=0
	strPasserArrived="select * from PasserSendArrived where PasserSn="&trim(rs1("SN"))&" order by SN"
	set rsArrived=conn.execute(strPasserArrived)
	If Not rsArrived.Bof Then rsArrived.MoveFirst 
	While Not rsArrived.Eof
		if i=0 then
			i=i+1
			TRcolor="#FFFF99"
		else
			i=i-1
			TRcolor="#AAF2A2"
		end if
%>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達流水號</strong></td>
			<td align="left"><%
			'送達流水號
			if trim(rsArrived("SN"))<>"" and not isnull(rsArrived("SN")) then
				response.write trim(rsArrived("SN"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>登錄人</strong></td>
			<td align="left"><%
			'登錄人
			if trim(rsArrived("RecordMemberID"))<>"" and not isnull(rsArrived("RecordMemberID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsArrived("RecordMemberID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達日期</strong></td>
			<td align="left"><%
			'送達日期
			if trim(rsArrived("ArrivedDate"))<>"" and not isnull(rsArrived("ArrivedDate")) then
				response.write gArrDT(trim(rsArrived("ArrivedDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達人員</strong></td>
			<td align="left"><%
			'送達人員
			if trim(rsArrived("SenderMemID"))<>"" and not isnull(rsArrived("SenderMemID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsArrived("SenderMemID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達圖片</strong></td>
			<td colspan="3"><Span <%=ShowMap(trim(rsArrived("SN")))%>>送達圖片</span></td>
		</tr>
		<tr>
			<td></td>
		</tr>
<%	rsArrived.MoveNext
	Wend
	rsArrived.close
	set rsArrived=nothing
%>
	</table>
	<br>
<%	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
end if
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
<%if CheckSelectData1="0" and CheckSelectData2="0" then%>
	alert("查無資料！");
	window.close();
<%end if%>
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgFileName){
	urlstr='../ProsecutionImage/ProsecutionImageDetail.asp?FileName='+ImgFileName.replace(/\+/g,'@2@')+'&SN=1';
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
function DP(){
	window.focus();
	window.print();
}
</script>
</html>
