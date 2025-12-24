<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/banner.asp"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>各式清冊/舉發單列印</title>
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
				Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
		next
		strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
	end if

	if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
	elseif trim(request("Sys_BillNo1"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo1"))&"' and '"&trim(request("Sys_BillNo1"))&"'"
	elseif trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.BillNo between '"&trim(request("Sys_BillNo2"))&"' and '"&trim(request("Sys_BillNo2"))&"'"
	end if

	if strwhere<>"" then
		strwhereToPrintCarData=strwhere
	else
		strwhereToPrintCarData=""
	end if

	if request("RecordDate")<>"" and request("RecordDate1")<>""then
		RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
		RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"
		if strwhere<>"" then
			strwhere=strwhere&" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and f.RecordMemberID="&Session("User_ID")
		else
			strwhere=" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and f.RecordMemberID="&Session("User_ID")
		end if
	end if
end if
DB_Display=request("DB_Display")
if DB_Display="show" then
	if trim(strwhere)<>"" then
		'strwhereToPrintCarData=strwhere
		if sys_City="基隆市" then
			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','5','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','5','9','a','j','A','F','H','K','L','T','V') and a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','5','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if

		'if trim(request("PBillSN"))="" then '與dci上下查詢不同
		strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL&" order by f.RecordDate"
		
		set rssn=conn.execute(strSQL)
		BillSN=""
		while Not rssn.eof
			if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
			BillSN=BillSN&trim(rssn("BillSN"))
			rssn.movenext
		wend
		rssn.close
		'end if

		if sys_City="基隆市" then
			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','5','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','5','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607')) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','5','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','5','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if
		
		strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','5','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','5','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"

		set chksuess=conn.execute(strSQL)
		filsuess=Cint(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and d.DCIRETURNSTATUS='-1' "&strwhere
		set chksuess=conn.execute(strSQL)
		fildel=Cint(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Cint(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN(+) and a.BillNo=f.BillNo(+) and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=Cint(Dbrs("cnt"))
		Dbrs.close
		
		if sys_City="基隆市" then
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.DciErrorCarData in ('1','3','5','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.DciErrorCarData in ('1','3','5','9','a','j','A','H','K','L','T','V') and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt=Cint(Dbrs("cnt"))
		Dbrs.close
		filsuess=filsuess-errCatCnt
		tmpSQL=strwhere
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><span class="style3">各式清冊/舉發單列印</span></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號 
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" size="35">
						
						(多個批號同時處理，各批號請用,隔開。如：95A361,95A382,95A486）						
						<br>
						舉發單號
						<input name="Sys_BillNo1" type="text" class="btn1" value="<%=UCase(request("Sys_BillNo1"))%>" size="14" maxlength="9">
						~
						<input name="Sys_BillNo2" type="text" class="btn1" value="<%=UCase(request("Sys_BillNo2"))%>" size="13" maxlength="9"> ( 列印 單筆 或 特定範圍 舉發單才需填寫)
						<br>
						<%if sys_City<>"嘉義縣" then%>
							建檔日期
							<input name="RecordDate" type="text" value="<%=request("RecordDate")%>" size="8" maxlength="6" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
							~
							<input name="RecordDate1" type="text" value="<%=request("RecordDate1")%>" size="8" maxlength="6" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
							(提供列印入案移送清冊/大宗清冊/大宗掛號單/郵費單使用)
						<%else%>
							<img src="space.gif" width="60" height="1">舉發單以及各式清冊需要原案件建檔人才可列印
							<br>
							
							<input name="RecordDate" type="Hidden" value="" size="8" class="btn1">
							<input name="RecordDate1" type="Hidden" value="" size="8" class="btn1">
						<%end if%>
						<br>
						<%if sys_City<>"嘉義縣" then%>
							<input type="button" name="btnSelt" value="查詢" onclick="funSelt('BatchSelt');"<%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(233,1)=false then
								response.write " disabled"
							end if
							%>>
						<%else%>
							<input type="button" name="btnSelt" value="查詢" onclick="funChiayiSelt('BatchSelt');"<%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(233,1)=false then
								response.write " disabled"
							end if
							%>>
						<%end if%>
						<input type="button" name="cancel" value="清除" onClick="location='AllBillPrint.asp'">

						<img src="space.gif" width="10" height="1"></img><strong>( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功 , <%=errCatCnt%> 筆無效  ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-filsuess-fildel-deldata-errCatCnt%>筆未處理. )</strong>
						<img src="space.gif" width="6" height="1"><a href="DciCarErrorData.asp" target="_blank">查看逕舉無效原因</a>
						
						<br>
						<img src="space.gif" width="60" height="1"></img><font size="2" >列印舉發單/各式清冊前，請先輸入 批號 或是 舉發單號 進行 查詢</font>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr>

		<td height="35" bgcolor="#FFDD77" align="left">
			<%if sys_City="基隆市" then %>
				<img src="space.gif" width="8" height="1">
				<font size="2">列印郵簡式舉發單，印表機紙張格式請選擇 Legal 8.5 X 14 </font>
				<br>
			<%end if%>
			<%if sys_City="花蓮縣" then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprintBill" value="列印 違規通知單( A4 郵簡式舉發單 )" onclick="funBillNoPrint(1)">
				<img src="space.gif" width="57" height="1">
			<%elseif sys_City="基隆市" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprintBill" value="列印 違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(0)">
				<img src="space.gif" width="57" height="1">
			<%end if%>
			<img src="space.gif" width="12" height="1"></img>
			<input type="button" name="Submit43635" value="整批入案 郵寄日期/大宗條碼資料註記" onclick="funBillMailInfoMark()">
			<br>
			<%if sys_City="澎湖縣"  then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(2);">
			<%elseif sys_City="彰化縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(7);">
			<%elseif sys_City="嘉義縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(10);">
			<%elseif sys_City="雲林縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(9);">
			<%end if%>

			<%if sys_City="南投縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillNonTouSendLegal();">
			<%elseif sys_City="台中縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillTaiChungSendLegal();">
			<%elseif sys_City="彰化縣" or sys_City="嘉義縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillCHCGLegal();">
			<%elseif sys_City="花蓮縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillHuaLienSendLegal();">
			<%elseif sys_City="宜蘭縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 (  B5  )  " onclick="funBillSendB5();">
			<%elseif sys_City="台中市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="台中市列印 送達證書 ( 橫式 ) " onclick="funBillTaiChungCitySendLegal();">
			<%end if%>
			<img src="space.gif" width="8" height="1"></img>
			<input type="button" name="btnprint" value="列印 送達證書 (  A4  )  " onclick="funBillSendLegal();">
			<!--<input type="button" name="Submit43635" value="列印違規相片（A4雙色舉發單）" onclick="funBillIimagePrint(3)">  -->
			<%if sys_City="彰化縣" then%>
				<br>
				<input type="button" name="btnprint" value="掛號郵件收回執 " onclick="funFastPostReceive();">
			<%end if%>
			<%if sys_City="花蓮縣" then %>		
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit43635" value="花蓮縣分局 列印違規通知單1（   點陣式 8 x 6 in    ）" onclick="funBillNoPrint(8)">			
			<%elseif sys_City="宜蘭縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprintBill" value="列印 違規通知單(   點陣式 8 x 6 in    )" onclick="funBillNoPrint(13)">
				<img src="space.gif" width="57" height="1">
			<%elseif sys_City="金門縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<!--<input type="button" name="Submit43635" value="金門縣 列印違規通知單（   點陣式 8 x 6 in    ）" onclick="funBillNoPrint(4)">-->
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(2);">
			<%elseif sys_City="連江縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit43635" value="連江縣 列印違規通知單（   點陣式 8 x 6 in    ）" onclick="funBillNoPrint(5)">
			<%elseif sys_City="南投縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit43635" value="列印違規通知單" onclick="funBillNoPrint(6)">
			<%elseif sys_City="台中縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit43635" value="列印違規通知單" onclick="funBillNoPrint(11)">
			<%elseif sys_City="台中市" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit43635" value="台中市 列印違規通知單" onclick="funBillNoPrint(12)">
			<%end if%>
			<br>
			<!--<img src="space.gif" width="8" height="1">
			<input type="button" name="btnprintBill" value="列印送達證書（Legal 8.5 X 14郵簡式）" onclick="funUrgeList()">-->
			<hr>			
			<!--<span class="style3">
			DCI檔案名稱
			<input name="textfield42324" type="text" value="" size="14" maxlength="13">
			</span>-->
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then '花蓮專用A3版%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList_HL()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊" onclick="funReportSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊" onclick="funReportSendList()">
		<%end if%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit42342" value="大宗掛號清冊" onclick="funMailList()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="退件清冊_寄存(未結案)" onclick="funReturnSendList_Store()">
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(全部)" onclick="funStoreSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(全部)" onclick="funStoreSendList()">
		<%end if%>
		<br>
			<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then '花蓮專用A3版%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList_HL()">
			<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList()">
			<%end if%>
			<%'if trim(request("Sys_ExchangeTypeID"))="W" then '入案%>
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4335" value="攔停移送清冊" onclick="funStopSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4335" value="攔停移送清冊" onclick="funStopSendList()">
		<%end if%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">
			<span class="style3"><img src="space.gif" width="80" height="8"></span>
			<input type="button" name="Submit488423" value="退件清冊_寄存(已結案)" onclick="funReturnSendList_Store_Close()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(已結案)" onclick="funStoreSendList_Close()">
		<br>
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="無效清冊" onclick="funUselessSendList_HL()">			
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="無效清冊" onclick="funUselessSendList()">	
		<%end if%>
		<%if sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊_A4" onclick="funReportSendList()" style="width: 135px; height: 27px;">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
		<%else%>
			<span class="style3"><img src="space.gif" width="117" height="8"></span>
			<span class="style3"><img src="space.gif" width="43" height="8"></span>
		<%end if%>

			
			<!-- <span class="style3"><img src="space.gif" width="163" height="8"></span> -->
			<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit488423" value="退件清冊_公示(未結案)" onclick="funReturnSendList_Gov()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(未結案)" onclick="funStoreSendList_UnClose()">
		<br>
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="結案清冊" onclick="funCaseCloseSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="結案清冊" onclick="funCaseCloseSendList()">
		<%end if%>
		<%if sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="攔停移送清冊_A4" onclick="funStopSendList()" style="width: 135px; height: 27px;">
			<span class="style3"><img src="space.gif" width="165" height="8"></span>
		<%else%>
			<span class="style3"><img src="space.gif" width="318" height="8"></span>
		<%end if%>
			<input type="button" name="Submit488423" value="退件清冊_公示(已結案)" onclick="funReturnSendList_Gov_Close()">
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(全部)" onclick="funGovSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(全部)" onclick="funGovSendList()">
		<%end if%>
		<br>
			<span class="style3"><img src="space.gif" width="427" height="8"></span>
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList_HL()">
		<%else%>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList()">
		<%end if%>
		<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
			<span class="style3"><img src="space.gif" width="38" height="8"></span>
			<input type="button" name="Submit4232" value="公告清冊" onclick="funOpenGovList()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
		<%else%>
			<span class="style3"><img src="space.gif" width="147" height="8"></span>
		<%end if%>
			<input type="button" name="Submit4232" value="公示送達清冊(已結案)" onclick="funGovSendList_Close()">
		<br>
			<span class="style3"><img src="space.gif" width="672" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(未結案)" onclick="funGovSendList_UnClose()">
		<br>
		<%if sys_City="嘉義縣" then%>
			<input type="button" name="Submit4233" value="退件清冊_寄存(未結案)_A4" onclick="funReturnSendList_Store_A4()">
			<input type="button" name="Submit488423" value="退件清冊_寄存(已結案)_A4" onclick="funReturnSendList_Store_Close_A4()">
			<input type="button" name="Submit488423" value="退件清冊_公示(未結案)_A4" onclick="funReturnSendList_Gov_A4()">
			<input type="button" name="Submit488423" value="退件清冊_公示(已結案)_A4" onclick="funReturnSendList_Gov_Close_A4()">
		<br>
		<%end if%>
		<HR>
		本批資料發文監理站日期
		<input name="Sys_SendOpenGovDocToStationDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendOpenGovDocToStationDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSendOpenGovDocToStationDate();">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		本批資料二次郵寄日期
		<input name="Sys_StoreAndSendMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendMailDate');">
		&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendMailDate();">
		<br><br>
	</td>
  </tr>
  <tr>
    <td><p align="center">&nbsp;</p>    </td></tr>
</table>
<input type="button" name="Submit4232" value="違規舉發單 / 清冊 列印設定說明" onclick="funPrintDetail()"> 各式清冊依據縣市需求分為 A4  或 A3 格式 </br> 
<br>
<font size="5"> 
	列印 <b>各式清冊 <br>
	超出頁面 <img src="space.gif" width="40" height="1"></b> 請確認 檔案 --> 列印格式--> 紙張設定<font size="3">  (請依據縣市需求選擇A4或A3)</font><br>
	<img src="space.gif" width="450" height="1">上下左右邊界請設定是否皆為 0mm 或是 5.08mm <br>

	<b>頁尾出現網址 </b> 請確認 檔案 --> 列印格式--> 頁首頁尾皆為空白 
	
</font>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="hd_PrintSum" value="0">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="PBillSN" value="<%=BillSN%>">
<input type="Hidden" name="printStyle" value="">
<input type="Hidden" name="Sys_MailDate" value="">
<input type="Hidden" name="Sys_JudeAgentSex" value="">
<input type="Hidden" name="Sys_Print" value="">
<input type="Hidden" name="Sys_strSQL" value="<%=tmpSQL%>">	
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var winopen;
function funSendOpenGovDocToStationDate(){
	if (myForm.Sys_SendOpenGovDocToStationDate.value!=''){
		//runServerScript("SendToStationDate.asp?SendOpenDate="+myForm.Sys_SendOpenGovDocToStationDate.value+"&BillSn="+myForm.PBillSN.value);
		var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
		xmlhttp.Open("post","SendToStationDate.asp",false);	
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");
		xmlhttp.send("SendOpenDate="+myForm.Sys_SendOpenGovDocToStationDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
		alert("儲存完成!!");
	}
}
function funStoreAndSendMailDate(){
	if (myForm.Sys_StoreAndSendMailDate.value!=''){
		//runServerScript("StoreAndSendMailDate.asp?StoreAndSendMailDate="+myForm.Sys_StoreAndSendMailDate.value+"&BillSn="+myForm.PBillSN.value);
		var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
		xmlhttp.Open("post","StoreAndSendMailDate.asp",false);	
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");
		xmlhttp.send("StoreAndSendMailDate="+myForm.Sys_StoreAndSendMailDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
		//alert(xmlhttp.responsetext);
		alert("儲存完成!!");
	}
}

function funChiayiSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==""&&myForm.Sys_BillNo1.value==""&&myForm.Sys_BillNo2.value==""&&myForm.RecordDate.value==""&&myForm.RecordDate1.value==""){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(error==0){
			runServerScript("chkAllBillPrint.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
		}
	}
}

function funSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==""&&myForm.Sys_BillNo1.value==""&&myForm.Sys_BillNo2.value==""&&myForm.RecordDate.value==""&&myForm.RecordDate1.value==""){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(error==0){
			myForm.hd_PrintSum.value="0";
			myForm.PBillSN.value="";
			myForm.DB_Move.value="";
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
function funDataDetail(SN){
	UrlStr="ViewBillBaseData_Car.asp?BillSn="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funUpdate(SN){
	UrlStr="../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funDel(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="Del";
	myForm.submit();
}
function funBillIimagePrint(StyleType){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=StyleType;
		funsubmit();
	}
}
function funBillNoPrint(StyleType){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=StyleType;
		runServerScript("BillNoPrint.asp?SQLstr=<%=tmpSQL%>&printStyle="+StyleType);
		funsubmit();
	}
}
function funFastPostReceive(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		UrlStr="FastPostReceive.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		UrlStr="PasserUrgeDeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillSendB5(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		//UrlStr="PasserUrgeDeliverList.asp";
		UrlStr="PasserUrgeHuaLien_DeliverListV.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillNonTouSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		UrlStr="PasserUrgeNanTou_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillTaiChungSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		UrlStr="PasserUrgeTaiChung_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="TaiChung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillCHCGLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		UrlStr="PasserUrgeCHCG_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="CHCG";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillHuaLienSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		UrlStr="PasserUrgeHuaLien_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funBillTaiChungCitySendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		UrlStr="PasserUrgeTaiChungCity_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funsubmit(){
	//if (!winopen.closed){winopen.close();}
	//var chkcnt=<%=Cint(filsuess)%>;
	if(myForm.printStyle.value=='0'){
		/*window.parent.frames("mainFrame").location="BillPrints.asp";
		myForm.action="BillPrints.asp";*/
		UrlStr="BillPrints.asp";
	}else if(myForm.printStyle.value=='2'){
		/*window.parent.frames("mainFrame").location="BillPrints_a4.asp";
		myForm.action="BillPrints_a4.asp";*/
		UrlStr="BillPrints_a4.asp";
	}else if(myForm.printStyle.value=='1'){
		UrlStr="BillPrints_legalA4.asp";
	}else if(myForm.printStyle.value=='3'){
		UrlStr="BillImagePrint.asp";
	}else if(myForm.printStyle.value=='4'){
		UrlStr="BillPrints_lattice.asp";
	}else if(myForm.printStyle.value=='5'){
		UrlStr="BillPrints_lattice_MU.asp";
	}else if(myForm.printStyle.value=='6'){
		UrlStr="BillPrints_lattice_NanTou.asp";
	}else if(myForm.printStyle.value=='7'){
		UrlStr="BillPrintsCHCG_a4.asp";
	}else if(myForm.printStyle.value=='8'){
		UrlStr="BillPrints_lattice_HuaLien2.asp";
	}else if(myForm.printStyle.value=='9'){
		UrlStr="BillPrintsYunLin_a4.asp";
	}else if(myForm.printStyle.value=='10'){
		UrlStr="BillPrintsChiayi_a4.asp";
	}else if(myForm.printStyle.value=='11'){
		UrlStr="BillPrints_lattice_TaiChung.asp";
	}else if(myForm.printStyle.value=='12'){
		UrlStr="BillPrints_lattice_City.asp";
	}else if(myForm.printStyle.value=='13'){
		UrlStr="BillPrints_lattice_YiLan.asp";
	}
	/*myForm.target="mainFrame";
	myForm.submit();
	myForm.action="";
	myForm.target="";*/
	/*myForm.btnprint.disabled=false;
	if(myForm.Sys_Print.value!=''){
		myForm.hd_PrintSum.value=parseInt(myForm.hd_PrintSum.value)+parseInt(myForm.Sys_Print.value);
		if(parseInt(myForm.hd_PrintSum.value)-parseInt(myForm.Sys_Print.value)>chkcnt){
			myForm.btnprint.disabled=true;
		}else{
			myForm.btnprint.disabled=false;
		}
	}
	setTimeout('',2000);
	newWin(UrlStr,"JudeBat",920,600,50,10,"yes","yes","yes","no");*/
	myForm.action=UrlStr;
	myForm.target="JudeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	/*if(myForm.printStyle.value!='4'){
		setTimeout('funchgprint()',4000);
	}*/
}
function funUrgeList(){
	UrlStr="JudeStyle.asp";
	newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
	myForm.action="JudeStyle.asp";
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funJudesubmit(){
	winopen.close();
	if(myForm.printStyle.value=='0'){
		UrlStr="BillPrints_legal.asp";		
		newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
		myForm.action=UrlStr;
		myForm.target="UrgeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		setTimeout('funchgprint()',2000);
	}else{
		UrlStr="PasserJudeA4.asp?PBillSN="+myForm.PBillSN.value;
		newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
	}
}
function funchgprint(){
	winopen.printWindow(true,5.08,5.08,5.08,5.08);
}
function funchgExecel(){
	UrlStr="DCIExchangeQry_Execel.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"inputWin",700,550,50,10,"yes","yes","yes","no");
}
function funPrintDetail(){
	UrlStr="PictureDetail.htm";
	newWin(UrlStr,"inputWin",1000,800,50,10,"yes","yes","yes","no");
}
function funPrintStyle(){

		UrlStr="SendStyle.asp";
		newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
		myForm.action="SendStyle.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
//大宗郵件
function funMailList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印大宗郵件清冊的舉發單！");
	}else{
		UrlStr="MailSendList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailSendList",300,125,200,100,"no","no","no","no");
	}
}
//大宗郵件2
function funMailList2(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印交寄大宗函件的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>&MailSendType=S";
		newWin(UrlStr,"MailReportList",300,220,350,200,"no","no","no","no");
	}
}
//郵費清單
function funMailMoneyList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印郵費單的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList",300,220,350,200,"no","no","no","no");
	}
}
//逕舉
function funReportSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印逕舉移送清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReportSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReportSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}
//逕舉_花蓮A3版
function funReportSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印逕舉移送清冊的舉發單！");
	}else{
		UrlStr="ReportSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停
function funStopSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="StopSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="StopSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停_花蓮A3
function funStopSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
		UrlStr="StopSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊
function funValidSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4",800,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊_花蓮A3版
function funValidSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊
function funUselessSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊_花蓮A3版
function funUselessSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//結案清冊
function funCaseCloseSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="CaseCloseSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"CaseCloseWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//結案清冊_花蓮A3版
function funCaseCloseSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="CaseCloseSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"CaseCloseWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(未結案)
function funReturnSendList_Store(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Store.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="ReturnSendList_Excel_A3_Store.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Store.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(已結案)
function funReturnSendList_Store_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="ReturnSendList_Excel_A3_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin41",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(未結案)
function funReturnSendList_Gov(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="ReturnSendList_Excel_A3_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(已結案)
function funReturnSendList_Gov_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="ReturnSendList_Excel_A3_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin65",800,700,0,0,"yes","yes","yes","no");
	}
}
//======================================
//退件清冊_寄存(未結案)
function funReturnSendList_Store_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Store.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(已結案)A4
function funReturnSendList_Store_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Store_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin41",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(未結案)A4
function funReturnSendList_Gov_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Gov.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(已結案)A4
function funReturnSendList_Gov_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin65",800,700,0,0,"yes","yes","yes","no");
	}
}
//================================================
//收受
function funGetSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="GetSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="GetSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//收受_花蓮A3版
function funGetSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="GetSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊
function funStoreSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin7",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊_花蓮A3版
function funStoreSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
		UrlStr="funStoreSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin7",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(未結案)
function funStoreSendList_UnClose(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="funStoreSendList_Excel_A3_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin71",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(已結案)
function funStoreSendList_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="funStoreSendList_Excel_A3_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin72",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊
function funGovSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊_花蓮
function funGovSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(已結案)
function funGovSendList_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="funGovSendList_Excel_A3_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin81",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(未結案)
function funGovSendList_UnClose(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="台中市" or sys_City="嘉義縣" then%>
		UrlStr="funGovSendList_Excel_A3_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin82",800,700,0,0,"yes","yes","yes","no");
	}
}
//車籍查詢
function funchgCarDataList(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印車籍清冊的舉發單！");
	}else{
		UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","no","yes","no");
	}
}
//車籍查詢_花蓮A3版
function funchgCarDataList_HL(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印車籍清冊的舉發單！");
	}else{
		UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","no","yes","no");
	}
}
//公告清冊
function funOpenGovList(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印公告清冊的舉發單！");
	}else{
		UrlStr="funOpenGovList_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
function funBillMailInfoMark(){
	UrlStr="BillMailInfoMark.asp";
	newWin(UrlStr,"inputWin",800,600,50,10,"yes","no","yes","no");
	myForm.action="BillMailInfoMark.asp";
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
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
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
</script>
<%conn.close%>