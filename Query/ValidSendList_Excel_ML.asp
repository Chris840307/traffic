<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 6800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

%>
<%
	StationArrayTemp=""
	strwhere=request("SQLstr")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
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
<title>有效清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<%
strSQL="select UnitName,UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
	if sys_City="屏東縣" then
		Sys_UnitName=replace(rsunit("UnitName"),"屏東縣政府警察局","")
	elseif sys_City="高雄市" then
		strT2="select UnitName from UnitInfo where UnitID='"&trim(rsunit("UnitTypeID"))&"'"
		set rsT2=conn.execute(strT2)
		if not rsT2.eof then
			Sys_UnitName=trim(rsT2("UnitName"))
		end if
		rsT2.close
		set rsT2=nothing
	else
		Sys_UnitName=trim(rsunit("UnitName"))
	end if
rsunit.close
%>
<form name=myForm method="post">
<%	
ExchangeTypeFlag="W"
stopBatchnumber=""
strExchangeType="select a.ExchangeTypeID,f.BillUnitID,a.Batchnumber from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 and rownum<=1 "&strwhere
set rsEType=conn.execute(strExchangeType)
if not rsEType.eof then
	if trim(rsEType("ExchangeTypeID"))="N" then
		ExchangeTypeFlag="N"
	else
		ExchangeTypeFlag="W"
	end if
	BillUnitIDtmp=trim(rsEType("BillUnitID"))
	stopBatchnumber=trim(rsEType("Batchnumber"))
else
	ExchangeTypeFlag="W"
	BillUnitIDtmp=""
end if
rsEType.close
set rsEType=Nothing

	EofFlag1=1
	EofFlag2=1
	PrintSN=0
	'每頁共幾筆
	PageTotalCnt=18	'逕舉
	PageTotalCnt2=12 '攔停

'逕舉===============================================================================
If Trim(request("BillType"))="0" Or Trim(request("BillType"))="2" then
	ReportUnitArrayTemp=""
	strU="select distinct(g.UnitTypeID) from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f ,UnitInfo g" &_
		" where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8)) and a.ExchangeTypeID='W') or (d.DCIreturnStatus=1 and a.ExchangeTypeID<>'W'))" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere &_
		" and f.BillUnitID=g.UnitID order by g.UnitTypeID"
	set rsU=conn.execute(strU)
	While Not rsU.Eof
		if ReportUnitArrayTemp="" then
			ReportUnitArrayTemp=trim(rsU("UnitTypeID"))
		else
			ReportUnitArrayTemp=ReportUnitArrayTemp&","&trim(rsU("UnitTypeID"))
		end if
	rsU.MoveNext
	Wend
	rsU.close
	set rsU=nothing
'response.write ReportUnitArrayTemp
ReportUnitArray=Split(ReportUnitArrayTemp,",")
For R=0 To ubound(ReportUnitArray)
	PrintSN=0
	if R>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	strSQL="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2' and f.RecordStateID=0" &_
		" and f.BIllUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(ReportUnitArray(R))&"')" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8)) and a.ExchangeTypeID='W') or (d.DCIreturnStatus=1 and a.ExchangeTypeID<>'W'))" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	if trim(rs1("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rs1("cnt"))/PageTotalCnt+0.9999999)
	end if
	rs1.close

	strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2," &_
		"f.Rule3,f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation" &_
		",a.ExchangeTypeID,a.DciReturnStatusID,a.FileName from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID='2'" &_
		" and f.RecordStateID=0" &_
		" and f.BIllUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(ReportUnitArray(R))&"')" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8)) and a.ExchangeTypeID='W') or (d.DCIreturnStatus=1 and a.ExchangeTypeID<>'W'))" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	If not rs1.Eof then 
		EofFlag1=0
	end if
	While Not rs1.Eof
		if (fix(PrintSN/PageTotalCnt)+1)>1 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<table width="710" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%
			strUN="select UnitName from UnitInfo where UnitID='"&Trim(ReportUnitArray(R))&"'"
			Set rsUN=conn.execute(strUN)
			If Not rsUN.eof Then
				response.write Trim(rsUN("UnitName"))
			End If
			rsUN.close
			Set rsUN=Nothing 
				%>有效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：逕舉 &nbsp; &nbsp; &nbsp; &nbsp; <%
			If sys_City="基隆市" Then
				response.write "作業批號：" & stopBatchnumber
			End if
			%></td>
			<td align="right">Page <%=fix(PrintSN/PageTotalCnt)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="center">編號</td>
			<td width="9%" align="center">單號<br>DCI檔名</td>
			<td width="9%" align="center">違規日期<br>違規時間</td>
			<td width="9%" align="center"><br>車號</td>
			<td width="8%" align="center">法條1</td>
			<td width="31%" align="center"><br>駕駛人 / 車主</td>
			<td width="10%" align="center">員警<br>舉發單位</td>
			<td width="11%" align="center">到案處所<br>代保管物</td>
			<td width="9%" align="center">投郵日期<br>貼條號碼</td>
		</tr>
<%		
		for i=1 to PageTotalCnt
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>		<tr>
			<td><%
			'序號編號
			response.write PrintSN
			%></td>
			<td><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'檔名
			if trim(rs1("FileName"))<>"" and not isnull(rs1("FileName")) then
				response.write "<font size=1>"&trim(rs1("FileName"))&"</font>"
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'違規日期違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write year(rs1("IllegalDate"))-1911&Right("00"&month(rs1("IllegalDate")),2)&Right("00"&day(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'車號,簡示車種
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車" 
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			%></td>
			<td><%
			'法條
			RuleStr=""
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))&"<br>"
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule2"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule2"))
				end if
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule3"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule3"))
				end if
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule4"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule4"))
				end if
			end if
			if RuleStr="" then
				response.write "&nbsp;"
			else
				response.write RuleStr
			end if
			%></td>
			<td><%
			'抓取BillBaseDCIReturn的資料
			DciOwner=""
			DciOwnerAddress=""
			DciDriverHomeAddress=""
			DCIStation=""
			if trim(rs1("BillNO"))="" or isnull(rs1("BillNO")) then
				if trim(rs1("ExchangeTypeID"))="N" then
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='W'" &_
						" and Status='Y'"
				else
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
						" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
				end if
			else
				if trim(rs1("ExchangeTypeID"))="N" then
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO='"&trim(rs1("BillNO"))&"'" &_
						" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='W'" &_
						" and Status='Y'"
				else
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO='"&trim(rs1("BillNO"))&"'" &_
						" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
						" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
				end if
			end if
			set rsBDci=conn.execute(strBillDci)
			if not rsBDci.eof then
				DciOwner=trim(rsBDci("Owner"))
				DciOwnerZip=trim(rsBDci("OwnerZip"))
				DciOwnerAddress=trim(rsBDci("OwnerAddress"))
				DciDriverHomeZip=trim(rsBDci("DriverHomeZip"))
				DciDriverHomeAddress=trim(rsBDci("DriverHomeAddress"))
				DCIStation=trim(rsBDci("DCIreturnStation"))
				DicExchangeType=trim(rsBDci("ExchangeTypeID"))
			end if
			rsBDci.close
			set rsBDci=nothing
			'車主
			if trim(rs1("BillTypeID"))="2" then
				response.write funcCheckFont(DciOwner,18,1)
			else
				response.write funcCheckFont(trim(rs1("Driver")),18,1)
			end if
			GetMailAddress=""
			if trim(rs1("BillTypeID"))="2" then
				if DciOwnerAddress<>"" and not isnull(DciOwnerAddress) then
					if DicExchangeType="A" then
						GetMailAddress=DciOwnerZip&DciOwnerAddress
					else
						if DciOwnerZip<>"" and not isnull(DciOwnerZip) then
							strZip="select * from Zip where ZipID='"&DciOwnerZip&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						end if
						GetMailAddress=ZipName&DciOwnerAddress
					end if 
				end if
			else
				if DciDriverHomeAddress<>"" and not isnull(DciDriverHomeAddress) then
					if DicExchangeType="A" then
						GetMailAddress=DciDriverHomeAddress
					else
						if DciDriverHomeZip<>"" and not isnull(DciDriverHomeZip) then
							strZip="select * from Zip where ZipID='"&DciDriverHomeZip&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						end if
						GetMailAddress=ZipName&DciDriverHomeAddress
					end if 
				end if
			end if
			response.write "<br>"
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,18,1)&"</font>"
			%></td>
			<td><%
			'員警
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUN=conn.execute(strUName)
				if not rsUN.eof then
					response.write "<font size=1>"&trim(rsUN("UnitName"))&"</font>"
				end if
				rsUN.close
				set rsUN=nothing
			end if
			%></td>
			<td><%
			'到案處所
			'攔停用BillBase=MemberStation 逕舉用BillBaseDCIReturn=DCIreturnStation
			if trim(rs1("BillTypeID"))="2" then 
				stationID=DCIStation
			else
				stationID=rs1("MemberStation")
			end if
			if trim(stationID)<>"" and not isnull(stationID) then
				strMemberStation="select DCIStationName from Station where DCIstationID='"&trim(stationID)&"'"
				set rsMS=conn.execute(strMemberStation)
				if not rsMS.eof then
					response.write "<font size=1>"&trim(rsMS("DCIStationName"))&"</font>"
				end if
				rsMS.close
				set rsMS=nothing
			end if
			response.write "<br>"
			'代保管物
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "0"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td><%
			'投郵日期貼條號碼
			strMailDate="select MailDate,MailNumber from BillMailHistory where BillSN="&trim(rs1("SN"))
			set rsMD=conn.execute(strMailDate)
			if not rsMD.eof then
				if trim(rsMD("MailDate"))<>"" then
					response.write gInitDT(rsMD("MailDate"))
					response.write "<br>"
					response.write rsMD("MailNumber")
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			rsMD.close
			set rsMD=nothing
			%></td>

		</tr>
<%
		rs1.MoveNext
		next
%>	
	</table>	
<%
	
	Wend
	rs1.close
	set rs1=nothing
if trim(PrintSN)<>"0" then
%>
共計： <%=PrintSN%>  &nbsp;筆<br>
前揭違規案件業已建檔傳送資料庫，請查核無誤後，於603表內蓋章，本件留存備查。
<br>
<%
	end If
Next 
End If 
'攔停=========================================================================================
If Trim(request("BillType"))="0" Or Trim(request("BillType"))="1" Then
	If EofFlag1=0 Then response.write "<div class=""PageNext"">&nbsp;</div>"
	StopUnitArrayTemp=""
	strU="select distinct(g.UnitTypeID) from DCILog a,MemberData b" &_
		",BillBaseDciReturn e,DCIReturnStatus d" &_
		",BillBase f,UnitInfo g where a.BillSN=f.SN and a.BillTypeID<>'2' and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere &_
		" and f.BillUnitID=g.UnitID order by g.UnitTypeID"
	set rsU=conn.execute(strU)
	While Not rsU.Eof
		if StopUnitArrayTemp="" then
			StopUnitArrayTemp=trim(rsU("UnitTypeID"))
		else
			StopUnitArrayTemp=StopUnitArrayTemp&","&trim(rsU("UnitTypeID"))
		end if
	rsU.MoveNext
	Wend
	rsU.close
	set rsU=Nothing
	'response.write StopUnitArrayTemp
	StopUnitArray=Split(StopUnitArrayTemp,",")
For S=0 To UBound(StopUnitArray)
	if S>0 then response.write "<div class=""PageNext"">&nbsp;</div>"

	strSQL="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and a.BillTypeID<>'2'" &_
		" and f.RecordStateID=0" &_
		" and f.BIllUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(StopUnitArray(S))&"')" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	if trim(rs1("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rs1("cnt"))/PageTotalCnt2+0.9999999)
	end if
	rs1.close
	PrintSN=0
	strSQL="select f.SN,a.BillNO,f.IllegalDate,f.CarNo,f.CarSimpleID,f.Rule1,f.Rule2," &_
		"f.Rule3,f.Rule4,f.BillTypeID,f.Driver,f.BillMem1,a.BillUnitID,f.MemberStation" &_
		",f.Driver,f.DriverZip,f.DriverAddress,f.Owner,f.OwnerZip,f.OwnerAddress" &_
		",a.ExchangeTypeID,a.DciReturnStatusID,a.FileName from DCILog a,MemberData b" &_
		",BillBaseDciReturn e,DCIReturnStatus d" &_
		",BillBase f where a.BillSN=f.SN and a.BillTypeID<>'2' and f.RecordStateID=0" &_
		" and f.BIllUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(StopUnitArray(S))&"')" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+)"&strwhere&" order by f.RecordMemberID,f.RecordDate"
	set rs1=conn.execute(strSQL)
	If not rs1.Eof then 
		EofFlag2=0
	end if
	While Not rs1.Eof
		if (fix(PrintSN/PageTotalCnt2)+1)>1 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<table width="710" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%
			strUN="select UnitName from UnitInfo where UnitID='"&Trim(StopUnitArray(S))&"'"
			Set rsUN=conn.execute(strUN)
			If Not rsUN.eof Then
				response.write Trim(rsUN("UnitName"))
			End If
			rsUN.close
			Set rsUN=Nothing 
				%>有效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：攔停 &nbsp; &nbsp; &nbsp; &nbsp; <%
			If sys_City="基隆市" Then
				response.write "作業批號：" & stopBatchnumber
			End if
			%></td>
			<td align="right">Page <%=fix(PrintSN/PageTotalCnt2)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="center">編號</td>
			<td width="9%" align="center">單號<br>DCI檔名</td>
			<td width="9%" align="center">違規日期<br>違規時間</td>
			<td width="9%" align="center"><br>車號</td>
			<td width="8%" align="center">法條1<br>法條2</td>
			<td width="31%" align="center">駕駛人<br> 車主</td>
			<td width="10%" align="center">員警<br>舉發單位</td>
			<td width="11%" align="center">到案處所<br>代保管物</td>
			<td width="9%" align="center">投郵日期<br>貼條號碼</td>
		</tr>
<%		
		for i=1 to PageTotalCnt2
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>		<tr>
			<td><%
			'序號編號
			response.write PrintSN
			%></td>
			<td><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'檔名
			if trim(rs1("FileName"))<>"" and not isnull(rs1("FileName")) then
				response.write "<font size=1>"&trim(rs1("FileName"))&"</font>"
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'違規日期違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write year(rs1("IllegalDate"))-1911&Right("00"&month(rs1("IllegalDate")),2)&Right("00"&day(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'車號,簡示車種
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車" 
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				end if
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			%></td>
			<td><%
			'法條
			RuleStr=""
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))&"<br>"
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule2"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule2"))
				end if
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule3"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule3"))
				end if
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if RuleStr="" then
					RuleStr=trim(rs1("Rule4"))
				else
					RuleStr=RuleStr&"<br>"&trim(rs1("Rule4"))
				end if
			end if
			if RuleStr="" then
				response.write "&nbsp;"
			else
				response.write RuleStr
			end if
			%></td>
			<td><%
			'抓取BillBaseDCIReturn的資料
			DciOwner=""
			DciOwnerAddress=""
			DciDriverHomeAddress=""
			DCIStation=""
			if trim(rs1("BillNO"))="" or isnull(rs1("BillNO")) then
				if trim(rs1("ExchangeTypeID"))="N" then
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='W'" &_
						" and Status='Y"
				else
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO is null and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
						" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
				end if
			else
				if trim(rs1("ExchangeTypeID"))="N" then
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO='"&trim(rs1("BillNO"))&"'" &_
						" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='W'" &_
						" and Status='Y'"
				else
					strBillDci="select * from BillBaseDCIReturn" &_
						" where BillNO='"&trim(rs1("BillNO"))&"'" &_
						" and CarNo='"&trim(rs1("CarNo"))&"' and" &_
						" ExchangeTypeID='"&trim(rs1("ExchangeTypeID"))&"'" &_
						" and Status='"&trim(rs1("DciReturnStatusID"))&"'"
				end if
			end if
			set rsBDci=conn.execute(strBillDci)
			if not rsBDci.eof then
				DciDriver=trim(rsBDci("Driver"))
				DciOwner=trim(rsBDci("Owner"))
				DciOwnerZip=trim(rsBDci("OwnerZip"))
				DciOwnerAddress=trim(rsBDci("OwnerAddress"))
				DciDriverHomeZip=trim(rsBDci("DriverHomeZip"))
				DciDriverHomeAddress=trim(rsBDci("DriverHomeAddress"))
				DCIStation=trim(rsBDci("DCIreturnStation"))
				DicExchangeType=trim(rsBDci("ExchangeTypeID"))
			end if
			rsBDci.close
			set rsBDci=nothing
			'駕駛人
			ZipName=""
			GetMailAddress=""
			response.write funcCheckFont(trim(rs1("Driver")),18,1)
			GetMailAddress=""
			if trim(rs1("DriverAddress"))<>"" then
				if trim(rs1("DriverZip"))<>"" then
					strZip="select * from Zip where ZipID='"&trim(rs1("DriverZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing
				end if
				GetMailAddress=trim(rs1("DriverZip"))&ZipName&replace(replace(trim(rs1("DriverAddress"))&"","臺","台"),ZipName,"")
			end if
			response.write "<br>"
			response.write "<font size=1>"&funcCheckFont(GetMailAddress,18,1)&"</font>"
			'車主
			ZipName2=""
			GetMailAddress2=""
			response.write "<br>"&funcCheckFont(trim(rs1("Owner")),18,1)
			GetMailAddress2=""
			if trim(rs1("OwnerAddress"))<>"" then
				if trim(rs1("OwnerZip"))<>"" then
					strZip="select * from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing
				end if
				GetMailAddress2=trim(rs1("OwnerZip"))&ZipName2&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName2,"")
			end if
			response.write "<br>"
			response.write "<font size=1>"&funcCheckFont(GetMailAddress2,18,1)&"</font>"
			%></td>
			<td><%
			'員警
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			response.write "<br>"
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUN=conn.execute(strUName)
				if not rsUN.eof then
					response.write "<font size=1>"&trim(rsUN("UnitName"))&"</font>"
				end if
				rsUN.close
				set rsUN=nothing
			end if
			%></td>
			<td><%
			'到案處所
			'攔停用BillBase=MemberStation 逕舉用BillBaseDCIReturn=DCIreturnStation
			if trim(rs1("BillTypeID"))="2" then 
				stationID=DCIStation
			else
				stationID=rs1("MemberStation")
			end if
			if trim(stationID)<>"" and not isnull(stationID) then
				strMemberStation="select DCIStationName from Station where DCIstationID='"&trim(stationID)&"'"
				set rsMS=conn.execute(strMemberStation)
				if not rsMS.eof then
					response.write "<font size=1>"&trim(rsMS("DCIStationName"))&"</font>"
				end if
				rsMS.close
				set rsMS=nothing
			end if
			response.write "<br>"
			'代保管物
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "0"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td><%
			'投郵日期貼條號碼
			strMailDate="select MailDate,MailNumber from BillMailHistory where BillSN="&trim(rs1("SN"))
			set rsMD=conn.execute(strMailDate)
			if not rsMD.eof then
				if trim(rsMD("MailDate"))<>"" then
					response.write gInitDT(rsMD("MailDate"))
					response.write "<br>"
					response.write rsMD("MailNumber")
				else
					response.write "&nbsp;"
				end if
			else
				response.write "&nbsp;"
			end if
			rsMD.close
			set rsMD=nothing
			%></td>

		</tr>
<%
		rs1.MoveNext
		next
%>	
	</table>
	
<%
	wend
	rs1.close
	set rs1=nothing
	if trim(PrintSN)<>"0" then
%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
	前揭違規案件業已建檔傳送資料庫，請查核無誤後，於603表內蓋章，本件留存備查。
<br>
<%	end If
Next 
End if
%>



<%
if EofFlag1=1 and EofFlag2=1 then
%>
	<table width="710" border="0" cellpadding="2" cellspacing="0">
		<tr>
			<td align="center" colspan="2">
				<font size="3"><%=Sys_UnitName%>有效清冊</font>
			</td>
		</tr>
		<tr>
			<td align="left">告發單別：</td>
			<td align="right">Page <%=fix(PrintSN/PageTotalCnt2)+1%> of <%=pagecnt%></td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="2" cellspacing="0">
		<tr>
			<td width="4%" height="28" align="center">編號</td>
			<td width="9%" align="center">單號<br>DCI檔名</td>
			<td width="9%" align="center">違規日期<br>違規時間</td>
			<td width="9%" align="center"><br>車號</td>
			<td width="8%" align="center">法條1<br>法條2</td>
			<td width="31%" align="center"><br>駕駛人 / 車主</td>
			<td width="10%" align="center">員警<br>舉發單位</td>
			<td width="11%" align="center">到案處所<br>代保管物</td>
			<td width="9%" align="center">投郵日期<br>貼條號碼</td>
		</tr>
	</table>
<%
end if
%>

</form>
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